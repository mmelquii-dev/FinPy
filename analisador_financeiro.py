import os
import json
import re
import threading
import itertools
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog
from groq import Groq
from dotenv import load_dotenv

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    TEM_DND = True
except Exception:
    TEM_DND = False

try:
    import pdfplumber
    TEM_PDF = True
except Exception:
    TEM_PDF = False

# Carrega as variáveis de ambiente
load_dotenv()

class AnalisadorFinanceiro:
    def __init__(self, api_key=None):
        self.api_key = api_key or os.getenv("GROQ_API_KEY")
        if not self.api_key:
            raise ValueError(
                "Chave de API do Groq não encontrada. "
                "Crie um arquivo .env contendo 'GROQ_API_KEY=sua_chave' ou informe-a diretamente."
            )
        self.client = Groq(api_key=self.api_key)
        self.df = None
        self.modelo_ia = "qwen/qwen3.8-27b"

    def _identificar_colunas_com_ia(self, colunas, amostra):
        """
        Usa o Groq para mapear os nomes reais das colunas do extrato bancário
        para os campos padronizados: Data, Descricao, Valor e Categoria.
        """
        prompt = f"""
        Você é um especialista em análise de extratos bancários brasileiros.
        Analise a lista de colunas e as primeiras linhas de dados deste extrato:
        Colunas: {colunas}
        Amostra de dados: {amostra}

        Identifique qual coluna corresponde a cada item:
        - data_col: coluna de datas
        - descricao_col: coluna de descrição, histórico, lançamento ou estabelecimento
        - valor_col: coluna de valor financeiro (positivo ou negativo)
        - categoria_col: coluna de categoria (se existir no arquivo, senão null)

        Responda APENAS um JSON válido no formato:
        {{
          "data_col": "nome_exato_da_coluna",
          "descricao_col": "nome_exato_da_coluna",
          "valor_col": "nome_exato_da_coluna",
          "categoria_col": null
        }}
        """
        try:
            resposta = self.client.chat.completions.create(
                model=self.modelo_ia,
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"},
                max_tokens=250,
                temperature=0.0
            )
            return json.loads(resposta.choices[0].message.content)
        except Exception as e:
            print(f"Aviso: Fallback na identificação de colunas por IA: {e}")
            return None

    def _desduplicar_colunas(self, colunas):
        """Garante nomes de colunas estritamente únicos em qualquer DataFrame."""
        vistas = {}
        novas = []
        for c in colunas:
            c_str = str(c).strip() if c is not None and str(c).strip() else "Col"
            if c_str in vistas:
                vistas[c_str] += 1
                novas.append(f"{c_str}_{vistas[c_str]}")
            else:
                vistas[c_str] = 0
                novas.append(c_str)
        return novas

    def _categorizar_heuristica(self, desc):
        """Classificação instantânea baseada em padrões bancários frequentes."""
        d = str(desc).upper()
        if any(k in d for k in ['UBER', '99 TECNOLOGIA', 'KIM+', 'POSTO', 'COMBUSTIVEL', 'ESTACIONAMENTO', 'METRO', 'CCR']):
            return 'TRANSPORTE'
        if any(k in d for k in ['FARMACIA', 'DROGARIA', 'PAGUE MENOS', 'HOSPITAL', 'CLINICA', 'CONSULTA', 'MEDIC', 'SAUDE', 'LABORATORIO']):
            return 'SAÚDE'
        if any(k in d for k in ['IFOOD', 'DELIVERY', 'RAPPI', 'AIQFOME']):
            return 'DELIVERY'
        if any(k in d for k in ['RESTAURANTE', 'LANCHES', 'SALGADOS', 'BURGUER', 'DELICATESSEN', 'PIZZA', 'PADARIA', 'CAFE', 'BAR', 'DOGAO', 'CHURRASC']):
            return 'RESTAURANTE'
        if any(k in d for k in ['ATACADAO', 'MERCADO', 'SUPERMERCADO', 'ASSAI', 'CARREFOUR', 'EXTRA', 'BIG', 'HORTIFRUTI', 'ACOUGUE']):
            return 'MERCADO'
        if any(k in d for k in ['ALUGUEL', 'CONDOMINIO', 'IMOBIL', 'QUILTO']):
            return 'ALUGUEL'
        if any(k in d for k in ['COELB', 'EMBASA', 'ENEL', 'CPFL', 'SABESP', 'ENERGISA', 'LUZ', 'AGUA', 'GAS', 'TELECOM', 'CLARO', 'VIVO', 'TIM', 'OI']):
            return 'CONTAS DE CONSUMO'
        if any(k in d for k in ['NETFLIX', 'SPOTIFY', 'PRIME VIDEO', 'YOUTUBE', 'APPLE', 'GOOGLE', 'MAX', 'DISNEY']):
            return 'ASSINATURAS'
        if any(k in d for k in ['B3', 'CDB', 'LCI', 'LCA', 'TESOURO', 'INVEST', 'ATIVO', 'RENDIMENTOS']):
            return 'INVESTIMENTOS'
        if any(k in d for k in ['AMAZON', 'MERCADO LIVRE', 'SHOPEE', 'SHEIN', 'MAGALU', 'CASAS BAHIA', 'ALIEXPRESS']):
            return 'COMPRAS'
        return None

    def _identificar_nome_banco(self, caminho_arquivo, df):
        """Identifica o nome amigável do banco a partir do conteúdo ou nome do arquivo."""
        caminho_lower = os.path.basename(caminho_arquivo).lower()
        amostra_txt = ""
        if 'Descricao' in df.columns:
            amostra_txt += " " + " ".join(df['Descricao'].dropna().head(20).astype(str))
        if 'Histórico' in df.columns:
            amostra_txt += " " + " ".join(df['Histórico'].dropna().head(20).astype(str))
        amostra_txt = (caminho_lower + " " + amostra_txt).lower()

        if 'inter' in amostra_txt:
            return 'Banco Inter'
        if 'c6' in amostra_txt or '01m4' in amostra_txt:
            return 'C6 Bank'
        if 'nubank' in amostra_txt or 'nu pagamentos' in amostra_txt or 'nu financeira' in amostra_txt:
            return 'Nubank'
        if 'itau' in amostra_txt or 'itaú' in amostra_txt:
            return 'Itaú'
        if 'bradesco' in amostra_txt:
            return 'Bradesco'
        if 'santander' in amostra_txt:
            return 'Santander'
        if 'brasil' in amostra_txt or 'bb ' in amostra_txt:
            return 'Banco do Brasil'
        if 'caixa' in amostra_txt or 'cef' in amostra_txt:
            return 'Caixa Econômica'

        nome_base = os.path.splitext(os.path.basename(caminho_arquivo))[0]
        return re.sub(r'[-_]', ' ', nome_base).strip().title()

    def _categorizar_transacoes_com_ia(self, descricoes_unicas):
        """
        Usa o Groq para categorizar gastos em lotes compactos indexados,
        evitando estourar limites de taxa de tokens (OTPM).
        """
        itens = [str(d).strip() for d in descricoes_unicas if str(d).strip() and str(d).strip() != 'nan']
        resultado = {}
        tamanho_lote = 15

        # Processa em lotes de 15 descrições
        for i in range(0, min(len(itens), 45), tamanho_lote):
            lote = itens[i:i + tamanho_lote]
            itens_str = "\n".join([f"{idx}: {desc}" for idx, desc in enumerate(lote)])
            prompt = f"""
            Classifique cada descrição de despesa em uma destas categorias:
            [ALUGUEL, MERCADO, SAÚDE, TRANSPORTE, RESTAURANTE, DELIVERY, ASSINATURAS, COMPRAS, OUTROS]

            Itens:
            {itens_str}

            Retorne APENAS um JSON no formato:
            {{"0": "CATEGORIA", "1": "CATEGORIA"}}
            """
            try:
                resposta = self.client.chat.completions.create(
                    model=self.modelo_ia,
                    messages=[{"role": "user", "content": prompt}],
                    response_format={"type": "json_object"},
                    max_tokens=250,
                    temperature=0.1
                )
                mapeamento = json.loads(resposta.choices[0].message.content)
                for k, cat in mapeamento.items():
                    try:
                        idx = int(k)
                        resultado[lote[idx]] = str(cat).upper()
                    except (ValueError, IndexError):
                        continue
            except Exception as e:
                print(f"Aviso no lote de categorização: {e}")
                continue

        return resultado

    def _carregar_csv_com_cabecalho_dinamico(self, caminho_arquivo):
        """
        Detecta a linha de cabeçalho real e o separador correto, ignorando
        linhas de metadados que bancos (Inter, Itaú, BB, etc.) colocam no topo do arquivo.
        """
        encodings = ['utf-8', 'latin1', 'cp1252']
        linhas_amostra = []
        encoding_usado = 'utf-8'

        for enc in encodings:
            try:
                with open(caminho_arquivo, 'r', encoding=enc, errors='replace') as f:
                    linhas_amostra = [f.readline() for _ in range(40)]
                encoding_usado = enc
                break
            except Exception:
                continue

        header_idx = 0
        sep = ','

        # Procura a linha que contém os títulos das colunas da tabela
        for i, line in enumerate(linhas_amostra):
            line_lower = line.lower()
            if ('data' in line_lower or 'date' in line_lower) and (';' in line or ',' in line or '\t' in line):
                if any(k in line_lower for k in ['valor', 'hist', 'desc', 'saldo', 'lancamento', 'lançamento']):
                    header_idx = i
                    c_pv = line.count(';')
                    c_virg = line.count(',')
                    c_tab = line.count('\t')
                    if c_pv >= c_virg and c_pv >= c_tab:
                        sep = ';'
                    elif c_tab >= c_virg:
                        sep = '\t'
                    else:
                        sep = ','
                    break

        try:
            df = pd.read_csv(
                caminho_arquivo,
                skiprows=header_idx,
                sep=sep,
                encoding=encoding_usado,
                on_bad_lines='skip'
            )
        except Exception:
            # Fallback
            df = pd.read_csv(
                caminho_arquivo,
                skiprows=header_idx,
                sep=None,
                engine='python',
                encoding=encoding_usado,
                on_bad_lines='skip'
            )

        df.columns = self._desduplicar_colunas(df.columns)
        return df

    def _carregar_pdf(self, caminho_arquivo):
        """
        Extrai transações de extratos em PDF (C6, Nubank, Itaú, Inter, BB, etc.)
        utilizando regex linha a linha de alta precisão ou tabelas estruturadas com pdfplumber.
        """
        if not TEM_PDF:
            raise ImportError("A biblioteca 'pdfplumber' não está instalada. Execute: pip install pdfplumber")

        regex_val = re.compile(r"([-+]?\s*R\$\s*[\d\.]+,\d{2}|[-+]?\s*[\d\.]+,\d{2})\s*$")
        regex_date = re.compile(r"^(\d{2}[/-]\d{2}(?:[/-]\d{2,4})?)(?:\s+(\d{2}[/-]\d{2}(?:[/-]\d{2,4})?))?\s+(.*)$")

        registros = []
        ano_detectado = str(datetime.now().year)

        with pdfplumber.open(caminho_arquivo) as pdf:
            texto_completo = ""
            # 1. Tenta extração de alta precisão baseada em texto linha a linha
            for pagina in pdf.pages:
                texto_pag = pagina.extract_text() or ""
                texto_completo += "\n" + texto_pag

                for linha in texto_pag.split('\n'):
                    l = linha.strip()
                    if not l:
                        continue
                    # Ignora linhas de saldo e metadados de cabeçalho
                    if any(s in l.lower() for s in ['saldo do dia', 'saldo final', 'saldo anterior', 'extrato período', 'extrato exportado', 'data data', 'tipo descrição', 'lançamento contábil']):
                        continue

                    mv = regex_val.search(l)
                    if mv:
                        val_str = mv.group(1).strip()
                        resto = l[:mv.start()].strip()
                        md = regex_date.match(resto)
                        if md:
                            d1, d2, desc = md.group(1), md.group(2), md.group(3)
                            registros.append({
                                'Data': d1,
                                'Data_Lancamento': d2 if d2 else d1,
                                'Descricao': desc.strip(),
                                'Valor': val_str
                            })

            # Detecta o ano no texto do extrato (ex: 2026)
            match_ano = re.search(r'\b(202\d)\b', texto_completo)
            if match_ano:
                ano_detectado = match_ano.group(1)

        # Se encontrou transações pelo método de texto linha a linha
        if registros:
            df = pd.DataFrame(registros)
            # Normaliza ano nas datas caso venham como DD/MM
            df['Data'] = df['Data'].apply(
                lambda d: f"{d}/{ano_detectado}" if re.match(r'^\d{2}[/-]\d{2}$', str(d).strip()) else (
                    re.sub(r'[/-](\d{2})$', r'/20\1', str(d).strip()) if re.match(r'^\d{2}[/-]\d{2}[/-]\d{2}$', str(d).strip()) else str(d).strip()
                )
            )
            df.columns = self._desduplicar_colunas(df.columns)
            return df

        # 2. Fallback: Se não encontrou transações por texto, tenta extrair tabelas
        linhas_tabela = []
        with pdfplumber.open(caminho_arquivo) as pdf:
            for pagina in pdf.pages:
                tabelas = pagina.extract_tables()
                if tabelas:
                    for tab in tabelas:
                        for row in tab:
                            if row and any(cell and str(cell).strip() for cell in row):
                                row_limpa = [str(c or '').replace('\n', ' ').strip() for c in row]
                                if any('saldo' in str(c).lower() for c in row_limpa[:2]):
                                    continue
                                linhas_tabela.append(row_limpa)

        if linhas_tabela and len(linhas_tabela) > 1:
            header_idx = -1
            for i, row in enumerate(linhas_tabela[:10]):
                row_str = " ".join(row).lower()
                primeira_celula = row[0].strip()
                if not re.match(r'^\d{2}[/-]\d{2}', primeira_celula) and not re.match(r'^[-+]?R?\$', primeira_celula):
                    if any(k in row_str for k in ['data', 'date', 'dt']) and any(k in row_str for k in ['valor', 'saldo', 'hist', 'desc', 'lanc', 'r$']):
                        header_idx = i
                        break

            if header_idx >= 0:
                cabecalho = self._desduplicar_colunas([c if c else f"Col_{idx}" for idx, c in enumerate(linhas_tabela[header_idx])])
                dados_linhas = linhas_tabela[header_idx + 1:]
            else:
                num_cols = len(linhas_tabela[0])
                if num_cols == 5:
                    cabecalho = ['Data', 'Data_Lancamento', 'Tipo', 'Descricao', 'Valor']
                elif num_cols == 4:
                    cabecalho = ['Data', 'Tipo', 'Descricao', 'Valor']
                elif num_cols == 3:
                    cabecalho = ['Data', 'Descricao', 'Valor']
                else:
                    cabecalho = [f"Col_{idx}" for idx in range(num_cols)]
                dados_linhas = linhas_tabela

            num_cols = len(cabecalho)
            dados = []
            for r in dados_linhas:
                if len(r) == num_cols:
                    dados.append(r)
                elif len(r) > num_cols:
                    dados.append(r[:num_cols])
                elif len(r) > 1:
                    dados.append(r + [''] * (num_cols - len(r)))

            if dados:
                df = pd.DataFrame(dados, columns=cabecalho)
                df.columns = self._desduplicar_colunas(df.columns)
                return df

        raise ValueError(f"Não foi possível extrair dados textuais de '{os.path.basename(caminho_arquivo)}'. Se o PDF for uma imagem digitalizada, exporte em formato CSV pelo app do seu banco.")

    def normalizar_dados(self, caminho_arquivo):
        """
        Lê e normaliza os dados do extrato (CSV ou PDF) de qualquer banco brasileiro,
        usando a API do Groq para detectar colunas e enriquecer/categorizar dados.
        """
        nome_arquivo = os.path.basename(caminho_arquivo)
        print(f"Lendo e normalizando '{nome_arquivo}' com Groq...")

        # 1. Leitura do arquivo conforme extensão
        if caminho_arquivo.lower().endswith('.pdf'):
            self.df = self._carregar_pdf(caminho_arquivo)
        else:
            self.df = self._carregar_csv_com_cabecalho_dinamico(caminho_arquivo)

        # Remove colunas completamente vazias
        self.df = self.df.dropna(how='all', axis=1)
        self.df.columns = self._desduplicar_colunas(self.df.columns)

        # Identifica nome amigável do banco
        self.df['Banco'] = self._identificar_nome_banco(caminho_arquivo, self.df)

        # Se houver colunas Histórico e Descrição (padrão de bancos como Inter), enriquece
        if 'Histórico' in self.df.columns and 'Descrição' in self.df.columns:
            self.df['Descricao_Completa'] = (
                self.df['Histórico'].fillna('').astype(str).str.strip() + " - " +
                self.df['Descrição'].fillna('').astype(str).str.strip()
            ).str.strip(' -')

        colunas_originais = [c for c in self.df.columns.tolist() if not c.startswith('_')]
        amostra_dados = self.df[colunas_originais].head(2).to_dict(orient='records')

        # 2. IA identifica as colunas correspondentes (se já não estiverem padronizadas)
        mapa_colunas = None
        if not ({'Data', 'Valor'}.issubset(self.df.columns)):
            mapa_colunas = self._identificar_colunas_com_ia(colunas_originais, amostra_dados)

        renomear = {}
        if mapa_colunas:
            if mapa_colunas.get('data_col') in self.df.columns:
                renomear[mapa_colunas['data_col']] = 'Data'
            if mapa_colunas.get('descricao_col') in self.df.columns:
                renomear[mapa_colunas['descricao_col']] = 'Descricao'
            if mapa_colunas.get('valor_col') in self.df.columns:
                renomear[mapa_colunas['valor_col']] = 'Valor'
            if mapa_colunas.get('categoria_col') and mapa_colunas['categoria_col'] in self.df.columns:
                renomear[mapa_colunas['categoria_col']] = 'Categoria'

        # Fallback heurístico caso a IA não tenha mapeado todas as essenciais
        if 'Data' not in renomear.values() and 'Data' not in self.df.columns:
            for c in self.df.columns:
                if any(x in c.lower() for x in ['data', 'date', 'dt']):
                    renomear[c] = 'Data'
                    break

        if 'Valor' not in renomear.values() and 'Valor' not in self.df.columns:
            for c in self.df.columns:
                if any(x in c.lower() for x in ['valor', 'value', 'quantia', 'r$']):
                    renomear[c] = 'Valor'
                    break

        if 'Descricao' not in renomear.values() and 'Descricao' not in self.df.columns:
            for c in self.df.columns:
                if any(x in c.lower() for x in ['desc', 'hist', 'titul', 'estab', 'transa']):
                    renomear[c] = 'Descricao'
                    break

        self.df.rename(columns=renomear, inplace=True)
        self.df.columns = self._desduplicar_colunas(self.df.columns)

        if 'Data' not in self.df.columns or 'Valor' not in self.df.columns:
            raise ValueError(f"Não foi possível identificar colunas de Data e Valor em '{nome_arquivo}'. Colunas encontradas: {colunas_originais}")

        # Se houver múltiplas colunas chamadas Data ou Valor, garante que é uma Series
        if isinstance(self.df['Data'], pd.DataFrame):
            self.df['Data'] = self.df['Data'].iloc[:, 0]
        if isinstance(self.df['Valor'], pd.DataFrame):
            self.df['Valor'] = self.df['Valor'].iloc[:, 0]

        if 'Descricao_Completa' in self.df.columns:
            self.df['Descricao'] = self.df['Descricao_Completa']
        elif 'Descricao' not in self.df.columns:
            self.df['Descricao'] = "Transação"

        # 3. Limpeza e conversão do Valor (trata formatos R$, 1.250,50 ou -1250.50)
        if self.df['Valor'].dtype == object:
            self.df['Valor'] = (
                self.df['Valor']
                .astype(str)
                .str.replace('R$', '', regex=False)
                .str.strip()
                .str.replace(' ', '', regex=False)
            )
            tem_virgula = self.df['Valor'].str.contains(',', na=False).any()
            if tem_virgula:
                self.df['Valor'] = (
                    self.df['Valor']
                    .str.replace('.', '', regex=False)
                    .str.replace(',', '.', regex=False)
                )

        self.df['Valor'] = pd.to_numeric(self.df['Valor'], errors='coerce')

        # 4. Limpeza e conversão da Data
        ano_atual = str(datetime.now().year)
        data_str = self.df['Data'].astype(str).str.strip()
        data_str = data_str.apply(
            lambda d: f"{d}/{ano_atual}" if re.match(r'^\d{2}[/-]\d{2}$', d) else (
                re.sub(r'[/-](\d{2})$', r'/20\1', d) if re.match(r'^\d{2}[/-]\d{2}[/-]\d{2}$', d) else d
            )
        )
        self.df['Data'] = pd.to_datetime(data_str, dayfirst=True, format='mixed', errors='coerce')

        # Remove linhas sem Data ou Valor válidos
        self.df = self.df.dropna(subset=['Data', 'Valor']).copy()

        # 5. Normalização ou categorização inteligente (Heurística rápida + IA Groq)
        if 'Categoria' not in self.df.columns or self.df['Categoria'].dropna().empty:
            self.df['Categoria'] = self.df['Descricao'].apply(self._categorizar_heuristica)
            nao_categorizadas = self.df.loc[(self.df['Valor'] < 0) & (self.df['Categoria'].isna()), 'Descricao'].dropna().unique()
            if len(nao_categorizadas) > 0:
                mapa_ia = self._categorizar_transacoes_com_ia(nao_categorizadas)
                for desc, cat in mapa_ia.items():
                    self.df.loc[self.df['Descricao'] == desc, 'Categoria'] = cat
            self.df['Categoria'] = self.df['Categoria'].fillna('OUTROS')
        else:
            self.df['Categoria'] = self.df['Categoria'].fillna('OUTROS').astype(str).str.strip().str.upper()

        # Ajusta receitas (valores positivos)
        self.df.loc[self.df['Valor'] > 0, 'Categoria'] = self.df.loc[self.df['Valor'] > 0, 'Categoria'].replace('OUTROS', 'RECEITA')

        # 6. Colunas adicionais para agrupamento e análise
        self.df['Mes_Ano'] = self.df['Data'].dt.to_period('M')
        self.df['Tipo'] = self.df['Valor'].apply(lambda x: 'Receita' if x > 0 else 'Despesa')

        return self.df

    def processar_arquivos(self, caminhos_arquivos):
        """
        Lê, normaliza e cruza múltiplos extratos (CSV e PDF),
        consolidando tudo em um único dataset unificado.
        """
        dfs = []
        for caminho in caminhos_arquivos:
            try:
                df_item = self.normalizar_dados(caminho)
                if not df_item.empty:
                    dfs.append(df_item)
            except Exception as e:
                print(f"Erro ao processar '{os.path.basename(caminho)}': {e}")
                raise e

        if not dfs:
            raise ValueError("Nenhum dado válido pôde ser extraído dos arquivos fornecidos.")

        df_consolidado = pd.concat(dfs, ignore_index=True)
        # Ordena cronologicamente
        df_consolidado = df_consolidado.sort_values(by='Data').reset_index(drop=True)
        self.df = df_consolidado
        return df_consolidado

    def analise_combinatoria_economia(self):
        """
        Aplica análise combinatória para encontrar os melhores cenários de corte de gastos.
        Pega as top categorias de despesas (ignorando 'MORADIA' e 'ESSENCIAL').
        """
        print("[2/5] Realizando Análise Combinatória de cenários...")

        # Filtra apenas despesas (valores negativos)
        despesas = self.df[self.df['Valor'] < 0].copy()
        despesas['Valor_Absoluto'] = despesas['Valor'].abs()

        # Agrupa por categoria e calcula médias mensais
        num_meses = max(1, self.df['Mes_Ano'].nunique())
        gastos_por_categoria = despesas.groupby('Categoria')['Valor_Absoluto'].sum() / num_meses

        # Remove categorias essenciais para a simulação de cortes
        categorias_nao_essenciais = gastos_por_categoria.drop(
            labels=['ALUGUEL', 'MERCADO', 'SAÚDE'], errors='ignore'
        )

        # Pega as maiores despesas não essenciais (até 5)
        top_cortes = categorias_nao_essenciais.nlargest(5).to_dict()
        categorias = list(top_cortes.keys())
        cenarios = []

        # Análise Combinatória: Combinações de corte
        tamanhos = [k for k in [2, 3] if k <= len(categorias)]
        if not tamanhos and len(categorias) == 1:
            tamanhos = [1]

        for tamanho_corte in tamanhos:
            combinacoes = list(itertools.combinations(categorias, tamanho_corte))
            for combo in combinacoes:
                economia_mensal = sum(top_cortes[cat] for cat in combo)
                cenarios.append({
                    'Categorias_Cortadas': combo,
                    'Economia_Mensal_Estimada': economia_mensal
                })

        # Ordena para pegar os cenários mais impactantes
        cenarios_ordenados = sorted(cenarios, key=lambda x: x['Economia_Mensal_Estimada'], reverse=True)[:3]
        return cenarios_ordenados, gastos_por_categoria

    def gerar_relatorio_ia(self, cenarios, gastos, bancos=None):
        """
        Usa a API do Groq para analisar os dados e os cenários combinatórios e gerar um relatório.
        """
        print("[3/5] Consultando Inteligência Artificial (Groq)...")

        resumo_gastos = gastos.nlargest(8).to_dict()
        texto_cenarios = "\n".join([f"- Cortando {c['Categorias_Cortadas']}: Economia de R$ {c['Economia_Mensal_Estimada']:.2f}/mês" for c in cenarios])
        bancos_txt = f"Contas bancárias integradas na análise: {', '.join(bancos)}.\n" if bancos and len(bancos) > 1 else ""

        prompt = f"""
        Você é um consultor financeiro sênior e cientista de dados.
        {bancos_txt}Analise o resumo consolidado de gastos médios mensais do seu cliente (cruzamento de contas):
        {resumo_gastos}

        Fizemos uma análise combinatória e encontramos estes cenários de economia cortando categorias não essenciais:
        {texto_cenarios}

        Com base nesses dados unificados, escreva um relatório profissional e direto ao ponto com a seguinte estrutura:
        1. **Resumo da Situação:** Uma análise do fluxo e perfil de gastos consolidado.
        2. **Insights Ocultos:** Padrões ou gargalos que você notou no cruzamento dos dados.
        3. **Dicas Baseadas na Análise Combinatória:** Avalie os cenários apresentados e sugira qual é o mais realista.
        4. **Próximos Passos:** 3 ações claras e imediatas para melhorar a saúde financeira.

        Responda em Português (Brasil).
        """

        try:
            chat_completion = self.client.chat.completions.create(
                messages=[{"role": "user", "content": prompt}],
                model=self.modelo_ia,
                temperature=0.5,
                max_tokens=800,
            )
            return chat_completion.choices[0].message.content
        except Exception as e:
            return f"Erro ao gerar relatório via IA: {e}"

    def exportar_resultados(self, df, relatorio_ia):
        """
        Gera uma pasta de trabalho (.xlsx) perfeitamente legível no Google Sheets e no Excel,
        contendo Dashboard visual, gráficos completos, cruzamento de bancos, transações e insights da IA.
        """
        print("[4/5] Gerando planilha compatível com Google Sheets com Gráficos e Dashboards...")

        df['Mes_Ano_Str'] = df['Mes_Ano'].astype(str)

        wb = openpyxl.Workbook()

        cor_primaria = "1F4E79"
        cor_cinza = "F2F4F7"
        fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        fill_cabecalho = PatternFill(start_color=cor_primaria, end_color=cor_primaria, fill_type="solid")

        # -------------------------------------------------------------
        # ABA 1: Dashboard & Gráficos
        # -------------------------------------------------------------
        ws_dash = wb.active
        ws_dash.title = "Dashboard & Gráficos"
        ws_dash.views.sheetView[0].showGridLines = True

        ws_dash["A1"] = "📊 DASHBOARD FINANCEIRO & CONTROLE DE GASTOS"
        ws_dash["A1"].font = Font(name="Calibri", size=16, bold=True, color=cor_primaria)

        # KPIs Gerais
        receitas_total = float(df[df['Valor'] > 0]['Valor'].sum())
        despesas_total = float(df[df['Valor'] < 0]['Valor'].abs().sum())
        saldo_total = receitas_total - despesas_total

        ws_dash["A3"] = "Total Receitas"
        ws_dash["A4"] = receitas_total
        ws_dash["B3"] = "Total Despesas"
        ws_dash["B4"] = despesas_total
        ws_dash["C3"] = "Saldo Consolidado"
        ws_dash["C4"] = saldo_total

        for col_l in ["A", "B", "C"]:
            ws_dash[f"{col_l}3"].font = Font(name="Calibri", size=10, bold=True, color="595959")
            ws_dash[f"{col_l}3"].alignment = Alignment(horizontal="center")
            ws_dash[f"{col_l}3"].fill = PatternFill(start_color=cor_cinza, fill_type="solid")
            ws_dash[f"{col_l}4"].font = Font(name="Calibri", size=13, bold=True, color=cor_primaria)
            ws_dash[f"{col_l}4"].number_format = "R$ #,##0.00"
            ws_dash[f"{col_l}4"].alignment = Alignment(horizontal="center")
            ws_dash[f"{col_l}4"].fill = PatternFill(start_color=cor_cinza, fill_type="solid")

        # Tabela 1: Despesas por Categoria para alimentar os gráficos
        despesas_df = df[df['Valor'] < 0].copy()
        despesas_df['Valor_Absoluto'] = despesas_df['Valor'].abs()
        resumo_cat = despesas_df.groupby('Categoria')['Valor_Absoluto'].sum().sort_values(ascending=False).reset_index()

        ws_dash["A6"] = "Categoria"
        ws_dash["B6"] = "Total Gasto (R$)"
        ws_dash["A6"].font = fonte_cabecalho
        ws_dash["A6"].fill = fill_cabecalho
        ws_dash["B6"].font = fonte_cabecalho
        ws_dash["B6"].fill = fill_cabecalho

        linha_atual = 7
        for _, r in resumo_cat.iterrows():
            ws_dash.cell(row=linha_atual, column=1, value=str(r['Categoria']))
            c_val = ws_dash.cell(row=linha_atual, column=2, value=float(r['Valor_Absoluto']))
            c_val.number_format = "R$ #,##0.00"
            linha_atual += 1

        total_cat_row = linha_atual
        if len(resumo_cat) > 0:
            ws_dash.cell(row=total_cat_row, column=1, value="TOTAL DESPESAS").font = Font(bold=True)
            c_total = ws_dash.cell(row=total_cat_row, column=2, value=f"=SUM(B7:B{total_cat_row-1})")
            c_total.font = Font(bold=True)
            c_total.number_format = "R$ #,##0.00"

            # Gráfico de Barras: Maiores Despesas
            chart_bar = BarChart()
            chart_bar.type = "col"
            chart_bar.style = 10
            chart_bar.title = "Maiores Despesas por Categoria (R$)"
            chart_bar.y_axis.title = "Total (R$)"
            chart_bar.x_axis.title = "Categoria"
            chart_bar.width = 17
            chart_bar.height = 11
            chart_bar.legend = None

            data_ref = Reference(ws_dash, min_col=2, min_row=6, max_row=total_cat_row-1)
            cats_ref = Reference(ws_dash, min_col=1, min_row=7, max_row=total_cat_row-1)
            chart_bar.add_data(data_ref, titles_from_data=True)
            chart_bar.set_categories(cats_ref)
            ws_dash.add_chart(chart_bar, "D3")

            # Gráfico de Pizza: Distribuição Percentual
            chart_pie = PieChart()
            chart_pie.title = "Distribuição Percentual de Gastos"
            chart_pie.width = 15
            chart_pie.height = 11
            chart_pie.add_data(data_ref, titles_from_data=True)
            chart_pie.set_categories(cats_ref)
            ws_dash.add_chart(chart_pie, "M3")

        proxima_linha = total_cat_row + 3

        # Tabela 2 e Gráfico: Comparativo por Banco / Conta (se houver mais de 1 conta)
        if df['Banco'].nunique() > 1:
            resumo_banco = df.groupby(['Banco', 'Tipo'])['Valor'].sum().unstack(fill_value=0).reset_index()
            if 'Receita' not in resumo_banco: resumo_banco['Receita'] = 0.0
            if 'Despesa' not in resumo_banco: resumo_banco['Despesa'] = 0.0
            resumo_banco['Despesa'] = resumo_banco['Despesa'].abs()

            ws_dash.cell(row=proxima_linha, column=1, value="Banco / Conta").font = fonte_cabecalho
            ws_dash.cell(row=proxima_linha, column=1).fill = fill_cabecalho
            ws_dash.cell(row=proxima_linha, column=2, value="Receitas").font = fonte_cabecalho
            ws_dash.cell(row=proxima_linha, column=2).fill = fill_cabecalho
            ws_dash.cell(row=proxima_linha, column=3, value="Despesas").font = fonte_cabecalho
            ws_dash.cell(row=proxima_linha, column=3).fill = fill_cabecalho

            curr_b_row = proxima_linha + 1
            for _, br in resumo_banco.iterrows():
                ws_dash.cell(row=curr_b_row, column=1, value=str(br['Banco']))
                c_rec = ws_dash.cell(row=curr_b_row, column=2, value=float(br.get('Receita', 0)))
                c_rec.number_format = "R$ #,##0.00"
                c_desp = ws_dash.cell(row=curr_b_row, column=3, value=float(br.get('Despesa', 0)))
                c_desp.number_format = "R$ #,##0.00"
                curr_b_row += 1

            chart_banco = BarChart()
            chart_banco.type = "col"
            chart_banco.style = 12
            chart_banco.title = "Cruzamento de Contas: Receitas vs Despesas por Banco"
            chart_banco.y_axis.title = "Valor (R$)"
            chart_banco.x_axis.title = "Banco / Conta"
            chart_banco.width = 17
            chart_banco.height = 11

            b_data = Reference(ws_dash, min_col=2, max_col=3, min_row=proxima_linha, max_row=curr_b_row-1)
            b_cats = Reference(ws_dash, min_col=1, min_row=proxima_linha+1, max_row=curr_b_row-1)
            chart_banco.add_data(b_data, titles_from_data=True)
            chart_banco.set_categories(b_cats)
            ws_dash.add_chart(chart_banco, f"D{proxima_linha}")

            proxima_linha = curr_b_row + 2

        # Tabela 3 e Gráfico: Evolução Mensal (quando há mais de um mês)
        if df['Mes_Ano'].nunique() > 1:
            resumo_mes = df.groupby(['Mes_Ano', 'Tipo'])['Valor'].sum().abs().unstack(fill_value=0).reset_index()
            resumo_mes['Mes_Ano_Str'] = resumo_mes['Mes_Ano'].astype(str)

            ws_dash.cell(row=proxima_linha, column=1, value="Mês").font = fonte_cabecalho
            ws_dash.cell(row=proxima_linha, column=1).fill = fill_cabecalho
            ws_dash.cell(row=proxima_linha, column=2, value="Receitas").font = fonte_cabecalho
            ws_dash.cell(row=proxima_linha, column=2).fill = fill_cabecalho
            ws_dash.cell(row=proxima_linha, column=3, value="Despesas").font = fonte_cabecalho
            ws_dash.cell(row=proxima_linha, column=3).fill = fill_cabecalho

            curr_m_row = proxima_linha + 1
            for _, mr in resumo_mes.iterrows():
                ws_dash.cell(row=curr_m_row, column=1, value=str(mr['Mes_Ano_Str']))
                c_rec = ws_dash.cell(row=curr_m_row, column=2, value=float(mr.get('Receita', 0)))
                c_rec.number_format = "R$ #,##0.00"
                c_desp = ws_dash.cell(row=curr_m_row, column=3, value=float(mr.get('Despesa', 0)))
                c_desp.number_format = "R$ #,##0.00"
                curr_m_row += 1

            chart_mensal = BarChart()
            chart_mensal.type = "col"
            chart_mensal.style = 11
            chart_mensal.title = "Evolução Mensal: Receitas vs Despesas"
            chart_mensal.y_axis.title = "Valor (R$)"
            chart_mensal.x_axis.title = "Mês"
            chart_mensal.width = 17
            chart_mensal.height = 11

            m_data = Reference(ws_dash, min_col=2, max_col=3, min_row=proxima_linha, max_row=curr_m_row-1)
            m_cats = Reference(ws_dash, min_col=1, min_row=proxima_linha+1, max_row=curr_m_row-1)
            chart_mensal.add_data(m_data, titles_from_data=True)
            chart_mensal.set_categories(m_cats)
            ws_dash.add_chart(chart_mensal, f"M{proxima_linha if df['Banco'].nunique() > 1 else total_cat_row + 3}")

        ws_dash.column_dimensions["A"].width = 24
        ws_dash.column_dimensions["B"].width = 20
        ws_dash.column_dimensions["C"].width = 20

        # -------------------------------------------------------------
        # ABA 2: Transações Consolidadas
        # -------------------------------------------------------------
        ws_base = wb.create_sheet(title="Transações")
        ws_base.views.sheetView[0].showGridLines = True

        col_nomes = ['Data', 'Banco / Conta', 'Descrição', 'Valor (R$)', 'Categoria', 'Tipo', 'Mês/Ano']
        for c_idx, nome in enumerate(col_nomes, 1):
            cell = ws_base.cell(row=1, column=c_idx, value=nome)
            cell.font = fonte_cabecalho
            cell.fill = fill_cabecalho
            cell.alignment = Alignment(horizontal="center")

        for r_idx, (_, row) in enumerate(df.iterrows(), 2):
            ws_base.cell(row=r_idx, column=1, value=pd.to_datetime(row['Data']).strftime('%d/%m/%Y'))
            ws_base.cell(row=r_idx, column=2, value=str(row.get('Banco', 'Geral')))
            ws_base.cell(row=r_idx, column=3, value=str(row.get('Descricao', '')))
            c_val = ws_base.cell(row=r_idx, column=4, value=float(row.get('Valor', 0)))
            c_val.number_format = "R$ #,##0.00"
            ws_base.cell(row=r_idx, column=5, value=str(row.get('Categoria', '')))
            ws_base.cell(row=r_idx, column=6, value=str(row.get('Tipo', '')))
            ws_base.cell(row=r_idx, column=7, value=str(row.get('Mes_Ano_Str', '')))

        for col in ws_base.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws_base.column_dimensions[col_letter].width = max(max_len + 3, 14)

        # -------------------------------------------------------------
        # ABA 3: Resumo por Categoria
        # -------------------------------------------------------------
        ws_resumo = wb.create_sheet(title="Resumo por Categoria")
        ws_resumo.views.sheetView[0].showGridLines = True

        resumo_geral = df.groupby(['Categoria', 'Tipo'])['Valor'].sum().reset_index()
        ws_resumo.cell(row=1, column=1, value="Categoria").font = fonte_cabecalho
        ws_resumo.cell(row=1, column=1).fill = fill_cabecalho
        ws_resumo.cell(row=1, column=2, value="Tipo").font = fonte_cabecalho
        ws_resumo.cell(row=1, column=2).fill = fill_cabecalho
        ws_resumo.cell(row=1, column=3, value="Total (R$)").font = fonte_cabecalho
        ws_resumo.cell(row=1, column=3).fill = fill_cabecalho

        for r_idx, (_, r) in enumerate(resumo_geral.iterrows(), 2):
            ws_resumo.cell(row=r_idx, column=1, value=str(r['Categoria']))
            ws_resumo.cell(row=r_idx, column=2, value=str(r['Tipo']))
            c = ws_resumo.cell(row=r_idx, column=3, value=float(r['Valor']))
            c.number_format = "R$ #,##0.00"

        ws_resumo.column_dimensions["A"].width = 24
        ws_resumo.column_dimensions["B"].width = 16
        ws_resumo.column_dimensions["C"].width = 20

        # -------------------------------------------------------------
        # ABA 4: Insights e Recomendações IA
        # -------------------------------------------------------------
        ws_ia = wb.create_sheet(title="Insights e Recomendações IA")
        ws_ia.views.sheetView[0].showGridLines = True

        ws_ia["A1"] = "💡 RELATÓRIO ESTRATÉGICO & INSIGHTS DA IA"
        ws_ia["A1"].font = Font(name="Calibri", size=14, bold=True, color=cor_primaria)

        linhas_relatorio = str(relatorio_ia).strip().split('\n')
        for r_idx, linha in enumerate(linhas_relatorio, 3):
            cell = ws_ia.cell(row=r_idx, column=1, value=linha)
            linha_s = linha.strip()
            if linha_s.startswith('#'):
                cell.font = Font(name="Calibri", size=13, bold=True, color=cor_primaria)
            elif linha_s.startswith('**') or linha_s.startswith('###') or any(linha_s.startswith(f"{n}.") for n in range(1, 10)):
                cell.font = Font(name="Calibri", size=11, bold=True, color="262626")
            else:
                cell.font = Font(name="Calibri", size=11, color="333333")

        ws_ia.column_dimensions["A"].width = 110

        wb.save('Relatorio_Financeiro_Final.xlsx')
        print("[5/5] Concluído! Planilha 'Relatorio_Financeiro_Final.xlsx' gerada com sucesso (pronta para Google Sheets e Excel).")

class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Analisador Financeiro Multibancos (AI & Data Science)")
        self.root.geometry("740x700")
        self.root.configure(padx=20, pady=15, bg="#F4F6F9")

        self.chave_groq = os.getenv("GROQ_API_KEY", "").strip()
        self.arquivos_selecionados = []

        # --- Cabeçalho ---
        header_frame = tk.Frame(root, bg="#F4F6F9")
        header_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(header_frame, text="📊 Pipeline de Cruzamento de Extratos Bancários", font=("Arial", 16, "bold"), bg="#F4F6F9", fg="#1F4E79").pack(anchor="w")
        tk.Label(header_frame, text="Importe múltiplos extratos (CSV e PDF) de diferentes bancos para análise e cruzamento.", font=("Arial", 10), bg="#F4F6F9", fg="#6C757D").pack(anchor="w")

        # --- Área de Drag & Drop ---
        self.drop_frame = tk.Frame(root, bg="#FFFFFF", bd=2, relief="groove", padx=15, pady=15)
        self.drop_frame.pack(fill=tk.X, pady=(0, 8))

        lbl_icon = tk.Label(self.drop_frame, text="📥", font=("Arial", 26), bg="#FFFFFF")
        lbl_icon.pack()

        lbl_drop = tk.Label(self.drop_frame, text="Arraste e solte seus extratos aqui (CSV ou PDF)", font=("Arial", 12, "bold"), bg="#FFFFFF", fg="#1F4E79")
        lbl_drop.pack(pady=(2, 2))

        lbl_sub = tk.Label(self.drop_frame, text="Suporta múltiplos extratos de bancos brasileiros (Nubank, Inter, Itaú, BB, etc.)", font=("Arial", 9), bg="#FFFFFF", fg="#6C757D")
        lbl_sub.pack()

        # Ativa drag and drop se disponível no sistema
        if TEM_DND:
            try:
                self.drop_frame.drop_target_register(DND_FILES)
                self.drop_frame.dnd_bind('<<Drop>>', self.ao_soltar_arquivos)
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind('<<Drop>>', self.ao_soltar_arquivos)
            except Exception:
                pass

        # --- Barra de Ações de Arquivos ---
        bar_frame = tk.Frame(root, bg="#F4F6F9")
        bar_frame.pack(fill=tk.X, pady=(0, 6))

        btn_add = tk.Button(bar_frame, text="📂 Adicionar Arquivos...", command=self.selecionar_arquivos, font=("Arial", 10, "bold"), bg="#E9ECEF", padx=10)
        btn_add.pack(side=tk.LEFT, padx=(0, 8))

        btn_remove = tk.Button(bar_frame, text="🗑️ Remover Selecionado", command=self.remover_arquivo, font=("Arial", 9), bg="#FFF0F0", padx=8)
        btn_remove.pack(side=tk.LEFT, padx=(0, 8))

        btn_clear = tk.Button(bar_frame, text="🧹 Limpar Todos", command=self.limpar_arquivos, font=("Arial", 9), bg="#F8F9FA", padx=8)
        btn_clear.pack(side=tk.LEFT)

        self.lbl_contador = tk.Label(bar_frame, text="0 arquivo(s) carregado(s)", font=("Arial", 9, "bold"), bg="#F4F6F9", fg="#6C757D")
        self.lbl_contador.pack(side=tk.RIGHT)

        # --- Lista de Arquivos Carregados ---
        list_container = tk.Frame(root, bg="#F4F6F9")
        list_container.pack(fill=tk.X, pady=(0, 8))

        self.listbox_arquivos = tk.Listbox(list_container, height=4, font=("Consolas", 10), selectmode=tk.SINGLE, bg="#FFFFFF")
        scrollbar_list = tk.Scrollbar(list_container, orient="vertical", command=self.listbox_arquivos.yview)
        self.listbox_arquivos.config(yscrollcommand=scrollbar_list.set)

        self.listbox_arquivos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_list.pack(side=tk.RIGHT, fill=tk.Y)

        # --- Botão Processar ---
        self.btn_executar = tk.Button(
            root,
            text="🚀 Cruzar Dados e Gerar Análise Consolidada (Google Sheets / Excel)",
            command=self.iniciar_processamento,
            font=("Arial", 11, "bold"),
            bg="#28A745",
            fg="white",
            state=tk.DISABLED,
            pady=7
        )
        self.btn_executar.pack(fill=tk.X, pady=(0, 8))

        # --- Caixa de Texto (Log e Relatório) ---
        tk.Label(root, text="Logs do Sistema & Relatório da IA:", font=("Arial", 10, "bold"), bg="#F4F6F9", fg="#1F4E79").pack(anchor="w")
        self.txt_log = scrolledtext.ScrolledText(root, width=80, height=10, font=("Consolas", 10), bg="#1e1e1e", fg="#00ff00")
        self.txt_log.pack(fill=tk.BOTH, expand=True, pady=(3, 0))

    def ao_soltar_arquivos(self, event):
        dados = event.data
        if not dados:
            return
        caminhos = self.root.tk.splitlist(dados)
        self.adicionar_caminhos(caminhos)

    def adicionar_caminhos(self, caminhos):
        for c in caminhos:
            c = c.strip('{}').strip()
            if c.startswith('file://'):
                c = c[7:]
            if os.path.isfile(c) and c.lower().endswith(('.csv', '.pdf', '.txt')):
                if c not in self.arquivos_selecionados:
                    self.arquivos_selecionados.append(c)
                    tipo = "PDF" if c.lower().endswith('.pdf') else "CSV"
                    self.listbox_arquivos.insert(tk.END, f"📄 [{tipo}] {os.path.basename(c)}")

        qtd = len(self.arquivos_selecionados)
        if qtd > 0:
            self.lbl_contador.config(text=f"📁 {qtd} extrato(s) pronto(s) para cruzamento", fg="#28A745")
            self.btn_executar.config(state=tk.NORMAL)
        else:
            self.lbl_contador.config(text="0 arquivo(s) carregado(s)", fg="#6C757D")
            self.btn_executar.config(state=tk.DISABLED)

    def selecionar_arquivos(self):
        caminhos = filedialog.askopenfilenames(
            title="Selecione os extratos (CSV ou PDF)",
            filetypes=(
                ("Extratos Bancários", "*.csv;*.pdf;*.txt"),
                ("Arquivos CSV", "*.csv"),
                ("Arquivos PDF", "*.pdf"),
                ("Todos os arquivos", "*.*")
            )
        )
        if caminhos:
            self.adicionar_caminhos(caminhos)

    def remover_arquivo(self):
        selecao = self.listbox_arquivos.curselection()
        if selecao:
            idx = selecao[0]
            self.listbox_arquivos.delete(idx)
            del self.arquivos_selecionados[idx]
            qtd = len(self.arquivos_selecionados)
            if qtd > 0:
                self.lbl_contador.config(text=f"📁 {qtd} extrato(s) pronto(s) para cruzamento", fg="#28A745")
                self.btn_executar.config(state=tk.NORMAL)
            else:
                self.lbl_contador.config(text="0 arquivo(s) carregado(s)", fg="#6C757D")
                self.btn_executar.config(state=tk.DISABLED)

    def limpar_arquivos(self):
        self.arquivos_selecionados.clear()
        self.listbox_arquivos.delete(0, tk.END)
        self.lbl_contador.config(text="0 arquivo(s) carregado(s)", fg="#6C757D")
        self.btn_executar.config(state=tk.DISABLED)

    def log(self, mensagem):
        self.txt_log.insert(tk.END, mensagem + "\n")
        self.txt_log.see(tk.END)
        self.root.update_idletasks()

    def iniciar_processamento(self):
        self.btn_executar.config(state=tk.DISABLED)
        self.txt_log.delete(1.0, tk.END)
        thread = threading.Thread(target=self.processar_dados)
        thread.start()

    def processar_dados(self):
        try:
            if not self.chave_groq:
                chave = simpledialog.askstring("Chave da API Groq", "Insira sua API Key do Groq (ou defina no arquivo .env):", show="*")
                if chave and chave.strip():
                    self.chave_groq = chave.strip()
                else:
                    self.log("❌ ERRO: Nenhuma chave da API Groq fornecida. Operação cancelada.")
                    self.btn_executar.config(state=tk.NORMAL)
                    return

            total_arquivos = len(self.arquivos_selecionados)
            self.log(f"Inicializando motor de IA para cruzar {total_arquivos} extrato(s)...")
            analisador = AnalisadorFinanceiro(api_key=self.chave_groq)

            self.log("[1/5] Lendo, normalizando e cruzando dados de todos os extratos...")
            df_consolidado = analisador.processar_arquivos(self.arquivos_selecionados)
            self.log(f"-> Total de transações unificadas: {len(df_consolidado)}")
            bancos_detectados = df_consolidado['Banco'].unique().tolist()
            self.log(f"-> Contas/Bancos detectados: {', '.join(bancos_detectados)}")

            self.log("[2/5] Rodando Análise Combinatória (Itertools) para cenários de economia...")
            cenarios_combo, gastos_agrupados = analisador.analise_combinatoria_economia()

            self.log("[3/5] Consultando LLM via Groq API. Aguardando insights...")
            relatorio = analisador.gerar_relatorio_ia(cenarios_combo, gastos_agrupados, bancos=bancos_detectados)

            self.log("[4/5] Gerando planilha compatível com Google Sheets com Gráficos e Dashboards...")
            analisador.exportar_resultados(df_consolidado, relatorio)

            self.log("[5/5] ✨ CONCLUÍDO COM SUCESSO! ✨")

            self.log("\n" + "="*50 + "\nRESUMO DO RELATÓRIO DA IA\n" + "="*50 + "\n")
            self.log(relatorio)

            messagebox.showinfo(
                "Sucesso",
                f"Cruzamento concluído com sucesso!\n\n"
                f"- Extratos processados: {total_arquivos}\n"
                f"- Contas/Bancos: {', '.join(bancos_detectados)}\n"
                f"- Total de transações: {len(df_consolidado)}\n\n"
                f"Arquivo gerado: Relatorio_Financeiro_Final.xlsx\n"
                f"(Pronto para abrir diretamente no Google Sheets ou no Excel!)"
            )

        except Exception as e:
            self.log(f"\n❌ ERRO FATAL: {str(e)}")
            messagebox.showerror("Erro de Execução", f"Ocorreu um erro no pipeline de dados:\n\n{str(e)}")
        finally:
            self.btn_executar.config(state=tk.NORMAL)

if __name__ == "__main__":
    if TEM_DND:
        janela_principal = TkinterDnD.Tk()
    else:
        janela_principal = tk.Tk()
    app = AppGUI(janela_principal)
    janela_principal.mainloop()

import os
import json
import re
import threading
import itertools
from datetime import datetime
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import tkinter as tk
import platform
import subprocess
import customtkinter as ctk
from tkinter import filedialog, messagebox
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
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

try:
    import pypdf
    TEM_PYPDF = True
except Exception:
    TEM_PYPDF = False

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# Carrega as variáveis de ambiente
load_dotenv()


# =====================================================================
# MONKEY-PATCH: Correção do Bug de Scroll do CustomTkinter no Linux
# =====================================================================
def _patch_ctk_scroll():
    """
    Corrige o bug no CustomTkinter em ambiente Linux onde event.widget
    pode ser transmitido como string para _check_if_valid_scroll,
    gerando AttributeError: 'str' object has no attribute 'master'.
    """
    original_check = getattr(ctk.CTkScrollableFrame, "_check_if_valid_scroll", None)
    if original_check:
        def _safe_check_if_valid_scroll(self, widget):
            if widget is None:
                return False
            if isinstance(widget, str):
                try:
                    widget = self.nametowidget(widget)
                except Exception:
                    return False
            curr = widget
            while curr is not None:
                if curr == self or curr == getattr(self, '_parent_canvas', None):
                    return True
                curr = getattr(curr, 'master', None)
            return False
        ctk.CTkScrollableFrame._check_if_valid_scroll = _safe_check_if_valid_scroll

_patch_ctk_scroll()


# =====================================================================
# FUNÇÕES DE PIPELINE & MATEMÁTICA FINANCEIRA (MODULARES E TESTÁVEIS)
# =====================================================================

def limpar_valor_monetario(val):
    """
    Normaliza representações monetárias de bancos brasileiros para float numérico:
    Trata strings como 'R$ 1.250,50', '-450,90', '1200.50' ou valores float/int diretos.
    """
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace('R$', '').replace(' ', '')
    if not s:
        return 0.0
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0


# Regras compiladas de Regex para categorização instantânea (Tier 1 do classificador híbrido)
REGRAS_REGEX = [
    ('TRANSPORTE', re.compile(r'\b(UBER|99|99APP|99 TECNOLOGIA|KIM\+|POSTO|COMBUSTIVEL|GASOLINA|ESTACIONAMENTO|METRO|CCR|SEM PARAR|VELOE|AUTO POSTO|IPIRANGA|SHELL|BR PETROBRAS)\b', re.IGNORECASE)),
    ('SAÚDE', re.compile(r'\b(FARMACIA|DROGARIA|PAGUE MENOS|DROGASIL|RAIA|HOSPITAL|CLINICA|CONSULTA|MEDIC|SAUDE|LABORATORIO|ODONTO|EXAME|DOR CONSULTORIA)\b', re.IGNORECASE)),
    ('DELIVERY', re.compile(r'\b(IFOOD|RAPPI|DELIVERY|AIQFOME|ZE DELIVERY|HABIB\'S DELIVERY|PEDIDOSJA)\b', re.IGNORECASE)),
    ('RESTAURANTE', re.compile(r'\b(RESTAURANTE|LANCHES?|BURGUER|BURGER|DELICATESSEN|PIZZA|PIZZARIA|PADARIA|CAFE|BAR|DOGAO|CHURRASC|BAKERY|DOCERIA|SUSHI|MCDONALD|SUBWAY|GIRAFAS|BOB\'S|OUTBACK)\b', re.IGNORECASE)),
    ('MERCADO', re.compile(r'\b(ATACADAO|MERCADO|SUPERMERCADO|HIPERMERCADO|ASSAI|CARREFOUR|EXTRA|BIG|HORTIFRUTI|ACOUGUE|PAO DE ACUCAR|SAMS CLUB|SUPER BOM|SUPERBOX)\b', re.IGNORECASE)),
    ('ALUGUEL', re.compile(r'\b(ALUGUEL|CONDOMINIO|IMOBIL|QUINTO ANDAR|QUILTO|HABITACIONAL|LOCACAO)\b', re.IGNORECASE)),
    ('CONTAS DE CONSUMO', re.compile(r'\b(COELB|EMBASA|ENEL|CPFL|SABESP|ENERGISA|LUZ|AGUA|GAS|TELECOM|CLARO|VIVO|TIM|OI|INTERNET|ENERGIA)\b', re.IGNORECASE)),
    ('ASSINATURAS', re.compile(r'\b(NETFLIX|SPOTIFY|PRIME VIDEO|YOUTUBE|APPLE|GOOGLE|MAX|DISNEY|HBO|DEEZER|GLOBOPLAY|PARAMOUNT|CRUNCHYROLL|OPENAI|CHATGPT)\b', re.IGNORECASE)),
    ('INVESTIMENTOS', re.compile(r'\b(B3|CDB|LCI|LCA|TESOURO|INVEST|ATIVO|RENDIMENTOS|DIVIDENDOS|APORTE|XP|CLEAR|RICO|BTG|INTER DTVM|NU INVEST)\b', re.IGNORECASE)),
    ('COMPRAS', re.compile(r'\b(AMAZON|MERCADO LIVRE|SHOPEE|SHEIN|MAGALU|MAGAZINE LUIZA|CASAS BAHIA|ALIEXPRESS|ZARA|RENNER|RIACHUELO|C&A|KABUM|PICHAU)\b', re.IGNORECASE))
]


def classificar_por_regex(descricao):
    """
    Classifica instantaneamente a descrição via Regex compilado de alta precisão.
    Retorna o nome da Categoria ou None caso deva ser enviado ao fallback de IA.
    """
    if not descricao or pd.isna(descricao):
        return None
    s = str(descricao).strip()
    for cat, regex in REGRAS_REGEX:
        if regex.search(s):
            return cat
    return None


def deduplicar_e_tratar_estornos(df):
    """
    Deduplicação Inteligente & Cancelamento de Estornos:
    1. Remove duplicatas estritas originadas por múltiplos uploads do mesmo extrato.
    2. Identifica e neutraliza pares de estorno (+V e -V) ocorridos no mesmo período,
       evitando inflar artificialmente o total de entradas e saídas.
    """
    if df is None or df.empty:
        return df, {"duplicatas_removidas": 0, "estornos_neutralizados": 0}

    df_copia = df.copy()
    linhas_iniciais = len(df_copia)

    # 1. Deduplicação de linhas idênticas
    subset_cols = [c for c in ['Data', 'Descricao', 'Valor', 'Banco'] if c in df_copia.columns]
    if subset_cols:
        df_copia = df_copia.drop_duplicates(subset=subset_cols, keep='first').reset_index(drop=True)
    duplicatas_removidas = linhas_iniciais - len(df_copia)

    # 2. Neutralização de estornos e cancelamentos
    estornos_neutralizados = 0
    termos_estorno = ['ESTORNO', 'CANCELAMENTO', 'REVERSAO', 'REVERSÃO', 'DEVOLUCAO', 'DEVOLUÇÃO', 'REEMBOLSO']
    indices_para_remover = set()

    positivos = df_copia[df_copia['Valor'] > 0]
    negativos = df_copia[df_copia['Valor'] < 0]

    for idx_pos, row_pos in positivos.iterrows():
        val_pos = abs(row_pos['Valor'])
        desc_pos = str(row_pos.get('Descricao', '')).upper()

        candidatos_neg = negativos[
            (negativos['Valor'].abs() == val_pos) &
            (~negativos.index.isin(indices_para_remover))
        ]

        for idx_neg, row_neg in candidatos_neg.iterrows():
            desc_neg = str(row_neg.get('Descricao', '')).upper()

            eh_estorno = (
                any(t in desc_pos for t in termos_estorno) or
                any(t in desc_neg for t in termos_estorno) or
                (len(desc_neg) > 5 and desc_neg[:8] in desc_pos) or
                (len(desc_pos) > 5 and desc_pos[:8] in desc_neg)
            )

            if eh_estorno:
                indices_para_remover.add(idx_pos)
                indices_para_remover.add(idx_neg)
                estornos_neutralizados += 1
                break

    if indices_para_remover:
        df_copia = df_copia.drop(index=list(indices_para_remover)).reset_index(drop=True)

    info = {
        "duplicatas_removidas": duplicatas_removidas,
        "estornos_neutralizados": estornos_neutralizados
    }
    return df_copia, info


def calcular_juros_compostos(pmt, taxa_mensal=0.0085, meses=12):
    """
    Calcula a projeção de valor futuro de aportes mensais recorrentes (anuidade com juros compostos):
    FV = PMT * (((1 + i)^n - 1) / i)
    Retorna: (total_investido, montante_acumulado, juros_auferidos)
    """
    if pmt <= 0 or meses <= 0:
        return 0.0, 0.0, 0.0
    if taxa_mensal <= 0:
        total = pmt * meses
        return total, total, 0.0

    fv = pmt * (((1 + taxa_mensal) ** meses - 1) / taxa_mensal)
    total_investido = pmt * meses
    juros = fv - total_investido
    return total_investido, fv, juros


def otimizar_metas_combinatoria(gastos_por_categoria, meta_economia=None, taxa_mensal=0.0085):
    """
    Lógica de Análise Combinatória Avançada (Variação do 0-1 Knapsack / Subset Sum Problem):
    
    1. FORMULAÇÃO MATEMÁTICA:
       Dado um conjunto de categorias não essenciais C = {C_1, C_2, ..., C_k} com gastos médios
       mensais v_i > 0 e pesos subjetivos de impacto perceptual w_i (onde categorias como Assinaturas
       e Delivery têm menor fricção de corte do que Restaurante ou Compras), e uma meta financeira M (ex: R$ 450).
       
       Objetivo: Encontrar o subconjunto S ⊆ C que minimize o impacto no padrão de vida do usuário:
       - Restrição primária: sum(v_i para i ∈ S) >= M (bater a meta pretendida).
       - Função objetivo de minimização:
         F(S) = (sum(v_i) - M) + lambda_1 * |S| + lambda_2 * sum(w_i)
         Isto penaliza excessos desnecessários além da meta, minimiza a quantidade de categorias sacrificadas
         (|S| menor significa menor esforço de adaptação de hábitos) e prioriza cortes em gastos de baixo atrito.
         
    2. CENÁRIOS GERADOS:
       - Plano Otimizado: Solução da minimização descrita acima.
       - Plano Agressivo: Maior corte viável para aporte máximo.
       - Plano Equilibrado: Alternativa intermediária de diversificação de cortes.
       
    3. PROJEÇÃO DE JUROS COMPOSTOS:
       Para cada plano, calcula o montante acumulado se o valor mensal for aportado a 100% do CDI
       (taxa de ~0.85% ao mês) nos horizontes de 1 ano (12m), 5 anos (60m) e 10 anos (120m).
    """
    pesos_friccao = {
        'ASSINATURAS': 1.0,
        'DELIVERY': 1.2,
        'OUTROS': 1.3,
        'TRANSPORTE': 1.5,
        'COMPRAS': 1.8,
        'RESTAURANTE': 2.0
    }
    categorias_essenciais = {'ALUGUEL', 'MERCADO', 'SAÚDE', 'CONTAS DE CONSUMO', 'RECEITA', 'INVESTIMENTOS'}
    nao_essenciais = gastos_por_categoria.drop(
        labels=[c for c in categorias_essenciais if c in gastos_por_categoria.index],
        errors='ignore'
    )

    if nao_essenciais.empty:
        return [], gastos_por_categoria

    itens = nao_essenciais.nlargest(6).to_dict()
    nomes = list(itens.keys())
    candidatos = []

    for r in range(1, len(nomes) + 1):
        for combo in itertools.combinations(nomes, r):
            soma = sum(itens[k] for k in combo)
            friccao = sum(pesos_friccao.get(k, 1.5) for k in combo)
            candidatos.append({
                'Categorias_Cortadas': list(combo),
                'Economia_Mensal_Estimada': soma,
                'friccao': friccao,
                'qtd_cortes': len(combo)
            })

    cenarios_selecionados = []
    if meta_economia and meta_economia > 0:
        batem_meta = [c for c in candidatos if c['Economia_Mensal_Estimada'] >= meta_economia]
        if batem_meta:
            batem_meta.sort(
                key=lambda x: (x['Economia_Mensal_Estimada'] - meta_economia, x['qtd_cortes'], x['friccao'])
            )
            otimizado = dict(batem_meta[0])
            otimizado['nome'] = 'Plano Otimizado (Menor Impacto / Knapsack)'
            cenarios_selecionados.append(otimizado)

            agressivo = dict(sorted(candidatos, key=lambda x: x['Economia_Mensal_Estimada'], reverse=True)[0])
            if agressivo['Categorias_Cortadas'] != otimizado['Categorias_Cortadas']:
                agressivo['nome'] = 'Plano Agressivo (Máxima Poupança)'
                cenarios_selecionados.append(agressivo)

            intermediarios = [
                c for c in batem_meta
                if c['Categorias_Cortadas'] != otimizado['Categorias_Cortadas'] and
                (len(cenarios_selecionados) < 2 or c['Categorias_Cortadas'] != agressivo['Categorias_Cortadas'])
            ]
            if intermediarios:
                equilibrado = dict(intermediarios[0])
                equilibrado['nome'] = 'Plano Equilibrado'
                cenarios_selecionados.append(equilibrado)
        else:
            maximo = dict(sorted(candidatos, key=lambda x: x['Economia_Mensal_Estimada'], reverse=True)[0])
            maximo['nome'] = f'Plano Limite (Meta de R$ {meta_economia:.2f} excede cortes possíveis)'
            cenarios_selecionados.append(maximo)
    else:
        ordenados = sorted(candidatos, key=lambda x: x['Economia_Mensal_Estimada'], reverse=True)
        for i, c in enumerate(ordenados[:3]):
            cd = dict(c)
            cd['nome'] = f'Cenário Recomendado {i+1}'
            cenarios_selecionados.append(cd)

    for c in cenarios_selecionados:
        eco = c['Economia_Mensal_Estimada']
        c['projecoes'] = {}
        for anos, meses in [(1, 12), (5, 60), (10, 120)]:
            inv, mnt, jur = calcular_juros_compostos(eco, taxa_mensal, meses)
            c['projecoes'][anos] = {'meses': meses, 'investido': inv, 'montante': mnt, 'juros': jur}

    return cenarios_selecionados, gastos_por_categoria

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
        """Classificação instantânea baseada em padrões bancários frequentes (Tier 1 via Regex)."""
        return classificar_por_regex(desc)

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

    def _carregar_pdf(self, caminho_arquivo, senha=None, callback_senha=None):
        """
        Extrai transações de extratos em PDF (C6, Nubank, Itaú, Inter, BB, etc.)
        utilizando regex linha a linha de alta precisão ou tabelas estruturadas com pdfplumber.
        Suporta arquivos protegidos com pop-up discreto de senha e descriptografia 100% em memória,
        sem salvar nenhum arquivo descriptografado intermediário no disco.
        """
        if not TEM_PDF:
            raise ImportError("A biblioteca 'pdfplumber' não está instalada. Execute: pip install pdfplumber")

        regex_val = re.compile(r"([-+]?\s*R\$\s*[\d\.]+,\d{2}|[-+]?\s*[\d\.]+,\d{2})\s*$")
        regex_date = re.compile(r"^(\d{2}[/-]\d{2}(?:[/-]\d{2,4})?)(?:\s+(\d{2}[/-]\d{2}(?:[/-]\d{2,4})?))?\s+(.*)$")

        senha_usada = senha
        nome_arq = os.path.basename(caminho_arquivo)

        # 0. Detecção antecipada de senha com pypdf caso disponível
        if TEM_PYPDF:
            try:
                reader = pypdf.PdfReader(caminho_arquivo)
                if reader.is_encrypted:
                    if not senha_usada and callback_senha:
                        senha_usada = callback_senha(nome_arq)
                    if not senha_usada:
                        raise PermissionError(f"O extrato '{nome_arq}' está protegido por senha e nenhuma chave foi informada.")
                    dec = reader.decrypt(senha_usada)
                    if dec == 0:
                        raise PermissionError(f"Senha incorreta para descriptografar '{nome_arq}'.")
            except PermissionError:
                raise
            except Exception:
                pass

        def _abrir_pdf_em_memoria():
            nonlocal senha_usada
            try:
                return pdfplumber.open(caminho_arquivo, password=senha_usada)
            except Exception as err:
                msg_err = str(err).lower()
                if any(k in msg_err for k in ["password", "encrypted", "senha", "pwd"]) and callback_senha:
                    senha_usada = callback_senha(nome_arq)
                    if senha_usada:
                        return pdfplumber.open(caminho_arquivo, password=senha_usada)
                raise PermissionError(f"O extrato '{nome_arq}' está criptografado e requer senha válida.")

        registros = []
        ano_detectado = str(datetime.now().year)

        with _abrir_pdf_em_memoria() as pdf:
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
        with _abrir_pdf_em_memoria() as pdf:
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

        raise ValueError(f"Não foi possível extrair dados textuais de '{nome_arq}'. Se o PDF for uma imagem digitalizada, exporte em formato CSV pelo app do seu banco.")

    def normalizar_dados(self, caminho_arquivo, callback_senha=None):
        """
        Lê e normaliza os dados do extrato (CSV ou PDF) de qualquer banco brasileiro,
        usando classificador híbrido em cascata (Regex Tier 1 + Fallback Groq LLM Tier 2).
        """
        nome_arquivo = os.path.basename(caminho_arquivo)
        print(f"Lendo e normalizando '{nome_arquivo}'...")

        # 1. Leitura do arquivo conforme extensão
        if caminho_arquivo.lower().endswith('.pdf'):
            self.df = self._carregar_pdf(caminho_arquivo, callback_senha=callback_senha)
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

        if isinstance(self.df['Data'], pd.DataFrame):
            self.df['Data'] = self.df['Data'].iloc[:, 0]
        if isinstance(self.df['Valor'], pd.DataFrame):
            self.df['Valor'] = self.df['Valor'].iloc[:, 0]

        if 'Descricao_Completa' in self.df.columns:
            self.df['Descricao'] = self.df['Descricao_Completa']
        elif 'Descricao' not in self.df.columns:
            self.df['Descricao'] = "Transação"

        # 3. Limpeza e normalização do Valor com função utilitária testada
        self.df['Valor'] = self.df['Valor'].apply(limpar_valor_monetario)

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

        # 5. Classificador Híbrido em Cascata (Tier 1: Regex -> Tier 2: LLM Groq em lote)
        if 'Categoria' not in self.df.columns or self.df['Categoria'].dropna().empty:
            self.df['Categoria'] = self.df['Descricao'].apply(classificar_por_regex)
            
            # Apenas despesas desconhecidas (sem match no regex) vão para o Groq
            nao_categorizadas = self.df.loc[
                (self.df['Valor'] < 0) & (self.df['Categoria'].isna()), 'Descricao'
            ].dropna().unique()
            
            if len(nao_categorizadas) > 0:
                print(f"-> {len(nao_categorizadas)} transações desconhecidas enviadas para Groq LLM...")
                mapa_ia = self._categorizar_transacoes_com_ia(nao_categorizadas)
                for desc, cat in mapa_ia.items():
                    self.df.loc[self.df['Descricao'] == desc, 'Categoria'] = cat
            
            self.df['Categoria'] = self.df['Categoria'].fillna('OUTROS')
        else:
            self.df['Categoria'] = self.df['Categoria'].fillna('OUTROS').astype(str).str.strip().str.upper()

        # Ajusta receitas (valores positivos)
        self.df.loc[self.df['Valor'] > 0, 'Categoria'] = self.df.loc[self.df['Valor'] > 0, 'Categoria'].replace('OUTROS', 'RECEITA')

        # 6. Deduplicação inteligente e tratamento de estornos no arquivo atual
        self.df, _ = deduplicar_e_tratar_estornos(self.df)

        # 7. Colunas adicionais para agrupamento e análise
        self.df['Mes_Ano'] = self.df['Data'].dt.to_period('M')
        self.df['Tipo'] = self.df['Valor'].apply(lambda x: 'Receita' if x > 0 else 'Despesa')

        return self.df

    def processar_arquivos(self, caminhos_arquivos, callback_senha=None):
        """
        Lê, normaliza e cruza múltiplos extratos (CSV e PDF),
        consolidando tudo em um único dataset unificado com deduplicação inteligente.
        """
        dfs = []
        for caminho in caminhos_arquivos:
            try:
                df_item = self.normalizar_dados(caminho, callback_senha=callback_senha)
                if not df_item.empty:
                    dfs.append(df_item)
            except Exception as e:
                print(f"Erro ao processar '{os.path.basename(caminho)}': {e}")
                raise e

        if not dfs:
            raise ValueError("Nenhum dado válido pôde ser extraído dos arquivos fornecidos.")

        df_consolidado = pd.concat(dfs, ignore_index=True)
        
        # Deduplicação e tratamento de estornos cross-extratos
        df_consolidado, self.info_deduplicacao = deduplicar_e_tratar_estornos(df_consolidado)
        
        # Ordena cronologicamente
        df_consolidado = df_consolidado.sort_values(by='Data').reset_index(drop=True)
        self.df = df_consolidado
        return df_consolidado

    def analise_combinatoria_economia(self, meta_economia=None, taxa_mensal=0.0085):
        """
        Aplica o Otimizador de Metas (Knapsack Problem / Subset Sum) para encontrar
        os melhores cenários de economia com menor impacto perceptual no padrão de vida.
        Calcula também a projeção de juros compostos (1, 5 e 10 anos) a 100% do CDI.
        """
        print("[2/5] Realizando Análise Combinatória Avançada (Knapsack & Projeções)...")

        # Filtra apenas despesas (valores negativos)
        despesas = self.df[self.df['Valor'] < 0].copy()
        despesas['Valor_Absoluto'] = despesas['Valor'].abs()

        # Agrupa por categoria e calcula médias mensais
        num_meses = max(1, self.df['Mes_Ano'].nunique())
        gastos_por_categoria = despesas.groupby('Categoria')['Valor_Absoluto'].sum() / num_meses

        cenarios, _ = otimizar_metas_combinatoria(
            gastos_por_categoria,
            meta_economia=meta_economia,
            taxa_mensal=taxa_mensal
        )
        return cenarios, gastos_por_categoria

    def gerar_relatorio_ia(self, cenarios, gastos, bancos=None):
        """
        Usa a API do Groq para analisar os dados, cenários combinatórios e
        projeções de juros compostos para gerar um relatório executivo.
        """
        print("[3/5] Consultando Inteligência Artificial (Groq)...")

        resumo_gastos = gastos.nlargest(8).to_dict()
        
        linhas_cenarios = []
        for c in cenarios:
            nome = c.get('nome', 'Cenário')
            cats = ", ".join(c.get('Categorias_Cortadas', []))
            eco = c.get('Economia_Mensal_Estimada', 0.0)
            proj = c.get('projecoes', {})
            txt_proj = ""
            if 5 in proj:
                txt_proj = f" | Projeção CDI em 5 anos: R$ {proj[5]['montante']:.2f} (Juros: R$ {proj[5]['juros']:.2f})"
            linhas_cenarios.append(f"- **{nome}**: Cortando [{cats}] -> Economia de R$ {eco:.2f}/mês{txt_proj}")
            
        texto_cenarios = "\n".join(linhas_cenarios)
        bancos_txt = f"Contas bancárias integradas na análise: {', '.join(bancos)}.\n" if bancos and len(bancos) > 1 else ""

        prompt = f"""
        Você é um consultor financeiro sênior e cientista de dados.
        {bancos_txt}Analise o resumo consolidado de gastos médios mensais do seu cliente (cruzamento de contas):
        {resumo_gastos}

        Fizemos uma análise combinatória matemática (Knapsack Problem) e calculamos projeções de juros compostos a 100% do CDI:
        {texto_cenarios}

        Com base nesses dados unificados, escreva um relatório profissional e direto ao ponto com a seguinte estrutura:
        1. **Resumo da Situação:** Uma análise do fluxo e perfil de gastos consolidado.
        2. **Insights Ocultos:** Padrões ou gargalos que você notou no cruzamento dos dados.
        3. **Dicas Baseadas na Análise Combinatória:** Avalie os cenários apresentados e sugira qual é o mais realista para o cliente aportar.
        4. **Impacto dos Juros Compostos:** Explique o efeito multiplicador de investir a economia mensal no CDI/ativos geradores de renda.
        5. **Próximos Passos:** 3 ações claras e imediatas para melhorar a saúde financeira.

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

        ws_dash["A1"] = "📊 DASHBOARD FINANCEIRO & CONTROLE DE GASTOS (MÉDIA MENSAL)"
        ws_dash["A1"].font = Font(name="Calibri", size=16, bold=True, color=cor_primaria)

        # KPIs Gerais em Média Mensal (para controle financeiro realista e planejamento)
        num_meses = max(1, df['Mes_Ano'].nunique())
        receitas_total = float(df[df['Valor'] > 0]['Valor'].sum())
        despesas_total = float(df[df['Valor'] < 0]['Valor'].abs().sum())
        saldo_total = receitas_total - despesas_total

        receitas_media = receitas_total / num_meses
        despesas_media = despesas_total / num_meses
        saldo_medio = saldo_total / num_meses
        poupanca_taxa = ((receitas_total - despesas_total) / receitas_total) if receitas_total > 0 else 0.0

        # Cards de KPI com foco em Média Mensal
        ws_dash["A3"] = "Receita Média / Mês"
        ws_dash["A4"] = receitas_media
        ws_dash["B3"] = "Despesa Média / Mês"
        ws_dash["B4"] = despesas_media
        ws_dash["C3"] = "Sobra Média / Mês"
        ws_dash["C4"] = saldo_medio
        ws_dash["D3"] = "Taxa Média de Poupança"
        ws_dash["D4"] = poupanca_taxa

        for col_l in ["A", "B", "C"]:
            ws_dash[f"{col_l}3"].font = Font(name="Calibri", size=10, bold=True, color="595959")
            ws_dash[f"{col_l}3"].alignment = Alignment(horizontal="center")
            ws_dash[f"{col_l}3"].fill = PatternFill(start_color=cor_cinza, fill_type="solid")
            ws_dash[f"{col_l}4"].font = Font(name="Calibri", size=13, bold=True, color=cor_primaria)
            ws_dash[f"{col_l}4"].number_format = "R$ #,##0.00"
            ws_dash[f"{col_l}4"].alignment = Alignment(horizontal="center")
            ws_dash[f"{col_l}4"].fill = PatternFill(start_color=cor_cinza, fill_type="solid")

        ws_dash["D3"].font = Font(name="Calibri", size=10, bold=True, color="595959")
        ws_dash["D3"].alignment = Alignment(horizontal="center")
        ws_dash["D3"].fill = PatternFill(start_color=cor_cinza, fill_type="solid")
        ws_dash["D4"].font = Font(name="Calibri", size=13, bold=True, color=cor_primaria)
        ws_dash["D4"].number_format = "0.0%"
        ws_dash["D4"].alignment = Alignment(horizontal="center")
        ws_dash["D4"].fill = PatternFill(start_color=cor_cinza, fill_type="solid")

        # Texto informativo do período consolidado
        periodo_txt = f"ℹ️ Análise Financeira Mensal calculada sobre {num_meses} mês(es) de histórico. Total consolidado no período: Receitas R$ {receitas_total:,.2f} | Despesas R$ {despesas_total:,.2f} | Saldo Acumulado R$ {saldo_total:,.2f}."
        ws_dash["A5"] = periodo_txt
        ws_dash["A5"].font = Font(name="Calibri", size=9, italic=True, color="595959")

        # Tabela 1: Despesas Médias por Categoria para alimentar os gráficos
        despesas_df = df[df['Valor'] < 0].copy()
        despesas_df['Valor_Absoluto'] = despesas_df['Valor'].abs()
        resumo_cat = despesas_df.groupby('Categoria')['Valor_Absoluto'].sum().reset_index()
        resumo_cat['Media_Mensal'] = resumo_cat['Valor_Absoluto'] / num_meses
        resumo_cat['Pct_Total'] = resumo_cat['Valor_Absoluto'] / (despesas_total if despesas_total > 0 else 1.0)
        resumo_cat = resumo_cat.sort_values(by='Media_Mensal', ascending=False).reset_index(drop=True)

        ws_dash["A7"] = "Categoria"
        ws_dash["B7"] = "Média Mensal (R$)"
        ws_dash["C7"] = "Total Período (R$)"
        ws_dash["D7"] = "% dos Gastos"
        for col_l in ["A", "B", "C", "D"]:
            ws_dash[f"{col_l}7"].font = fonte_cabecalho
            ws_dash[f"{col_l}7"].fill = fill_cabecalho

        linha_atual = 8
        for _, r in resumo_cat.iterrows():
            ws_dash.cell(row=linha_atual, column=1, value=str(r['Categoria']))
            c_med = ws_dash.cell(row=linha_atual, column=2, value=float(r['Media_Mensal']))
            c_med.number_format = "R$ #,##0.00"
            c_tot = ws_dash.cell(row=linha_atual, column=3, value=float(r['Valor_Absoluto']))
            c_tot.number_format = "R$ #,##0.00"
            c_pct = ws_dash.cell(row=linha_atual, column=4, value=float(r['Pct_Total']))
            c_pct.number_format = "0.0%"
            linha_atual += 1

        total_cat_row = linha_atual
        if len(resumo_cat) > 0:
            ws_dash.cell(row=total_cat_row, column=1, value="TOTAL DESPESAS").font = Font(bold=True)
            c_total_med = ws_dash.cell(row=total_cat_row, column=2, value=f"=SUM(B8:B{total_cat_row-1})")
            c_total_med.font = Font(bold=True)
            c_total_med.number_format = "R$ #,##0.00"

            c_total_per = ws_dash.cell(row=total_cat_row, column=3, value=f"=SUM(C8:C{total_cat_row-1})")
            c_total_per.font = Font(bold=True)
            c_total_per.number_format = "R$ #,##0.00"

            c_total_pct = ws_dash.cell(row=total_cat_row, column=4, value=f"=SUM(D8:D{total_cat_row-1})")
            c_total_pct.font = Font(bold=True)
            c_total_pct.number_format = "0.0%"

            # Gráfico de Barras: Maiores Despesas Médias Mensais (aponta para Col B = Média Mensal)
            chart_bar = BarChart()
            chart_bar.type = "col"
            chart_bar.style = 10
            chart_bar.title = "Maiores Despesas por Categoria (Média Mensal R$)"
            chart_bar.y_axis.title = "Média Mensal (R$)"
            chart_bar.x_axis.title = "Categoria"
            chart_bar.width = 17
            chart_bar.height = 11
            chart_bar.legend = None

            data_ref = Reference(ws_dash, min_col=2, min_row=7, max_row=total_cat_row-1)
            cats_ref = Reference(ws_dash, min_col=1, min_row=8, max_row=total_cat_row-1)
            chart_bar.add_data(data_ref, titles_from_data=True)
            chart_bar.set_categories(cats_ref)
            ws_dash.add_chart(chart_bar, "F3")

            # Gráfico de Pizza: Distribuição Percentual de Gastos
            chart_pie = PieChart()
            chart_pie.title = "Distribuição Percentual de Gastos"
            chart_pie.width = 15
            chart_pie.height = 11
            chart_pie.add_data(data_ref, titles_from_data=True)
            chart_pie.set_categories(cats_ref)
            ws_dash.add_chart(chart_pie, "O3")

        proxima_linha = total_cat_row + 3

        # Tabela 2 e Gráfico: Comparativo por Banco / Conta (se houver mais de 1 conta)
        if df['Banco'].nunique() > 1:
            resumo_banco = df.groupby(['Banco', 'Tipo'])['Valor'].sum().unstack(fill_value=0).reset_index()
            if 'Receita' not in resumo_banco: resumo_banco['Receita'] = 0.0
            if 'Despesa' not in resumo_banco: resumo_banco['Despesa'] = 0.0
            resumo_banco['Despesa_Total'] = resumo_banco['Despesa'].abs()
            resumo_banco['Receita_Total'] = resumo_banco['Receita']
            resumo_banco['Receita_Media'] = resumo_banco['Receita_Total'] / num_meses
            resumo_banco['Despesa_Media'] = resumo_banco['Despesa_Total'] / num_meses

            ws_dash.cell(row=proxima_linha, column=1, value="Banco / Conta").font = fonte_cabecalho
            ws_dash.cell(row=proxima_linha, column=1).fill = fill_cabecalho
            ws_dash.cell(row=proxima_linha, column=2, value="Receita Média/Mês").font = fonte_cabecalho
            ws_dash.cell(row=proxima_linha, column=2).fill = fill_cabecalho
            ws_dash.cell(row=proxima_linha, column=3, value="Despesa Média/Mês").font = fonte_cabecalho
            ws_dash.cell(row=proxima_linha, column=3).fill = fill_cabecalho
            ws_dash.cell(row=proxima_linha, column=4, value="Saldo Total").font = fonte_cabecalho
            ws_dash.cell(row=proxima_linha, column=4).fill = fill_cabecalho

            curr_b_row = proxima_linha + 1
            for _, br in resumo_banco.iterrows():
                ws_dash.cell(row=curr_b_row, column=1, value=str(br['Banco']))
                c_rec_m = ws_dash.cell(row=curr_b_row, column=2, value=float(br.get('Receita_Media', 0)))
                c_rec_m.number_format = "R$ #,##0.00"
                c_desp_m = ws_dash.cell(row=curr_b_row, column=3, value=float(br.get('Despesa_Media', 0)))
                c_desp_m.number_format = "R$ #,##0.00"
                c_tot_b = ws_dash.cell(row=curr_b_row, column=4, value=float(br.get('Receita_Total', 0) - br.get('Despesa_Total', 0)))
                c_tot_b.number_format = "R$ #,##0.00"
                curr_b_row += 1

            chart_banco = BarChart()
            chart_banco.type = "col"
            chart_banco.style = 12
            chart_banco.title = "Cruzamento de Contas: Médias Mensais por Banco (R$)"
            chart_banco.y_axis.title = "Média Mensal (R$)"
            chart_banco.x_axis.title = "Banco / Conta"
            chart_banco.width = 17
            chart_banco.height = 11

            b_data = Reference(ws_dash, min_col=2, max_col=3, min_row=proxima_linha, max_row=curr_b_row-1)
            b_cats = Reference(ws_dash, min_col=1, min_row=proxima_linha+1, max_row=curr_b_row-1)
            chart_banco.add_data(b_data, titles_from_data=True)
            chart_banco.set_categories(b_cats)
            ws_dash.add_chart(chart_banco, f"F{proxima_linha}")

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
        ws_dash.column_dimensions["B"].width = 22
        ws_dash.column_dimensions["C"].width = 22
        ws_dash.column_dimensions["D"].width = 20

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
        resumo_geral['Valor_Absoluto'] = resumo_geral['Valor'].abs()
        resumo_geral['Media_Mensal'] = resumo_geral['Valor_Absoluto'] / num_meses

        ws_resumo.cell(row=1, column=1, value="Categoria").font = fonte_cabecalho
        ws_resumo.cell(row=1, column=1).fill = fill_cabecalho
        ws_resumo.cell(row=1, column=2, value="Tipo").font = fonte_cabecalho
        ws_resumo.cell(row=1, column=2).fill = fill_cabecalho
        ws_resumo.cell(row=1, column=3, value="Média Mensal (R$)").font = fonte_cabecalho
        ws_resumo.cell(row=1, column=3).fill = fill_cabecalho
        ws_resumo.cell(row=1, column=4, value="Total Período (R$)").font = fonte_cabecalho
        ws_resumo.cell(row=1, column=4).fill = fill_cabecalho

        for r_idx, (_, r) in enumerate(resumo_geral.iterrows(), 2):
            ws_resumo.cell(row=r_idx, column=1, value=str(r['Categoria']))
            ws_resumo.cell(row=r_idx, column=2, value=str(r['Tipo']))
            c_med = ws_resumo.cell(row=r_idx, column=3, value=float(r['Media_Mensal']))
            c_med.number_format = "R$ #,##0.00"
            c_tot = ws_resumo.cell(row=r_idx, column=4, value=float(r['Valor_Absoluto']))
            c_tot.number_format = "R$ #,##0.00"

        ws_resumo.column_dimensions["A"].width = 24
        ws_resumo.column_dimensions["B"].width = 16
        ws_resumo.column_dimensions["C"].width = 22
        ws_resumo.column_dimensions["D"].width = 22

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
        self.root.title("FinPy • Analisador Financeiro Multibancos (AI & Data Science)")
        self.root.geometry("920x880")
        self.root.minsize(780, 580)

        self.chave_groq = os.getenv("GROQ_API_KEY", "").strip()
        self.arquivos_selecionados = []
        self.mostrar_chave = False
        self.tela_cheia = False
        self.maximizado = False

        # Referências dos componentes dinâmicos
        self.kpi_labels = {}
        self.fig_matplotlib = None
        self.canvas_matplotlib = None

        # Container Principal Scrollável (com patch contra erros no Linux)
        self.main_container = ctk.CTkScrollableFrame(
            root,
            corner_radius=0,
            fg_color="transparent"
        )
        self.main_container.pack(fill="both", expand=True, padx=10, pady=8)

        # 1. Header com Logo, Título, Controles de Tela Cheia e Tema
        self._criar_header()

        # 2. Cards de KPI no Topo (Entradas, Saídas, Poupança %, Maior Ralo)
        self._criar_kpi_cards()

        # 3. Card de Configuração da API Key e Meta de Economia (Knapsack)
        self._criar_card_api_e_meta()

        # 4. Zona de Upload Drag & Drop
        self._criar_zona_upload()

        # 5. Lista de Arquivos Carregados
        self._criar_lista_arquivos()

        # 6. Barra de Ação & Progresso
        self._criar_barra_execucao()
        self._atualizar_lista_arquivos_ui()

        # 7. Tabview para Gráficos Embutidos, Relatório IA, Logs e Planilha
        self._criar_tabview_resultados()

        # Configura atalhos de teclado e scroll universal
        self._configurar_eventos_globais()

    def _criar_header(self):
        header = ctk.CTkFrame(self.main_container, fg_color="transparent")
        header.pack(fill="x", pady=(0, 8))

        title_box = ctk.CTkFrame(header, fg_color="transparent")
        title_box.pack(side="left")

        lbl_title = ctk.CTkLabel(
            title_box,
            text="📊 FinPy • Analisador Multibancos",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        lbl_title.pack(anchor="w")

        lbl_sub = ctk.CTkLabel(
            title_box,
            text="Pipeline de Extratos • Otimizador de Metas (Knapsack) • CDI • Groq AI",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        lbl_sub.pack(anchor="w")

        ctrl_box = ctk.CTkFrame(header, fg_color="transparent")
        ctrl_box.pack(side="right")

        self.lbl_status = ctk.CTkLabel(
            ctrl_box,
            text="● Sistema Pronto",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#2EA043"
        )
        self.lbl_status.pack(side="top", anchor="e", pady=(0, 2))

        botoes_box = ctk.CTkFrame(ctrl_box, fg_color="transparent")
        botoes_box.pack(side="top", anchor="e")

        self.btn_maximizar = ctk.CTkButton(
            botoes_box,
            text="🗖 Maximizar",
            width=90,
            height=26,
            font=ctk.CTkFont(size=11),
            fg_color=("gray70", "gray25"),
            hover_color=("gray60", "gray35"),
            command=self.alternar_maximizar
        )
        self.btn_maximizar.pack(side="left", padx=(0, 4))

        self.btn_tela_cheia = ctk.CTkButton(
            botoes_box,
            text="⛶ Tela Cheia",
            width=95,
            height=26,
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self.alternar_tela_cheia
        )
        self.btn_tela_cheia.pack(side="left", padx=(0, 4))

        theme_menu = ctk.CTkOptionMenu(
            botoes_box,
            values=["Dark", "Light", "System"],
            command=ctk.set_appearance_mode,
            width=85,
            height=26,
            font=ctk.CTkFont(size=11)
        )
        theme_menu.pack(side="left")

    def _criar_kpi_cards(self):
        """Cards de KPI Dinâmicos no Topo com visual executivo e moderno."""
        kpi_container = ctk.CTkFrame(self.main_container, fg_color="transparent")
        kpi_container.pack(fill="x", pady=(0, 10))

        # Configura grid 4 colunas com peso igual
        for i in range(4):
            kpi_container.grid_columnconfigure(i, weight=1)

        cards_data = [
            ("entradas", "💰 Receita Média/Mês", "R$ --", "#2EA043", "Média mensal de receitas"),
            ("saidas", "💸 Despesa Média/Mês", "R$ --", "#F85149", "Média mensal de despesas"),
            ("poupanca", "📈 Taxa Poupança", "--%", "#388BFD", "Sobra líquida / Entradas"),
            ("ralo", "🚨 Maior Ralo Mensal", "--", "#E5A93C", "Gargalo médio mensal")
        ]

        for col, (chave, titulo, val_padrao, cor_val, sub) in enumerate(cards_data):
            card = ctk.CTkFrame(kpi_container, corner_radius=10, fg_color=("gray90", "gray17"))
            card.grid(row=0, column=col, padx=4, sticky="nsew", ipady=4)

            lbl_t = ctk.CTkLabel(card, text=titulo, font=ctk.CTkFont(size=11, weight="bold"), text_color="gray")
            lbl_t.pack(anchor="w", padx=10, pady=(6, 1))

            lbl_v = ctk.CTkLabel(card, text=val_padrao, font=ctk.CTkFont(size=15, weight="bold"), text_color=cor_val)
            lbl_v.pack(anchor="w", padx=10)
            self.kpi_labels[chave] = lbl_v

            lbl_s = ctk.CTkLabel(card, text=sub, font=ctk.CTkFont(size=10), text_color="gray")
            lbl_s.pack(anchor="w", padx=10, pady=(0, 6))

    def _criar_card_api_e_meta(self):
        """Card combinado de Groq API Key e Meta de Economia (Knapsack)."""
        card = ctk.CTkFrame(self.main_container, corner_radius=10, fg_color=("gray90", "gray17"))
        card.pack(fill="x", pady=(0, 10), padx=2, ipady=4)

        # Seção API Key
        lbl_key = ctk.CTkLabel(card, text="🔑 Groq API Key:", font=ctk.CTkFont(size=12, weight="bold"))
        lbl_key.pack(side="left", padx=(12, 6))

        self.entry_chave = ctk.CTkEntry(
            card,
            placeholder_text="gsk_... (opcional se já no .env)",
            show="*",
            height=30
        )
        self.entry_chave.pack(side="left", fill="x", expand=True, padx=(0, 4))
        if self.chave_groq:
            self.entry_chave.insert(0, self.chave_groq)

        self.btn_eye = ctk.CTkButton(
            card,
            text="👁️",
            width=30,
            height=30,
            fg_color=("gray75", "gray25"),
            hover_color=("gray65", "gray35"),
            command=self._alternar_visibilidade_chave
        )
        self.btn_eye.pack(side="left", padx=(0, 10))

        # Seção Meta de Economia (Otimizador Knapsack)
        lbl_meta = ctk.CTkLabel(card, text="🎯 Meta Economia (R$):", font=ctk.CTkFont(size=12, weight="bold"))
        lbl_meta.pack(side="left", padx=(4, 6))

        self.entry_meta = ctk.CTkEntry(
            card,
            placeholder_text="Ex: 450.00",
            width=90,
            height=30
        )
        self.entry_meta.insert(0, "450.00")
        self.entry_meta.pack(side="left", padx=(0, 12))

    def _alternar_visibilidade_chave(self):
        self.mostrar_chave = not self.mostrar_chave
        self.entry_chave.configure(show="" if self.mostrar_chave else "*")

    def _criar_zona_upload(self):
        self.drop_frame = ctk.CTkFrame(
            self.main_container,
            corner_radius=12,
            border_width=2,
            border_color="#1f538d",
            fg_color=("gray95", "gray15")
        )
        self.drop_frame.pack(fill="x", pady=(0, 10), padx=2, ipady=12)

        lbl_icon = ctk.CTkLabel(self.drop_frame, text="📥", font=ctk.CTkFont(size=34))
        lbl_icon.pack()

        lbl_drop = ctk.CTkLabel(
            self.drop_frame,
            text="Arraste e solte seus extratos aqui (CSV ou PDF com/sem senha)",
            font=ctk.CTkFont(size=15, weight="bold")
        )
        lbl_drop.pack(pady=(2, 2))

        lbl_sub = ctk.CTkLabel(
            self.drop_frame,
            text="Suporta múltiplos extratos: Nubank, Inter, C6, Itaú, Bradesco, BB, Caixa, etc. Senha solicitada em memória.",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        lbl_sub.pack(pady=(0, 8))

        btn_browse = ctk.CTkButton(
            self.drop_frame,
            text="📂 Selecionar Arquivos...",
            command=self.selecionar_arquivos,
            width=200,
            height=32,
            corner_radius=8,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        btn_browse.pack()

        if TEM_DND:
            try:
                self.drop_frame.drop_target_register(DND_FILES)
                self.drop_frame.dnd_bind('<<Drop>>', self.ao_soltar_arquivos)
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind('<<Drop>>', self.ao_soltar_arquivos)
            except Exception:
                pass

    def _criar_lista_arquivos(self):
        bar_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        bar_frame.pack(fill="x", pady=(0, 4))

        self.lbl_contador = ctk.CTkLabel(
            bar_frame,
            text="📁 Extratos para Cruzamento (0)",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.lbl_contador.pack(side="left")

        btn_clear = ctk.CTkButton(
            bar_frame,
            text="🧹 Limpar Todos",
            command=self.limpar_arquivos,
            width=100,
            height=26,
            fg_color=("gray70", "gray25"),
            hover_color=("gray60", "gray35"),
            font=ctk.CTkFont(size=11)
        )
        btn_clear.pack(side="right")

        self.container_arquivos = ctk.CTkFrame(
            self.main_container,
            corner_radius=8,
            fg_color=("gray90", "gray13")
        )
        self.container_arquivos.pack(fill="x", pady=(0, 10))

    def _atualizar_lista_arquivos_ui(self):
        for widget in self.container_arquivos.winfo_children():
            widget.destroy()

        if not self.arquivos_selecionados:
            lbl_vazio = ctk.CTkLabel(
                self.container_arquivos,
                text="Nenhum extrato carregado. Arraste arquivos na caixa acima ou clique em 'Selecionar Arquivos'.",
                font=ctk.CTkFont(size=12),
                text_color="gray"
            )
            lbl_vazio.pack(pady=16)
            self.lbl_contador.configure(text="📁 Extratos para Cruzamento (0)", text_color="gray")
            self.btn_executar.configure(state="disabled")
            return

        for caminho in self.arquivos_selecionados:
            card = ctk.CTkFrame(self.container_arquivos, corner_radius=6, fg_color=("gray85", "gray18"))
            card.pack(fill="x", pady=2, padx=4)

            is_pdf = caminho.lower().endswith('.pdf')
            badge_txt = "PDF" if is_pdf else "CSV"
            badge_color = "#6A1B9A" if is_pdf else "#0D47A1"

            lbl_badge = ctk.CTkLabel(
                card,
                text=badge_txt,
                fg_color=badge_color,
                text_color="white",
                corner_radius=4,
                width=38,
                font=ctk.CTkFont(size=10, weight="bold")
            )
            lbl_badge.pack(side="left", padx=(6, 8), pady=4)

            nome_arq = os.path.basename(caminho)
            lbl_nome = ctk.CTkLabel(
                card,
                text=nome_arq,
                font=ctk.CTkFont(size=12, weight="bold")
            )
            lbl_nome.pack(side="left", padx=(0, 8))

            try:
                tam_kb = os.path.getsize(caminho) / 1024
                tam_str = f"{tam_kb:.1f} KB" if tam_kb < 1024 else f"{tam_kb/1024:.2f} MB"
            except Exception:
                tam_str = ""

            lbl_tam = ctk.CTkLabel(card, text=tam_str, font=ctk.CTkFont(size=11), text_color="gray")
            lbl_tam.pack(side="left")

            btn_del = ctk.CTkButton(
                card,
                text="✕",
                width=24,
                height=22,
                corner_radius=4,
                fg_color="transparent",
                hover_color="#C62828",
                text_color=("gray20", "gray80"),
                command=lambda p=caminho: self.remover_arquivo(p)
            )
            btn_del.pack(side="right", padx=6)

        qtd = len(self.arquivos_selecionados)
        self.lbl_contador.configure(text=f"📁 Extratos para Cruzamento ({qtd})", text_color="#2EA043")
        self.btn_executar.configure(state="normal")

    def _criar_barra_execucao(self):
        self.btn_executar = ctk.CTkButton(
            self.main_container,
            text="🚀 EXECUTAR CRUZAMENTO & ANÁLISE DE DADOS",
            height=44,
            corner_radius=8,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#2EA043",
            hover_color="#238636",
            command=self.iniciar_processamento,
            state="disabled"
        )
        self.btn_executar.pack(fill="x", pady=(0, 4))

        self.progress_bar = ctk.CTkProgressBar(self.main_container, mode="indeterminate", height=4)
        self.progress_bar.pack(fill="x", pady=(0, 8))
        self.progress_bar.set(0)

    def _criar_tabview_resultados(self):
        self.tabview = ctk.CTkTabview(self.main_container, height=360, corner_radius=10)
        self.tabview.pack(fill="both", expand=True, pady=(0, 10))

        tab_graficos = self.tabview.add("📊 Gráficos Embutidos")
        tab_ia = self.tabview.add("💡 Relatório da IA")
        tab_log = self.tabview.add("📋 Logs do Pipeline")
        tab_planilha = self.tabview.add("📁 Planilha Gerada")

        # Tab 1: Gráficos Embutidos (Matplotlib + CTk)
        self.frame_graficos = ctk.CTkFrame(tab_graficos, fg_color="transparent")
        self.frame_graficos.pack(fill="both", expand=True, padx=4, pady=4)

        self.lbl_graficos_placeholder = ctk.CTkLabel(
            self.frame_graficos,
            text="Os gráficos interativos (Rosca de Gastos, Fluxo de Caixa e Juros Compostos CDI)\nserão renderizados aqui após a execução do pipeline.",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        self.lbl_graficos_placeholder.pack(expand=True, pady=40)

        # Tab 2: Relatório IA
        self.txt_relatorio = ctk.CTkTextbox(
            tab_ia,
            height=280,
            font=ctk.CTkFont(size=12),
            corner_radius=6
        )
        self.txt_relatorio.pack(fill="both", expand=True, padx=4, pady=4)

        # Tab 3: Logs
        self.txt_log = ctk.CTkTextbox(
            tab_log,
            height=280,
            font=ctk.CTkFont(family="Consolas", size=11),
            text_color="#80FF80",
            fg_color=("#1a1a1a", "#0f0f0f"),
            corner_radius=6
        )
        self.txt_log.pack(fill="both", expand=True, padx=4, pady=4)

        # Tab 4: Arquivo Gerado & Ações
        self.frame_planilha = ctk.CTkFrame(tab_planilha, corner_radius=8, fg_color="transparent")
        self.frame_planilha.pack(fill="both", expand=True, padx=12, pady=12)

        lbl_icon_p = ctk.CTkLabel(self.frame_planilha, text="📈", font=ctk.CTkFont(size=36))
        lbl_icon_p.pack(pady=(6, 2))

        self.lbl_p_titulo = ctk.CTkLabel(
            self.frame_planilha,
            text="Relatorio_Financeiro_Final.xlsx",
            font=ctk.CTkFont(size=15, weight="bold")
        )
        self.lbl_p_titulo.pack()

        self.lbl_p_desc = ctk.CTkLabel(
            self.frame_planilha,
            text="Planilha consolidada com Dashboards visuais, gráficos interativos e cruzamento de bancos.\nCompatível nativamente com Google Sheets e Microsoft Excel.",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        self.lbl_p_desc.pack(pady=(2, 10))

        self.btn_abrir_planilha = ctk.CTkButton(
            self.frame_planilha,
            text="📂 Abrir Planilha no Computador",
            command=self._abrir_planilha,
            width=230,
            height=34,
            corner_radius=8,
            font=ctk.CTkFont(size=12, weight="bold"),
            state="disabled"
        )
        self.btn_abrir_planilha.pack()

    def _atualizar_kpis(self, df):
        """Calcula e atualiza os valores dinâmicos dos cards de KPI no topo em formato de Média Mensal."""
        try:
            num_meses = max(1, df['Mes_Ano'].nunique())
            entradas_total = float(df[df['Valor'] > 0]['Valor'].sum())
            saidas_total = float(df[df['Valor'] < 0]['Valor'].abs().sum())

            entradas_media = entradas_total / num_meses
            saidas_media = saidas_total / num_meses
            poupanca_pct = ((entradas_total - saidas_total) / entradas_total * 100) if entradas_total > 0 else 0.0

            despesas = df[df['Valor'] < 0]
            if not despesas.empty:
                gastos_cat = despesas.groupby('Categoria')['Valor'].sum().abs()
                maior_cat = str(gastos_cat.idxmax())
                val_maior_media = float(gastos_cat.max()) / num_meses
                texto_ralo = f"{maior_cat} (R$ {val_maior_media:,.0f}/mês)"
            else:
                texto_ralo = "Nenhum"

            self.kpi_labels['entradas'].configure(text=f"R$ {entradas_media:,.2f}/mês")
            self.kpi_labels['saidas'].configure(text=f"R$ {saidas_media:,.2f}/mês")
            self.kpi_labels['poupanca'].configure(text=f"{poupanca_pct:.1f}%")
            self.kpi_labels['ralo'].configure(text=texto_ralo)
        except Exception as e:
            print(f"Aviso ao atualizar KPIs: {e}")

    def _atualizar_graficos(self, df, cenarios, gastos):
        """Renderiza os gráficos embutidos de Rosca, Fluxo Temporal e Juros Compostos no CTk."""
        try:
            if self.lbl_graficos_placeholder:
                self.lbl_graficos_placeholder.destroy()
                self.lbl_graficos_placeholder = None

            if self.canvas_matplotlib is not None:
                self.canvas_matplotlib.get_tk_widget().destroy()
                self.canvas_matplotlib = None

            # Cria figura estilizada em tema Dark
            fig = Figure(figsize=(9.2, 3.8), facecolor='#161b22', dpi=100)
            ax1 = fig.add_subplot(131)
            ax2 = fig.add_subplot(132)
            ax3 = fig.add_subplot(133)

            # 1. Gráfico de Rosca (Donut) das Despesas por Categoria
            ax1.set_facecolor('#161b22')
            despesas_cat = gastos.sort_values(ascending=False)
            if len(despesas_cat) > 5:
                top5 = despesas_cat.head(5)
                outras_soma = despesas_cat.iloc[5:].sum()
                valores = top5.tolist() + [outras_soma]
                labels = top5.index.tolist() + ['OUTRAS']
            else:
                valores = despesas_cat.tolist()
                labels = despesas_cat.index.tolist()

            cores_paleta = ['#388BFD', '#2EA043', '#FF7B72', '#D29922', '#A371F7', '#56D364']
            if valores and sum(valores) > 0:
                wedges, texts, autotexts = ax1.pie(
                    valores, labels=labels, autopct='%1.0f%%', startangle=140,
                    colors=cores_paleta[:len(valores)],
                    wedgeprops=dict(width=0.45, edgecolor='#161b22', linewidth=2),
                    textprops=dict(color='#C9D1D9', fontsize=8)
                )
                for at in autotexts:
                    at.set_color('white')
                    at.set_fontsize(8)
            ax1.set_title('Despesas por Categoria', color='white', fontsize=10, fontweight='bold')

            # 2. Fluxo de Caixa Mensal (Temporal)
            ax2.set_facecolor('#161b22')
            df['Mes_Ano_Str'] = df['Data'].dt.strftime('%b/%y')
            meses_unicos = df['Mes_Ano_Str'].unique().tolist()
            entradas_mes = []
            saidas_mes = []
            for m in meses_unicos:
                df_m = df[df['Mes_Ano_Str'] == m]
                entradas_mes.append(float(df_m[df_m['Valor'] > 0]['Valor'].sum()))
                saidas_mes.append(float(df_m[df_m['Valor'] < 0]['Valor'].abs().sum()))

            x_pos = np.arange(len(meses_unicos))
            largura = 0.35
            ax2.bar(x_pos - largura/2, entradas_mes, largura, label='Entradas', color='#2EA043', alpha=0.85)
            ax2.bar(x_pos + largura/2, saidas_mes, largura, label='Saídas', color='#F85149', alpha=0.85)
            ax2.set_xticks(x_pos)
            ax2.set_xticklabels(meses_unicos, color='#8B949E', fontsize=8)
            ax2.tick_params(colors='#8B949E', labelsize=8)
            ax2.set_title('Fluxo de Caixa Mensal', color='white', fontsize=10, fontweight='bold')
            ax2.legend(facecolor='#21262d', edgecolor='#30363d', labelcolor='#C9D1D9', fontsize=7)
            ax2.grid(True, linestyle='--', alpha=0.15, color='#8B949E')

            # 3. Projeção de Juros Compostos (CDI) nos horizontes 1, 5 e 10 anos
            ax3.set_facecolor('#161b22')
            cenario_top = cenarios[0] if cenarios else {}
            proj = cenario_top.get('projecoes', {})
            prazos = ['1 Ano', '5 Anos', '10 Anos']
            if 1 in proj and 5 in proj and 10 in proj:
                investido = [proj[1]['investido'], proj[5]['investido'], proj[10]['investido']]
                montante = [proj[1]['montante'], proj[5]['montante'], proj[10]['montante']]
            else:
                investido = [5400, 27000, 54000]
                montante = [5650, 35000, 93000]

            xp = np.arange(len(prazos))
            ax3.bar(xp - largura/2, investido, largura, label='Aportado', color='#388BFD', alpha=0.8)
            ax3.bar(xp + largura/2, montante, largura, label='Montante CDI', color='#2EA043', alpha=0.9)
            ax3.set_xticks(xp)
            ax3.set_xticklabels(prazos, color='#8B949E', fontsize=8)
            ax3.tick_params(colors='#8B949E', labelsize=8)
            eco_top = cenario_top.get('Economia_Mensal_Estimada', 0.0)
            ax3.set_title(f'Projeção CDI (R$ {eco_top:,.0f}/mês)', color='white', fontsize=10, fontweight='bold')
            ax3.legend(facecolor='#21262d', edgecolor='#30363d', labelcolor='#C9D1D9', fontsize=7)
            ax3.grid(True, linestyle='--', alpha=0.15, color='#8B949E')

            fig.tight_layout()

            self.canvas_matplotlib = FigureCanvasTkAgg(fig, master=self.frame_graficos)
            self.canvas_matplotlib.draw()
            self.canvas_matplotlib.get_tk_widget().pack(fill="both", expand=True)
            self.fig_matplotlib = fig

        except Exception as e:
            print(f"Erro ao atualizar gráficos embutidos: {e}")

    def solicitar_senha_pdf(self, nome_arquivo):
        """
        Pop-up seguro e thread-safe na interface solicitando a senha para PDF criptografado.
        Descriptografa puramente em memória sem salvar intermediários no disco.
        """
        evento = threading.Event()
        resultado = {"senha": None}

        def _abrir_popup():
            dialog = ctk.CTkInputDialog(
                text=f"O extrato '{nome_arquivo}' está protegido por senha.\nDigite a chave para descriptografar em memória:",
                title="Extrato PDF Protegido"
            )
            resultado["senha"] = dialog.get_input()
            evento.set()

        self.root.after(0, _abrir_popup)
        evento.wait()
        return resultado["senha"]

    def obter_meta_economia(self):
        """Obtém a meta de economia definida pelo usuário ou fallback para R$ 450.00."""
        try:
            val_str = self.entry_meta.get().strip()
            return limpar_valor_monetario(val_str)
        except Exception:
            return 450.0

    def alternar_tela_cheia(self, event=None):
        self.tela_cheia = not self.tela_cheia
        try:
            self.root.attributes('-fullscreen', self.tela_cheia)
            if self.tela_cheia:
                self.btn_tela_cheia.configure(text="🗗 Janela")
            else:
                self.btn_tela_cheia.configure(text="⛶ Tela Cheia")
        except Exception:
            pass

    def sair_tela_cheia(self, event=None):
        if self.tela_cheia:
            self.tela_cheia = False
            try:
                self.root.attributes('-fullscreen', False)
                self.btn_tela_cheia.configure(text="⛶ Tela Cheia")
            except Exception:
                pass

    def alternar_maximizar(self, event=None):
        self.maximizado = not self.maximizado
        try:
            if platform.system() == 'Windows':
                self.root.state('zoomed' if self.maximizado else 'normal')
            else:
                self.root.attributes('-zoomed', self.maximizado)
            if self.maximizado:
                self.btn_maximizar.configure(text="🗗 Restaurar")
            else:
                self.btn_maximizar.configure(text="🗖 Maximizar")
        except Exception:
            pass

    def _configurar_eventos_globais(self):
        self.root.bind("<F11>", self.alternar_tela_cheia)
        self.root.bind("<Escape>", self.sair_tela_cheia)

        def _scroll_suave(event):
            widget_str = str(event.widget).lower()
            if 'text' in widget_str:
                return
            try:
                if platform.system() == 'Windows':
                    self.main_container._parent_canvas.yview_scroll(-int(event.delta / 60), "units")
                elif platform.system() == 'Darwin':
                    self.main_container._parent_canvas.yview_scroll(-int(event.delta), "units")
                else:
                    if event.num == 4:
                        self.main_container._parent_canvas.yview_scroll(-2, "units")
                    elif event.num == 5:
                        self.main_container._parent_canvas.yview_scroll(2, "units")
            except Exception:
                pass

        if platform.system() == 'Linux':
            self.root.bind_all("<Button-4>", _scroll_suave, add="+")
            self.root.bind_all("<Button-5>", _scroll_suave, add="+")
        else:
            self.root.bind_all("<MouseWheel>", _scroll_suave, add="+")

    def _abrir_planilha(self):
        caminho = os.path.abspath('Relatorio_Financeiro_Final.xlsx')
        if not os.path.exists(caminho):
            messagebox.showinfo("Aviso", "A planilha ainda não foi gerada. Execute a análise primeiro.")
            return

        try:
            if platform.system() == 'Windows':
                os.startfile(caminho)
            elif platform.system() == 'Darwin':
                subprocess.Popen(['open', caminho])
            else:
                subprocess.Popen(['xdg-open', caminho])
        except Exception as e:
            messagebox.showerror("Erro ao abrir", f"Não foi possível abrir o arquivo automaticamente:\n{e}")

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

        self._atualizar_lista_arquivos_ui()

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

    def remover_arquivo(self, caminho):
        if caminho in self.arquivos_selecionados:
            self.arquivos_selecionados.remove(caminho)
            self._atualizar_lista_arquivos_ui()

    def limpar_arquivos(self):
        self.arquivos_selecionados.clear()
        self._atualizar_lista_arquivos_ui()

    def log(self, mensagem):
        self.txt_log.insert("end", mensagem + "\n")
        self.txt_log.see("end")
        self.root.update_idletasks()

    def iniciar_processamento(self):
        chave_input = self.entry_chave.get().strip()
        if chave_input:
            self.chave_groq = chave_input

        if not self.chave_groq:
            messagebox.showwarning("Chave Groq Ausente", "Por favor, insira sua Chave da API do Groq no campo superior antes de continuar.")
            self.entry_chave.focus()
            return

        self.btn_executar.configure(state="disabled")
        self.btn_abrir_planilha.configure(state="disabled")
        self.txt_log.delete("1.0", "end")
        self.txt_relatorio.delete("1.0", "end")
        self.tabview.set("📋 Logs do Pipeline")
        self.progress_bar.start()
        self.lbl_status.configure(text="🟡 Processando...", text_color="#FFC107")

        thread = threading.Thread(target=self.processar_dados)
        thread.start()

    def processar_dados(self):
        try:
            total_arquivos = len(self.arquivos_selecionados)
            meta_informada = self.obter_meta_economia()
            self.log(f"Inicializando motor de IA e cruzamento de dados para {total_arquivos} extrato(s)...")
            self.log(f"Meta de Economia Mensal configurada: R$ {meta_informada:,.2f}")
            
            analisador = AnalisadorFinanceiro(api_key=self.chave_groq)

            self.log("[1/5] Lendo, descriptografando (se necessário) e cruzando dados...")
            df_consolidado = analisador.processar_arquivos(
                self.arquivos_selecionados,
                callback_senha=self.solicitar_senha_pdf
            )
            self.log(f"-> Total de transações após normalização: {len(df_consolidado)}")
            bancos_detectados = df_consolidado['Banco'].unique().tolist()
            self.log(f"-> Contas/Bancos detectados: {', '.join(bancos_detectados)}")

            if hasattr(analisador, 'info_deduplicacao'):
                d_info = analisador.info_deduplicacao
                self.log(f"-> Deduplicação inteligente: {d_info['duplicatas_removidas']} duplicatas eliminadas | {d_info['estornos_neutralizados']} estornos neutralizados.")

            self.log("[2/5] Rodando Otimizador de Metas (Knapsack) & Projeções de Juros Compostos (CDI)...")
            cenarios_combo, gastos_agrupados = analisador.analise_combinatoria_economia(meta_economia=meta_informada)

            self.log("[3/5] Consultando LLM via Groq API. Aguardando insights executivos...")
            relatorio = analisador.gerar_relatorio_ia(cenarios_combo, gastos_agrupados, bancos=bancos_detectados)

            self.log("[4/5] Gerando planilha compatível com Google Sheets com Gráficos e Dashboards...")
            analisador.exportar_resultados(df_consolidado, relatorio)

            self.log("[5/5] ✨ CONCLUÍDO COM SUCESSO! ✨")

            # Atualiza Aba de Relatório IA
            self.txt_relatorio.insert("end", relatorio)
            
            # Atualiza KPIs no topo e gráficos embutidos no thread principal
            self.root.after(0, lambda: self._atualizar_kpis(df_consolidado))
            self.root.after(0, lambda: self._atualizar_graficos(df_consolidado, cenarios_combo, gastos_agrupados))
            
            self.tabview.set("📊 Gráficos Embutidos")
            self.lbl_status.configure(text="🟢 Concluído", text_color="#2EA043")
            self.btn_abrir_planilha.configure(state="normal")

            messagebox.showinfo(
                "Sucesso",
                f"Cruzamento concluído com sucesso!\n\n"
                f"- Extratos processados: {total_arquivos}\n"
                f"- Contas/Bancos: {', '.join(bancos_detectados)}\n"
                f"- Total de transações: {len(df_consolidado)}\n"
                f"- Meta de economia mensal: R$ {meta_informada:,.2f}\n\n"
                f"Arquivo gerado: Relatorio_Financeiro_Final.xlsx\n"
                f"(Gráficos interativos renderizados na tela e prontos na planilha!)"
            )

        except Exception as e:
            self.log(f"\n❌ ERRO FATAL: {str(e)}")
            self.lbl_status.configure(text="🔴 Erro", text_color="#F44336")
            messagebox.showerror("Erro de Execução", f"Ocorreu um erro no pipeline de dados:\n\n{str(e)}")
        finally:
            self.progress_bar.stop()
            self.progress_bar.set(0)
            self.btn_executar.configure(state="normal")

if __name__ == "__main__":
    if TEM_DND:
        class AppWindow(ctk.CTk, TkinterDnD.DnDWrapper):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                self.TkdndVersion = TkinterDnD._require(self)
        janela_principal = AppWindow()
    else:
        janela_principal = ctk.CTk()

    app = AppGUI(janela_principal)
    janela_principal.mainloop()



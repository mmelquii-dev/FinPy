import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import pytest
import pandas as pd
from analisador_financeiro import (
    limpar_valor_monetario,
    classificar_por_regex,
    deduplicar_e_tratar_estornos,
    calcular_juros_compostos,
    otimizar_metas_combinatoria,
)


def test_limpar_valor_monetario():
    # Formato padrão brasileiro
    assert limpar_valor_monetario("R$ 1.250,50") == 1250.50
    assert limpar_valor_monetario("-450,90") == -450.90
    assert limpar_valor_monetario(" 35,00 ") == 35.00

    # Formato numérico ou com ponto decimal
    assert limpar_valor_monetario("1200.50") == 1200.50
    assert limpar_valor_monetario(300) == 300.0
    assert limpar_valor_monetario(-15.75) == -15.75

    # Valores nulos, vazios ou inválidos
    assert limpar_valor_monetario(None) == 0.0
    assert limpar_valor_monetario("") == 0.0
    assert limpar_valor_monetario("invalido") == 0.0


def test_classificador_regex():
    # Transporte
    assert classificar_por_regex("Uber *Trip 1234") == "TRANSPORTE"
    assert classificar_por_regex("99 TECNOLOGIA APP") == "TRANSPORTE"
    assert classificar_por_regex("POSTO IPIRANGA COMBUSTIVEL") == "TRANSPORTE"

    # Delivery & Restaurante
    assert classificar_por_regex("IFOOD *RESTAURANTE SAO PAULO") == "DELIVERY"
    assert classificar_por_regex("RAPPI BRASIL DELIVERY") == "DELIVERY"
    assert classificar_por_regex("RESTAURANTE OUTBACK STEAKHOUSE") == "RESTAURANTE"
    assert classificar_por_regex("BURGUER KING LANCHES") == "RESTAURANTE"

    # Mercado & Saúde
    assert classificar_por_regex("SUPERMERCADO CARREFOUR BA") == "MERCADO"
    assert classificar_por_regex("ATACADAO DISTRIBUICAO") == "MERCADO"
    assert classificar_por_regex("DROGASIL MORUMBI") == "SAÚDE"
    assert classificar_por_regex("FARMACIA PAGUE MENOS") == "SAÚDE"

    # Assinaturas & Compras
    assert classificar_por_regex("NETFLIX.COM MENSALIDADE") == "ASSINATURAS"
    assert classificar_por_regex("SPOTIFY BRASIL PREMIUM") == "ASSINATURAS"
    assert classificar_por_regex("AMAZON MARKETPLACE") == "COMPRAS"

    # Casos sem correspondência rápida (devem retornar None para acionar fallback Groq)
    assert classificar_por_regex("TRANSFERENCIA TED CLIENTE DESCONHECIDO") is None
    assert classificar_por_regex("LANCAMENTO DIVERSO 98765") is None


def test_deduplicar_e_tratar_estornos():
    dados = [
        {"Data": "2026-08-01", "Descricao": "IFOOD *LANCHE", "Valor": -50.0, "Banco": "Nubank"},
        {"Data": "2026-08-01", "Descricao": "IFOOD *LANCHE", "Valor": -50.0, "Banco": "Nubank"},  # Duplicata exata
        {"Data": "2026-08-02", "Descricao": "ESTORNO IFOOD *LANCHE", "Valor": 50.0, "Banco": "Nubank"},  # Estorno correspondente
        {"Data": "2026-08-05", "Descricao": "Salário Mensal XYZ", "Valor": 5000.0, "Banco": "Nubank"},
        {"Data": "2026-08-10", "Descricao": "Uber Viagem", "Valor": -35.0, "Banco": "Inter"},
    ]
    df = pd.DataFrame(dados)

    df_limpo, info = deduplicar_e_tratar_estornos(df)

    # 1 duplicata exata removida + 1 par de estorno neutralizado (-50 e +50 eliminados)
    assert info["duplicatas_removidas"] == 1
    assert info["estornos_neutralizados"] == 1
    assert len(df_limpo) == 2

    # Restam apenas o Salário e o Uber
    descricoes_restantes = df_limpo["Descricao"].tolist()
    assert "Salário Mensal XYZ" in descricoes_restantes
    assert "Uber Viagem" in descricoes_restantes


def test_otimizador_metas_knapsack():
    gastos = pd.Series({
        "ALUGUEL": 1500.0,        # Essencial (não deve ser cortado)
        "MERCADO": 600.0,         # Essencial (não deve ser cortado)
        "SAÚDE": 100.0,           # Essencial (não deve ser cortado)
        "DELIVERY": 350.0,        # Não essencial
        "RESTAURANTE": 280.0,     # Não essencial
        "ASSINATURAS": 120.0,     # Não essencial
        "TRANSPORTE": 160.0       # Não essencial
    })

    meta = 450.0
    cenarios, _ = otimizar_metas_combinatoria(gastos, meta_economia=meta)

    assert len(cenarios) > 0
    plano_otimizado = cenarios[0]

    # Verifica que atingiu ou superou a meta
    assert plano_otimizado["Economia_Mensal_Estimada"] >= meta

    # Verifica que nenhuma categoria essencial foi indicada para corte
    cortadas = plano_otimizado["Categorias_Cortadas"]
    for essencial in ["ALUGUEL", "MERCADO", "SAÚDE"]:
        assert essencial not in cortadas

    # Verifica se as projeções de juros compostos foram embutidas
    assert "projecoes" in plano_otimizado
    assert 1 in plano_otimizado["projecoes"]
    assert 5 in plano_otimizado["projecoes"]
    assert 10 in plano_otimizado["projecoes"]


def test_projecao_juros_compostos():
    pmt = 450.0
    taxa = 0.0085  # ~0.85% ao mês (100% CDI)

    # 1 ano (12 meses)
    inv1, mnt1, jur1 = calcular_juros_compostos(pmt, taxa, 12)
    assert inv1 == 450.0 * 12 == 5400.0
    assert mnt1 > inv1
    assert pytest.approx(jur1, 0.01) == mnt1 - inv1
    assert 5600.0 < mnt1 < 5800.0

    # 5 anos (60 meses)
    inv5, mnt5, jur5 = calcular_juros_compostos(pmt, taxa, 60)
    assert inv5 == 450.0 * 60 == 27000.0
    assert mnt5 > 34000.0
    assert jur5 > 7000.0

    # 10 anos (120 meses)
    inv10, mnt10, jur10 = calcular_juros_compostos(pmt, taxa, 120)
    assert inv10 == 450.0 * 120 == 54000.0
    assert mnt10 > 90000.0
    assert jur10 > 36000.0

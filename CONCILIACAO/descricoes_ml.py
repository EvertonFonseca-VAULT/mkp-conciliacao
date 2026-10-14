"""Guarda TODAS as linhas que o scraper encontrou no detalhe do pedido no
ML/MP (nao so os totais bucketizados em enriquecer.py/scraper_ml.py) --
inclusive as sublinhas de detalhamento (rotulo comeca com "-", ex. "- Tarifa
de cobranca" dentro do grupo "Frete") que o resto do pipeline ignora.

Serve pra pesquisar o que um valor e de verdade antes de decidir como
lancar no CMM (ex.: "e tarifa de cobranca dentro de Frete, vou lancar como
Cobranca") -- pedido do Everton, 15/09/2026.
"""

import pandas as pd

COLUNAS = ["PEDIDO_MKP", "ROTULO", "VALOR", "DETALHE"]


def gerar_descricoes_ml(ml_dados):
    """`ml_dados`: {PEDIDO_MKP: resumo_pedido_ou_None} vindo de
    scraper_ml.resumir_pedidos (cada resumo com "linhas_raw", a lista bruta
    (rotulo, valor) do detalhe do pedido)."""
    linhas = []
    for pedido, dados in (ml_dados or {}).items():
        if not dados or "erro" in dados:
            continue
        for rotulo, valor in dados.get("linhas_raw") or []:
            if not rotulo:
                continue
            rotulo = rotulo.strip()
            linhas.append({
                "PEDIDO_MKP": pedido,
                "ROTULO": rotulo,
                "VALOR": valor,
                "DETALHE": rotulo.startswith("-"),
            })
    return pd.DataFrame(linhas, columns=COLUNAS)


if __name__ == "__main__":
    ml_dados = {
        "123": {
            "linhas_raw": [
                ("Frete", 10.0),
                ("- Tarifa de cobranca", 2.5),
                ("Tarifa de venda total", -8.0),
            ],
        },
        "456": {"erro": "timeout"},
        "789": None,
    }
    df = gerar_descricoes_ml(ml_dados)
    assert len(df) == 3, df
    assert bool(df.iloc[1]["DETALHE"]) and df.iloc[1]["ROTULO"] == "- Tarifa de cobranca"
    assert not bool(df.iloc[0]["DETALHE"])
    print("descricoes_ml.py: ok")

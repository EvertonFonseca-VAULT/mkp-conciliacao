"""Mercado COLETAS ES — só o carregamento do AB muda em relação ao FULL
(aba 'Sheet0', layout de 60 colunas em vez de 40). O resto do cruzamento
com o TOTVS/SRI é o mesmo processo genérico, reaproveitado do FULL."""

import os
import sys

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "FULL"))
from processar_full import (  # noqa: F401 (re-exportado pra quem importar deste módulo)
    marcar_flags, filtrar_validos, carregar_totvs, cruzar_totvs,
    resolver_pendentes_via_sri, gerar_para_processar,
)

from tipagem_es import LAYOUT_ES, PT_PARA_EN


def carregar_ab(caminho):
    """O nome da aba varia entre exports ('Relatorio' na maioria, 'Sheet0'
    em pelo menos um lote de 01.09.2026) — tenta 'Relatorio' primeiro,
    senão usa a única aba do arquivo."""
    planilhas = pd.ExcelFile(caminho).sheet_names
    aba = "Relatorio" if "Relatorio" in planilhas else planilhas[0]
    df = pd.read_excel(caminho, sheet_name=aba)
    # Export em portugues (43 colunas, Power Query do ELETRO): traduz pro ingles.
    df = df.rename(columns=PT_PARA_EN)
    if set(df.columns) == set(PT_PARA_EN.values()):
        return df
    # Mesmas 60 colunas em outra ordem (export do Koncili) tambem vale: reordena.
    if len(df.columns) != len(LAYOUT_ES) or set(df.columns) != set(LAYOUT_ES):
        raise ValueError(f"Layout do Mercado COLETAS ES mudou. Esperado: {LAYOUT_ES}\nEncontrado: {list(df.columns)}")
    return df[LAYOUT_ES]

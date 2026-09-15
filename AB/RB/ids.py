"""Normaliza IDs de pedido (EXTERNAL_REFERENCE/PED MKP) para string, sem
sufixo ".0" -- mesmo problema ja visto no cruzamento TOTVS do FULL/ES
(coluna float64 por causa de NaN)."""

import pandas as pd


def limpar_id(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    if isinstance(valor, float):
        return str(int(valor)) if valor.is_integer() else str(valor)
    return str(valor).strip()

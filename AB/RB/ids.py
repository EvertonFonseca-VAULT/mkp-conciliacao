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


def limpar_pedido_ab(valor):
    """Como limpar_id, mas nunca devolve "" -- uma linha do AB sem
    EXTERNAL_REFERENCE vira "VAZIO", chave que nunca bate com o RB (que
    tambem tem lancamentos sem PED_MKP, mas de motivos completamente
    diferentes -- agregar os dois lados so por estarem igualmente vazios
    juntava coisas sem relacao nenhuma: achado real do FULL 18/09/2026,
    uma linha "payment" de -15,28 vs 15 lancamentos do RB somando
    -49.642,08, virando um DIF fantasma de R$49.626,80 na fila
    VERIFICAR). Pedido do Everton, 23/09/2026."""
    valor = limpar_id(valor)
    return "VAZIO" if valor == "" else valor

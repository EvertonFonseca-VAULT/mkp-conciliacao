"""Le historicos_totvs.json (gerado por extrair_historicos.py a partir da
aba HISTORICOS do CONTROLE_V2.xlsx) -- loja -> RAIZ do historico TOTVS +
codigos _REPASSE/_COMISSAO/_FRETE/_ESTORNO/_ANUNCIO, pra mostrar no
dashboard (ver dashboard.py) e colar direto no TOTVS ao resolver um
VERIFICAR."""

import json
import os

_CAMINHO_JSON = os.path.join(os.path.dirname(__file__), "historicos_totvs.json")

with open(_CAMINHO_JSON, encoding="utf-8") as _f:
    _POR_RAIZ = {item["raiz"]: item for item in json.load(_f)}


def buscar_historico(raiz):
    """None se essa raiz nao estiver na tabela HISTORICOS (ex. ELETRO, que
    usa uma raiz propria desse pipeline, fora da planilha)."""
    return _POR_RAIZ.get(raiz)


if __name__ == "__main__":
    h = buscar_historico("MKP_MELI_FU")
    assert h and h["repasse"] == "MKP_MELI_FU_REPASSE", h
    assert buscar_historico("XPTO_NAO_EXISTE") is None
    print("historicos.py: ok")

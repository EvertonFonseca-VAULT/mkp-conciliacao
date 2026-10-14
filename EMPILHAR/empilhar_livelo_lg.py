"""Empilhado do LIVELO LG -- ponto de entrada; implementacao e premissas em empilhar_repasses.py.

Fonte: MKP - LIVELO/LG/<ano>/<mes.ano>/LIVELO LG *.xlsx (arquivo mensal fechado, sem AB/TOTVS).
Alvo: ARQUIVOS REPASSES COM CHAVES PRIMÁRIAS/LG_LIVELO.xlsx (abas Base + Calendario).

Uso:
    python empilhar_livelo_lg.py      # empilha o que faltar (dados e/ou calendario), grava direto
    python empilhar_livelo_lg.py --teste
"""
import sys

from empilhar_repasses import main, teste

if __name__ == "__main__":
    if "--teste" in sys.argv:
        teste("LIVELO_LG")
    else:
        main("LIVELO_LG", sys.argv[1:])

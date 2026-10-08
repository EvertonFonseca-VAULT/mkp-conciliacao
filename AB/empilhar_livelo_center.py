"""Empilhado do LIVELO CENTER -- ponto de entrada; implementacao e premissas em empilhar_repasses.py.

Fonte: MKP - LIVELO/CENTER/<ano>/<mes.ano>/LIVELO CENTER *.xlsx (arquivo mensal fechado, sem AB/TOTVS).
Alvo: ARQUIVOS REPASSES COM CHAVES PRIMÁRIAS/CENTER_LIVELO.xlsx (abas EMPILHADO + CALENDARIO).

Uso:
    python empilhar_livelo_center.py      # empilha o que faltar (dados e/ou calendario), grava direto
    python empilhar_livelo_center.py --teste
"""
import sys

from empilhar_repasses import main, teste

if __name__ == "__main__":
    if "--teste" in sys.argv:
        teste("LIVELO_CENTER")
    else:
        main("LIVELO_CENTER", sys.argv[1:])

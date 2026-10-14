"""Empilhado do COLOMBO CENTER -- ponto de entrada; implementacao e premissas em empilhar_repasses.py.

Fonte: MKP - COLOMBO/CENTER/AB/<ano>/<mes.ano>/COLOMBO CENTER *.xlsx (exportacao diaria "AB", sem
resumo embutido -- a data do repasse vem do nome do arquivo, nao de dentro dele).
Alvo: ARQUIVOS REPASSES COM CHAVES PRIMÁRIAS/CENTER_COLOMBO.xlsx (abas CENTER + Calendário).

Uso:
    python empilhar_colombo_center.py      # empilha o que faltar (dados e/ou calendario), grava direto
    python empilhar_colombo_center.py --teste
"""
import sys

from empilhar_repasses import main, teste

if __name__ == "__main__":
    if "--teste" in sys.argv:
        teste("COLOMBO_CENTER")
    else:
        main("COLOMBO_CENTER", sys.argv[1:])

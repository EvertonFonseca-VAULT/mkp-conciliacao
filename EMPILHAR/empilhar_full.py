"""Empilhado do FULL -- ponto de entrada; implementacao e premissas em empilhar_canais.py.

Para Processar do dia: C:\\RODAR TOTVS\\PROCESSAR\\FULL\\PROCESSAR_MERCADO FULL DD.MM.AAAA.xlsx (com ano).

Uso:
    python empilhar_full.py 10.09.2026                       # monta e confere; NAO grava nada
    python empilhar_full.py 10.09.2026 11.09.2026 --publicar # idem e grava no empilhado do X:
    python empilhar_full.py --teste                          # teste offline
"""

import sys

from empilhar_canais import main, teste

if __name__ == "__main__":
    if "--teste" in sys.argv:
        teste("FULL")
    else:
        main("FULL", sys.argv[1:])

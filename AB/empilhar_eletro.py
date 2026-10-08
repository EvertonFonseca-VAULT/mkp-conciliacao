"""Empilhado do ELETRO -- ponto de entrada; implementacao e premissas em empilhar_canais.py.

Para Processar do dia: C:\\RODAR TOTVS\\PROCESSAR\\ELETRO\\PROCESSAR_MERCADO ELETRODOMESTICO DD.MM.AAAA.xlsx.
AB em ELETRODOMESTICO\\AAAA\\MM.AAAA\\MERCADO ELETRODOMESTICO DD.MM.AAAA.xlsx (koncili.py).

Uso:
    python empilhar_eletro.py 04.09.2026                        # monta e confere; NAO grava nada
    python empilhar_eletro.py 04.09.2026 11.09.2026 --publicar  # idem e grava no empilhado do X:
    python empilhar_eletro.py --teste                           # teste offline
"""

import sys

from empilhar_canais import main, teste

if __name__ == "__main__":
    if "--teste" in sys.argv:
        teste("ELETRO")
    else:
        main("ELETRO", sys.argv[1:])

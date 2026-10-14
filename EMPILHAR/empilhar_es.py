"""Empilhado do ES -- ponto de entrada; implementacao e premissas em empilhar_canais.py.

Para Processar do dia: C:\\RODAR TOTVS\\PROCESSAR\\COLETAS ES\\PROCESSAR_MERCADO COLETAS ES DD.MM.xlsx
(sem ano no nome -- rodar_canal.py grava assim pra ES; ver _carregar_processar_glob_sem_ano).

Uso:
    python empilhar_es.py 10.09.2026                       # monta e confere; NAO grava nada
    python empilhar_es.py 10.09.2026 11.09.2026 --publicar # idem e grava no empilhado do X:
    python empilhar_es.py --teste                          # teste offline
"""

import sys

from empilhar_canais import main, teste

if __name__ == "__main__":
    if "--teste" in sys.argv:
        teste("ES")
    else:
        main("ES", sys.argv[1:])

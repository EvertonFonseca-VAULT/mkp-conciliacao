"""Empilhado do WEB -- ponto de entrada; implementacao e premissas em empilhar_canais.py.

Para Processar do dia: C:\\RODAR TOTVS\\PROCESSAR\\WEB\\PROCESSAR_MERCADO WEB DD.MM.AAAA.xlsx.
VALOR (W) = SUM(I:K); SEQUENCIA (X) = X<anterior>+1; CHAVE_PRIMARIA (Y) = CONCATENATE(X,"_WEB_MERCADO").
Nome e flags preservados por compatibilidade com o app desktop (Mkp/app/worker.py).

Uso:
    python empilhar_web.py 10.09.2026                       # monta e confere; NAO grava nada
    python empilhar_web.py 10.09.2026 11.09.2026 --publicar # idem e grava no empilhado do X:
    python empilhar_web.py --teste                          # teste offline
"""

import sys

from empilhar_canais import main, teste

if __name__ == "__main__":
    if "--teste" in sys.argv:
        teste("WEB")
    else:
        main("WEB", sys.argv[1:])

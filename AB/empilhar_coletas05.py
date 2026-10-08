"""Empilhado do COLETAS_CO05 -- ponto de entrada; implementacao e premissas em empilhar_canais.py.

Para Processar do dia: C:\\RODAR TOTVS\\PROCESSAR\\COLETAS_CO05\\PROCESSAR_MERCADO COLETAS_CO05 DD.MM.AAAA.xlsx.
Mesmo layout 25 colunas do WEB. O que muda (lido do proprio arquivo, ver _convencoes): VALOR (W) =
SUM(I,J,L,K); SEQUENCIA (X) = numero fixo da linha (X2=2); CHAVE_PRIMARIA (Y) =
CONCATENATE(X,"_COLETAS_CO05_MERCADO"); INSTALLMENTS (O) gravado como texto. Nome e flags preservados
por compatibilidade com o app desktop (Mkp/app/worker.py).

Uso:
    python empilhar_coletas05.py 11.09.2026                        # monta e confere; NAO grava nada
    python empilhar_coletas05.py 11.09.2026 14.09.2026 --publicar  # idem e grava no empilhado do X:
    python empilhar_coletas05.py --teste                           # teste offline
"""

import sys

from empilhar_canais import main, teste

if __name__ == "__main__":
    if "--teste" in sys.argv:
        teste("COLETAS_CO05")
    else:
        main("COLETAS_CO05", sys.argv[1:])

"""Empilhado do MOVEIS -- ponto de entrada; implementacao e premissas em empilhar_canais.py.

Para Processar do dia: C:\\RODAR TOTVS\\PROCESSAR\\MOVEIS\\PROCESSAR_MERCADO MOVEIS DD.MM.AAAA.xlsx.
O que muda em relacao aos demais (lido do proprio arquivo, ver _convencoes): VALOR (W) = I+J++K, sem
SUM; SEQUENCIA (X) = numero fixo linha-1; CHAVE_PRIMARIA (Y) = CONCATENATE(X,"_MÓVEIS_MERCADO") (com
acento). AB vem com cabecalhos em PORTUGUES (traduzidos) e DATE gravada como texto com virgula
(AAAA-MM-DDTHH:MM:SS,000-03:00). INSTALLMENTS (O) gravado como texto. Nome e flags preservados por
compatibilidade com o app desktop (Mkp/app/worker.py).

Uso:
    python empilhar_moveis.py 04.09.2026                        # monta e confere; NAO grava nada
    python empilhar_moveis.py 04.09.2026 11.09.2026 --publicar  # idem e grava no empilhado do X:
    python empilhar_moveis.py --teste                           # teste offline
"""

import sys

from empilhar_canais import main, teste

if __name__ == "__main__":
    if "--teste" in sys.argv:
        teste("MOVEIS")
    else:
        main("MOVEIS", sys.argv[1:])

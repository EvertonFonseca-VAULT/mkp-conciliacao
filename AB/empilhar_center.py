"""Empilhado do CENTER (Mercado Pago) -- monta as linhas do dia e grava direto no .xlsx do X:.

Mesmas premissas do empilhar_web.py:
Nao carrega o TOTVS. O "Nº QUE ESTA NO TOTVS" (coluna U) vem do arquivo Para Processar do dia
(C:\\RODAR TOTVS\\PROCESSAR\\CENTER\\PROCESSAR_MERCADO CENTER DD.MM.AAAA.xlsx), cruzado com o AB por SOURCE_ID.
Como esse arquivo vem sem duplicados, as linhas do AB cujo SOURCE_ID nao esta nele usam o SHIPPING_ID / ORDER_ID /
EXTERNAL_REFERENCE da propria linha que exista no conjunto de numeros do Para Processar (se der 0 ou 2, para e avisa).
  * sem pasta de backup e sem copia local: le o .xlsx do X: e grava um ".novo" na mesma pasta;
    so no fim troca o original (os.replace, sem trafego). Se der erro antes, o original fica intacto.
  * VALOR (W), SEQUENCIA (X) e sufixo da CHAVE_PRIMARIA (Y) sao lidos do proprio arquivo (_convencoes),
    nao do codigo -- o script segue exatamente o que ja esta gravado no empilhado.
  * nao mexe em calendario/tabela a nao ser que existam no arquivo.
  * sem cache em pickle: monta, confere e grava na mesma execucao.

A implementacao fica em empilhar_canais.py (modulo compartilhado dos canais -- ver Canal em CONTEXT.md).
Este arquivo e so o ponto de entrada do canal CENTER; nome e flags preservados por compatibilidade
com o app desktop (Mkp/app/worker.py).

Uso (coloque este arquivo na mesma pasta do empilhar.py):
    python empilhar_center.py 10.09.2026                       # monta e confere; NAO grava nada
    python empilhar_center.py 10.09.2026 11.09.2026 --publicar # idem e grava no empilhado do X:
    python empilhar_center.py --teste                          # teste offline
Rodar de novo com um dia que ja esta no empilhado nao duplica nada (o dia e pulado).
"""

import sys

from empilhar_canais import main, teste

if __name__ == "__main__":
    if "--teste" in sys.argv:
        teste("CENTER")
    else:
        main("CENTER", sys.argv[1:])

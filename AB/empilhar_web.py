"""Empilhado do WEB (Mercado Pago) -- monta as linhas do dia e grava direto no .xlsx do X:.

Nao carrega o TOTVS. O "Nº QUE ESTA NO TOTVS" (coluna U) vem do arquivo Para Processar do dia
(C:\\RODAR TOTVS\\PROCESSAR\\WEB\\PROCESSAR_MERCADO WEB DD.MM.AAAA.xlsx), cruzado com o AB por SOURCE_ID.
Como esse arquivo vem sem duplicados, as linhas do AB cujo SOURCE_ID nao esta nele usam o SHIPPING_ID / ORDER_ID /
EXTERNAL_REFERENCE da propria linha que exista no conjunto de numeros do Para Processar (se der 0 ou 2, para e avisa).
O que muda em relacao ao FULL/ES:
  * sem pasta de backup e sem copia local: le o .xlsx do X: e grava um ".novo" na mesma pasta;
    so no fim troca o original (os.replace, sem trafego). Se der erro antes, o original fica intacto.
  * VALOR (coluna W) e a formula SUM(I:K); SEQUENCIA (X) = X<linha anterior>+1;
    CHAVE_PRIMARIA (Y) = CONCATENATE(X,"_WEB_MERCADO").
  * nao mexe em calendario/tabela a nao ser que existam no arquivo (o exemplo nao tem).
  * sem cache em pickle: monta, confere e grava na mesma execucao.

A implementacao fica em empilhar_canais.py (modulo compartilhado dos canais -- ver Canal em CONTEXT.md).
Este arquivo e so o ponto de entrada do canal WEB; nome e flags preservados por compatibilidade
com o app desktop (Mkp/app/worker.py).

Uso (coloque este arquivo na mesma pasta do empilhar.py):
    python empilhar_web.py 10.09.2026                       # monta e confere; NAO grava nada
    python empilhar_web.py 10.09.2026 11.09.2026 --publicar # idem e grava no empilhado do X:
    python empilhar_web.py --teste                          # teste offline
Rodar de novo com um dia que ja esta no empilhado nao duplica nada (o dia e pulado).
"""

import sys

from empilhar_canais import main, teste

if __name__ == "__main__":
    if "--teste" in sys.argv:
        teste("WEB")
    else:
        main("WEB", sys.argv[1:])

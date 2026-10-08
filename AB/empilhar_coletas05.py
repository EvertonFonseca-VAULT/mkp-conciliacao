"""Empilhado do COLETAS 05 (Mercado Pago) -- monta as linhas do dia e grava direto no .xlsx do X:.

Mesmas premissas do empilhar_web.py:
Nao carrega o TOTVS. O "Nº QUE ESTA NO TOTVS" (coluna U) vem do arquivo Para Processar do dia
(C:\\RODAR TOTVS\\PROCESSAR\\COLETAS_CO05\\PROCESSAR_MERCADO COLETAS_CO05 DD.MM.AAAA.xlsx), cruzado com o AB por SOURCE_ID.
Como esse arquivo vem sem duplicados, as linhas do AB cujo SOURCE_ID nao esta nele usam o SHIPPING_ID / ORDER_ID /
EXTERNAL_REFERENCE da propria linha que exista no conjunto de numeros do Para Processar (se der 0 ou 2, para e avisa).
  * sem pasta de backup e sem copia local: le o .xlsx do X: e grava um ".novo" na mesma pasta;
    so no fim troca o original (os.replace, sem trafego). Se der erro antes, o original fica intacto.
  * Mesmo layout de 25 colunas do WEB (A..Y). O que muda, e o script LE DO PROPRIO ARQUIVO (ver _convencoes):
      - VALOR (W): formula SUM(I,J,L,K) (GROSS + MP_FEE + SHIPPING_FEE + FINANCING_FEE), aceita tambem SUM(I:K) e SUM(I+J+K);
      - SEQUENCIA (X): NUMERO FIXO = numero da linha (X2 = 2), nao formula; se a ultima linha do arquivo tiver formula,
        segue formula (X<anterior>+1); o deslocamento (linha x sequencia) tambem e lido da ultima linha;
      - CHAVE_PRIMARIA (Y) = CONCATENATE(X,"_COLETAS_CO05_MERCADO").
  * INSTALLMENTS (O) e gravado como TEXTO, igual ao exemplo do Coletas 05 (no WEB e numero).
  * nao mexe em calendario/tabela a nao ser que existam no arquivo.
  * sem cache em pickle: monta, confere e grava na mesma execucao.

A implementacao fica em empilhar_canais.py (modulo compartilhado dos canais -- ver Canal em CONTEXT.md).
Este arquivo e so o ponto de entrada do canal COLETAS_CO05; nome e flags preservados por compatibilidade
com o app desktop (Mkp/app/worker.py).

Uso (coloque este arquivo na mesma pasta do empilhar.py):
    python empilhar_coletas05.py 11.09.2026                        # monta e confere; NAO grava nada
    python empilhar_coletas05.py 11.09.2026 14.09.2026 --publicar  # idem e grava no empilhado do X:
    python empilhar_coletas05.py --teste                           # teste offline
Rodar de novo com um dia que ja esta no empilhado nao duplica nada (o dia e pulado).
"""

import sys

from empilhar_canais import main, teste

if __name__ == "__main__":
    if "--teste" in sys.argv:
        teste("COLETAS_CO05")
    else:
        main("COLETAS_CO05", sys.argv[1:])

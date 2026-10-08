"""Empilhado do PRIME (Mercado Pago) -- monta as linhas do dia e grava direto no .xlsx do X:.

Mesmas premissas do empilhar_web.py:
Nao carrega o TOTVS. O "PEDIDO TOTVS" (coluna AJ) vem do arquivo Para Processar do dia
(C:\\RODAR TOTVS\\PROCESSAR\\PRIME\\PROCESSAR_MERCADO PRIME DD.MM.AAAA.xlsx), cruzado com o AB por SOURCE_ID.
Como esse arquivo vem sem duplicados, as linhas do AB cujo SOURCE_ID nao esta nele usam o SHIPPING_ID / ORDER_ID /
EXTERNAL_REFERENCE da propria linha que exista no conjunto de numeros do Para Processar (se der 0 ou 2, para e avisa).
  * sem pasta de backup e sem copia local: le o .xlsx do X: e grava um ".novo" na mesma pasta;
    so no fim troca o original (os.replace, sem trafego). Se der erro antes, o original fica intacto.
  * VALOR LIQUIDO (coluna AL) e a formula SUM(I+J+K+L) (le do proprio arquivo, pode ser SUM(I:L) tambem);
    SEQUENCIA (AM) = AM<linha anterior>+1; CHAVE_PRIMARIA (AN) = CONCATENATE(AM,"_PRIME_MERCADO").
  * nao mexe em calendario/tabela a nao ser que existam no arquivo.
  * sem cache em pickle: monta, confere e grava na mesma execucao.

O que muda em relacao ao WEB:
  * o empilhado do PRIME tem 40 colunas (A..AN) em outra ordem: entra SELLER_AMOUNT, TAX_DETAIL, TAX_AMOUNT_TELCO,
    POS_*, STORE_*, CURRENCY, TAXES_DISAGGREGATED, METADATA, REFUND_ID, EFFECTIVE_COUPON_AMOUNT; nao tem
    "DESCRICAO 1/2" nem "ESTORNO" separados (sao RECORD_TYPE / DESCRIPTION / NET_DEBIT_AMOUNT).
  * o AB do PRIME vem com UMA ABA POR DIA DE REPASSE (nome da aba = dia, ex.: "11", "14"),
    e o arquivo pode ter mais de um dia. A aba lida e a do dia do repasse informado.
  * REFUND_ID nao existe no AB atual: a coluna AH fica vazia (se um dia o AB trouxer, e usada).

A implementacao fica em empilhar_canais.py (modulo compartilhado dos canais -- ver Canal em CONTEXT.md).
Este arquivo e so o ponto de entrada do canal PRIME; nome e flags preservados por compatibilidade
com o app desktop (Mkp/app/worker.py).

Uso (coloque este arquivo na mesma pasta do empilhar.py):
    python empilhar_prime.py 11.09.2026                        # monta e confere; NAO grava nada
    python empilhar_prime.py 11.09.2026 14.09.2026 --publicar  # idem e grava no empilhado do X:
    python empilhar_prime.py 14.09.2026 --ab="X:\\...\\MERCADO_PRIME_11_09_2026.xlsx"   # AB de outro dia (aba = dia)
    python empilhar_prime.py --teste                           # teste offline
Rodar de novo com um dia que ja esta no empilhado nao duplica nada (o dia e pulado).
"""

import sys

from empilhar_canais import main, teste

if __name__ == "__main__":
    if "--teste" in sys.argv:
        teste("PRIME")
    else:
        main("PRIME", sys.argv[1:])

"""Empilhado do COLETAS ES (Mercado Pago) -- monta as linhas do dia e grava direto no .xlsx do X:.

Mesmas premissas do empilhar_center.py/empilhar_web.py (ver Canal em CONTEXT.md):
Nao carrega o TOTVS. O "Nº QUE ESTA NO TOTVS" (coluna U) vem do arquivo Para Processar do dia
(C:\\RODAR TOTVS\\PROCESSAR\\COLETAS ES\\PROCESSAR_MERCADO COLETAS ES DD.MM.AAAA.xlsx -- ja gerado por
rodar_es.py, que faz o cruzamento com o TOTVS/SRI antes disso), cruzado com o AB por SOURCE_ID. Como esse
arquivo vem sem duplicados, as linhas do AB cujo SOURCE_ID nao esta nele usam o SHIPPING_ID / ORDER_ID /
EXTERNAL_REFERENCE da propria linha que exista no conjunto de numeros do Para Processar (se der 0 ou 2,
para e avisa).
  * sem pasta de backup e sem copia local: le o .xlsx do X: e grava um ".novo" na mesma pasta;
    so no fim troca o original (os.replace, sem trafego). Se der erro antes, o original fica intacto.
  * VALOR (W), SEQUENCIA (X) e sufixo da CHAVE_PRIMARIA (Y) sao lidos do proprio arquivo (_convencoes),
    nao do codigo -- o script segue exatamente o que ja esta gravado no empilhado.
  * sem cache em pickle: monta, confere e grava na mesma execucao.

A implementacao fica em empilhar_canais.py (modulo compartilhado dos canais -- ver Canal em CONTEXT.md).
Este arquivo e so o ponto de entrada do canal ES.

Uso:
    python empilhar_es.py 10.09.2026                       # monta e confere; NAO grava nada
    python empilhar_es.py 10.09.2026 11.09.2026 --publicar # idem e grava no empilhado do X:
    python empilhar_es.py --teste                          # teste offline
Rodar de novo com um dia que ja esta no empilhado nao duplica nada (o dia e pulado).
"""

import sys

from empilhar_canais import main, teste

if __name__ == "__main__":
    if "--teste" in sys.argv:
        teste("ES")
    else:
        main("ES", sys.argv[1:])

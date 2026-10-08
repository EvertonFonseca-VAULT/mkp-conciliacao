"""Empilhado do MOVEIS (Mercado Pago) -- monta as linhas do dia e grava direto no .xlsx do X:.

Mesmas premissas do empilhar_web.py / empilhar_coletas05.py:
Nao carrega o TOTVS. O "Nº QUE ESTA NO TOTVS" (coluna U) vem do arquivo Para Processar do dia
(C:\\RODAR TOTVS\\PROCESSAR\\MOVEIS\\PROCESSAR_MERCADO MOVEIS DD.MM.AAAA.xlsx), cruzado com o AB por SOURCE_ID.
Como esse arquivo vem sem duplicados, as linhas do AB cujo SOURCE_ID nao esta nele usam o SHIPPING_ID / ORDER_ID /
EXTERNAL_REFERENCE da propria linha que exista no conjunto de numeros do Para Processar (se der 0 ou 2, para e avisa).
  * sem pasta de backup e sem copia local: le o .xlsx do X: e grava um ".novo" na mesma pasta;
    so no fim troca o original (os.replace, sem trafego). Se der erro antes, o original fica intacto.
  * Mesmo layout de 25 colunas (A..Y). O que muda em relacao ao WEB / Coletas 05, e o script LE DO PROPRIO ARQUIVO (ver _convencoes):
      - VALOR (W): formula I+J++K (GROSS + MP_FEE + FINANCING_FEE), sem SUM (o "++" e copiado como esta no arquivo);
        aceita tambem SUM(I,J,L,K), SUM(I:K) e SUM(I+J+K);
      - SEQUENCIA (X): NUMERO FIXO = linha - 1 (X2 = 1), nao formula; se a ultima linha tiver formula, segue formula (X<anterior>+1);
      - CHAVE_PRIMARIA (Y) = CONCATENATE(X,"_MÓVEIS_MERCADO") (com acento, igual ao exemplo).
  * O AB do Moveis vem com cabecalhos em PORTUGUES (DATA DE LIBERAÇÃO, VALOR LÍQUIDO CREDITADO, ...) e a data como data/hora
    do Excel: o script traduz os cabecalhos e grava DATE como texto AAAA-MM-DDTHH:MM:SS,000-03:00 (virgula, como no exemplo do
    empilhado; ajuste o date_sufixo no cadastro do canal em empilhar_canais.py se precisar do formato com ponto). Tambem aceita
    AB com cabecalhos em ingles.
  * INSTALLMENTS (O) e gravado como TEXTO, igual ao exemplo.
  * nao mexe em calendario/tabela a nao ser que existam no arquivo.
  * sem cache em pickle: monta, confere e grava na mesma execucao.

A implementacao fica em empilhar_canais.py (modulo compartilhado dos canais -- ver Canal em CONTEXT.md).
Este arquivo e so o ponto de entrada do canal MOVEIS; nome e flags preservados por compatibilidade
com o app desktop (Mkp/app/worker.py).

Uso (coloque este arquivo na mesma pasta do empilhar.py):
    python empilhar_moveis.py 04.09.2026                        # monta e confere; NAO grava nada
    python empilhar_moveis.py 04.09.2026 11.09.2026 --publicar  # idem e grava no empilhado do X:
    python empilhar_moveis.py --teste                           # teste offline
Rodar de novo com um dia que ja esta no empilhado nao duplica nada (o dia e pulado).
"""

import sys

from empilhar_canais import main, teste

if __name__ == "__main__":
    if "--teste" in sys.argv:
        teste("MOVEIS")
    else:
        main("MOVEIS", sys.argv[1:])

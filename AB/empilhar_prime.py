"""Empilhado do PRIME -- ponto de entrada; implementacao e premissas em empilhar_canais.py.

Para Processar do dia: C:\\RODAR TOTVS\\PROCESSAR\\PRIME\\PROCESSAR_MERCADO PRIME DD.MM.AAAA.xlsx
(o numero do TOTVS vai na coluna AJ, "PEDIDO TOTVS").
O que muda em relacao aos demais: 40 colunas (A..AN), outra ordem (SELLER_AMOUNT, TAX_*, POS_*, STORE_*,
CURRENCY, METADATA, REFUND_ID, ...); AB vem com UMA ABA POR DIA DE REPASSE (nome = dia, ex. "11"), le a
aba do dia informado; VALOR LIQUIDO (AL) = SUM(I+J+K+L); SEQUENCIA (AM) = AM<anterior>+1;
CHAVE_PRIMARIA (AN) = CONCATENATE(AM,"_PRIME_MERCADO"). Nome e flags preservados por compatibilidade com
o app desktop (Mkp/app/worker.py).

Uso:
    python empilhar_prime.py 11.09.2026                        # monta e confere; NAO grava nada
    python empilhar_prime.py 11.09.2026 14.09.2026 --publicar  # idem e grava no empilhado do X:
    python empilhar_prime.py 14.09.2026 --ab="X:\\...\\MERCADO_PRIME_11_09_2026.xlsx"   # AB de outro dia (aba = dia)
    python empilhar_prime.py --teste                           # teste offline
"""

import sys

from empilhar_canais import main, teste

if __name__ == "__main__":
    if "--teste" in sys.argv:
        teste("PRIME")
    else:
        main("PRIME", sys.argv[1:])

"""Pergunta o mes e o arquivo AB do COLETAS ES -- mesma logica de
escolher_full.py (mes/arquivo, ano mais recente sozinho), so que apontando
pra pasta/prefixo do ES. Reaproveita `escolher_data` de escolher_full.py em
vez de duplicar (mesmo padrao que processar_es.py ja usa pra reusar
processar_full.py). Gera o mesmo "Para Processar" que
`rodar_es.py --data DD.MM.AAAA` ja gera (C:\\RODAR TOTVS\\PROCESSAR\\COLETAS
ES, sem mudar nada pro TOTVS). Uso via escolher_es.bat (2 cliques, salvo em
C:\\RODAR TOTVS). Pedido do Everton, 18/09/2026.
"""

import os
import subprocess
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "FULL"))
from escolher_full import escolher_data  # noqa: E402

from rodar_es import PASTA_AB_REDE_PADRAO  # noqa: E402


def main():
    data = escolher_data(PASTA_AB_REDE_PADRAO, "MERCADO COLETAS ES")
    print(f"\nProcessando {data}...")
    subprocess.run(
        [sys.executable, os.path.join(os.path.dirname(__file__), "rodar_es.py"), "--data", data],
        check=True,
    )


def _teste():
    """So confere a fiacao (prefixo/pasta certos, escolher_data reaproveitado
    de verdade) -- a logica de navegacao pasta/arquivo em si ja e coberta
    pelo self-test de escolher_full.py (mesma funcao)."""
    import tempfile

    with tempfile.TemporaryDirectory() as tmp:
        pasta_mes = os.path.join(tmp, "2026", "09.2026")
        os.makedirs(pasta_mes)
        open(os.path.join(pasta_mes, "MERCADO COLETAS ES 05.09.2026.xlsx"), "w").close()

        import builtins
        respostas = iter(["", ""])
        builtins.input = lambda *_: next(respostas)
        try:
            data = escolher_data(tmp, "MERCADO COLETAS ES")
        finally:
            del builtins.input
        assert data == "05.09.2026", data

    assert PASTA_AB_REDE_PADRAO.endswith("AB\\COLETAS ES"), PASTA_AB_REDE_PADRAO
    print("escolher_es.py: ok")


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        _teste()
    else:
        main()

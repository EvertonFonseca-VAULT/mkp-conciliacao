"""Pergunta o mes e o arquivo AB (dentro de PASTA_AB_REDE_PADRAO, rede) em
vez de digitar a data DD.MM.AAAA na mao -- gera o mesmo "Para Processar" que
`rodar_full.py --data DD.MM.AAAA` ja gera hoje (mesmo destino de sempre,
C:\\RODAR TOTVS\\PROCESSAR\\FULL -- nada muda pro TOTVS nem pro
interativo.py/resolver_para_processar). O ano e escolhido sozinho (o mais
recente) -- so pergunta mes e arquivo, como pedido; pra processar um ano
antigo continua existindo o rodar_full.bat de sempre (digita a data
inteira). Uso via escolher_full.bat (2 cliques, salvo em C:\\RODAR TOTVS).
Pedido do Everton, 18/09/2026.
"""

import glob
import os
import re
import subprocess
import sys

from rodar_full import PASTA_AB_REDE_PADRAO


def _listar_pastas(pasta):
    return sorted(
        (d for d in os.listdir(pasta) if os.path.isdir(os.path.join(pasta, d))),
        reverse=True,
    )


def _escolher(opcoes, rotulo):
    for i, opcao in enumerate(opcoes, 1):
        marca = "  <- mais recente" if i == 1 else ""
        print(f"  {i}) {opcao}{marca}")
    escolha = input(f"Escolha {rotulo} (Enter = mais recente): ").strip()
    if not escolha:
        return opcoes[0]
    if escolha.isdigit() and 1 <= int(escolha) <= len(opcoes):
        return opcoes[int(escolha) - 1]
    return escolha  # deixa digitar o nome direto tambem


def escolher_data(pasta_raiz=PASTA_AB_REDE_PADRAO, prefixo_arquivo="MERCADO FULL"):
    """Pasta_raiz = <pasta_raiz>/<ano>/<mes.ano>/<prefixo_arquivo>
    <data>.xlsx -- ano sempre o mais recente (sem perguntar, ver docstring
    do modulo), mes e arquivo perguntados. Retorna a data (DD.MM.AAAA)
    extraida do nome do arquivo escolhido. `prefixo_arquivo` generico (nao
    so "MERCADO FULL") pra reaproveitar em outros canais com a mesma
    estrutura de pasta -- ver escolher_es.py (mesmo padrao que
    processar_es.py ja usa pra reusar processar_full.py). Pedido do
    Everton, 18/09/2026."""
    # So pastas de ano (4 digitos) -- filtra lixo tipo "Macro Alissson" que
    # tambem mora em PASTA_AB_REDE_PADRAO e, sem o filtro, "vencia" o 2026 na
    # ordenacao alfabetica (M > 2), virando o "ano mais recente" errado.
    # Pedido do Everton, 18/09/2026.
    anos = [d for d in _listar_pastas(pasta_raiz) if re.fullmatch(r"\d{4}", d)]
    if not anos:
        sys.exit(f"Nenhuma pasta de ano em {pasta_raiz}.")
    pasta_ano = os.path.join(pasta_raiz, anos[0])

    meses = _listar_pastas(pasta_ano)
    if not meses:
        sys.exit(f"Nenhuma pasta de mes em {pasta_ano}.")
    print(f"\nMeses disponiveis em {pasta_ano}:")
    mes = _escolher(meses, "o mes")

    pasta_mes = os.path.join(pasta_ano, mes)
    arquivos = sorted(glob.glob(os.path.join(pasta_mes, f"{prefixo_arquivo} *.xlsx")), reverse=True)
    if not arquivos:
        sys.exit(f"Nenhum arquivo '{prefixo_arquivo} *.xlsx' em {pasta_mes}.")
    nomes = [os.path.basename(a) for a in arquivos]
    print(f"\nArquivos disponiveis em {pasta_mes}:")
    escolhido = _escolher(nomes, "o arquivo")

    regex_data = re.compile(rf"{re.escape(prefixo_arquivo)} (\d{{2}}\.\d{{2}}\.\d{{4}})\.xlsx$", re.IGNORECASE)
    m = regex_data.search(escolhido)
    if not m:
        sys.exit(f"Nome fora do padrao '{prefixo_arquivo} DD.MM.AAAA.xlsx': {escolhido}")
    return m.group(1)


def main():
    data = escolher_data()
    print(f"\nProcessando {data}...")
    subprocess.run(
        [sys.executable, os.path.join(os.path.dirname(__file__), "rodar_full.py"), "--data", data],
        check=True,
    )


def _teste():
    import tempfile

    with tempfile.TemporaryDirectory() as tmp:
        for ano, mes, dias in [("2025", "12.2025", ["30.12.2025"]),
                                ("2026", "08.2026", ["31.08.2026"]),
                                ("2026", "09.2026", ["01.09.2026", "02.09.2026"])]:
            pasta_mes = os.path.join(tmp, ano, mes)
            os.makedirs(pasta_mes)
            for dia in dias:
                open(os.path.join(pasta_mes, f"MERCADO FULL {dia}.xlsx"), "w").close()
        # Pasta lixo (nome, nao ano) que mora do lado dos anos de verdade na
        # rede (ex. "Macro Alissson") -- "M" > "2" alfabeticamente, entao sem
        # o filtro de 4 digitos ela "vencia" 2026 como "ano mais recente".
        # Bug real reportado pelo Everton, 18/09/2026.
        os.makedirs(os.path.join(tmp, "Macro Alissson"))

        # ano mais recente escolhido sozinho (2026, nao 2025 nem a pasta
        # lixo); mes com Enter
        # -> o mais recente (09.2026, nao 08.2026, mesmo pasta antiga tendo
        # ordenacao alfabetica proxima); arquivo com Enter -> o mais recente
        # (02.09.2026, nao 01.09.2026).
        import builtins
        respostas = iter(["", ""])
        builtins.input = lambda *_: next(respostas)
        try:
            data = escolher_data(pasta_raiz=tmp)
        finally:
            del builtins.input
        assert data == "02.09.2026", data

        # escolhendo por numero (2 = 01.09.2026, o 2o da lista mais recente->antigo)
        respostas = iter(["", "2"])
        builtins.input = lambda *_: next(respostas)
        try:
            data2 = escolher_data(pasta_raiz=tmp)
        finally:
            del builtins.input
        assert data2 == "01.09.2026", data2

    print("escolher_full.py: ok")


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        _teste()
    else:
        main()

"""Teste rápido de conexão com o banco Marketplace no Fabric.

Cria a tabela dbo._teste_conexao e insere algumas linhas em lote (não apaga
a tabela — fica lá pra você conferir pelo VSCode/SSMS). Serve só pra validar
autenticação + conectividade antes de rodar o importar.py inteiro (que
insere volume grande de uma vez).

Uso:
    python testar_conexao.py

Pra limpar depois de conferir:
    DROP TABLE dbo._teste_conexao;
"""

import time

from conexao import conectar

N_LINHAS = 100


def main():
    inicio = time.time()
    print("Conectando (só abre o navegador se não houver token em cache)...")
    conn = conectar(autocommit=False)
    print(f"  conectado em {time.time() - inicio:.1f}s")

    cur = conn.cursor()

    print("Criando tabela de teste...")
    cur.execute("IF OBJECT_ID('dbo._teste_conexao') IS NOT NULL DROP TABLE dbo._teste_conexao;")
    cur.execute("CREATE TABLE dbo._teste_conexao (id INT IDENTITY PRIMARY KEY, valor NVARCHAR(50));")
    conn.commit()

    print(f"Inserindo {N_LINHAS} linhas em lote...")
    t0 = time.time()
    cur.executemany(
        "INSERT INTO dbo._teste_conexao (valor) VALUES (?)",
        [(f"linha {i}",) for i in range(N_LINHAS)],
    )
    conn.commit()
    duracao = time.time() - t0
    print(f"  {N_LINHAS} linhas em {duracao:.2f}s ({N_LINHAS / duracao:.0f} linhas/s)")

    print("Lendo de volta...")
    cur.execute("SELECT COUNT(*) FROM dbo._teste_conexao")
    print("  total na tabela:", cur.fetchone()[0])

    conn.close()
    print("\nOK — conexão e insert em lote funcionando.")
    print("Tabela dbo._teste_conexao ficou no banco pra você conferir.")
    print("Pra limpar depois: DROP TABLE dbo._teste_conexao;")


if __name__ == "__main__":
    main()

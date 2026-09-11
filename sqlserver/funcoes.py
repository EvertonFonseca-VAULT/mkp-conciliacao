"""Consultas de negócio sobre o banco Marketplace (SQL Server / Microsoft Fabric)."""

from conexao import conectar as _conectar


def _query(sql, params=None):
    """Abre conexão, executa e fecha sempre (mesmo se `execute` falhar)."""
    conn = _conectar()
    try:
        cur = conn.cursor()
        cur.execute(sql, params or ())
        return cur.fetchall()
    finally:
        conn.close()


def pendencias_conciliacao():
    """Linhas onde EB, AB ou RB ainda não batem (ou faltam)."""
    return _query(
        """
        SELECT canal, loja, vertente, data, valor_eb, valor_ab, valor_rb, status_calculado
        FROM vw_conciliacao_status
        WHERE status_calculado != 'CONCILIADO'
        ORDER BY data DESC
        """
    )


def resumo_por_canal(data_inicio, data_fim):
    """Soma de valor_ab por canal num período, pra visão gerencial rápida."""
    return _query(
        """
        SELECT canal, loja, SUM(valor_ab) AS total_ab, COUNT(*) AS dias
        FROM vw_conciliacao_status
        WHERE data BETWEEN ? AND ?
        GROUP BY canal, loja
        ORDER BY total_ab DESC
        """,
        (data_inicio, data_fim),
    )


def canais_sem_cadastro_completo():
    """Canais criados automaticamente no import, ainda com vertente 'SEM ID'."""
    return _query("SELECT canal, loja FROM canais WHERE vertente = 'SEM ID'")


if __name__ == "__main__":
    print("Pendências de conciliação:")
    for linha in pendencias_conciliacao():
        print(" ", linha)

    print("\nCanais sem vertente cadastrada:")
    for linha in canais_sem_cadastro_completo():
        print(" ", linha)

"""Consultas de negócio sobre o banco Marketplace (SQL Server / Microsoft Fabric)."""

from conexao import conectar as _conectar


def pendencias_conciliacao():
    """Linhas onde EB, AB ou RB ainda não batem (ou faltam)."""
    conn = _conectar()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT canal, loja, vertente, data, valor_eb, valor_ab, valor_rb, status_calculado
        FROM vw_conciliacao_status
        WHERE status_calculado != 'CONCILIADO'
        ORDER BY data DESC
        """
    )
    resultado = cur.fetchall()
    conn.close()
    return resultado


def resumo_por_canal(data_inicio, data_fim):
    """Soma de valor_ab por canal num período, pra visão gerencial rápida."""
    conn = _conectar()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT canal, loja, SUM(valor_ab) AS total_ab, COUNT(*) AS dias
        FROM vw_conciliacao_status
        WHERE data BETWEEN ? AND ?
        GROUP BY canal, loja
        ORDER BY total_ab DESC
        """,
        (data_inicio, data_fim),
    )
    resultado = cur.fetchall()
    conn.close()
    return resultado


def canais_sem_cadastro_completo():
    """Canais criados automaticamente no import, ainda com vertente 'SEM ID'."""
    conn = _conectar()
    cur = conn.cursor()
    cur.execute("SELECT canal, loja FROM canais WHERE vertente = 'SEM ID'")
    resultado = cur.fetchall()
    conn.close()
    return resultado


if __name__ == "__main__":
    print("Pendências de conciliação:")
    for linha in pendencias_conciliacao():
        print(" ", linha)

    print("\nCanais sem vertente cadastrada:")
    for linha in canais_sem_cadastro_completo():
        print(" ", linha)

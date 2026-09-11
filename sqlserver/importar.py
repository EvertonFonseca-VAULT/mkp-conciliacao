"""
Importa as bases do Magalu para o banco Marketplace no SQL Server (Microsoft
Fabric SQL database).

Versão T-SQL/SQL Server do importar.py (SQLite) da raiz do projeto — mesma
lógica de negócio (mesmos dois tipos de arquivo reconhecidos, mesmo dedup
por hash, mesmo recálculo de valor_ab), só troca o motor de banco.

Pré-requisitos (uma vez só):
    pip install mssql-python azure-identity pandas openpyxl

Usa o driver "mssql-python" da própria Microsoft (pip-only, embute o driver
ODBC dentro do pacote) em vez do pyodbc — não precisa instalar nenhum
driver ODBC no Windows nem ter direitos de administrador.

Autenticação: login Azure AD/Entra ID via navegador, com cache de token em
disco (ver conexao.py) — só pede login de novo quando o token expira ou é
revogado, não a cada execução. Precisa de rede liberada até
*.fabric.microsoft.com. O banco já existe no Fabric (não precisa de
CREATE DATABASE).
"""

import glob
import hashlib
import os
import re

import pandas as pd

from conexao import DATABASE, SERVER, conectar

SCHEMA = os.path.join(os.path.dirname(__file__), "schema.sql")
PASTA_HISTORICO = os.path.join(os.path.dirname(__file__), "..", "Projeto Imports")
PASTA_NOVOS = os.path.join(os.path.dirname(__file__), "..", "Dados", "magalu")


def _bulk_insert(conn, tabela, colunas, df):
    """Insere em lote via executemany (sem depender de sqlalchemy/pyodbc)."""
    dados = df[colunas].astype(object).where(pd.notnull(df[colunas]), None).values.tolist()
    if not dados:
        return
    placeholders = ", ".join(["?"] * len(colunas))
    conn.cursor().executemany(
        f"INSERT INTO {tabela} ({', '.join(colunas)}) VALUES ({placeholders})",
        dados,
    )
    conn.commit()


def criar_schema(conn):
    """Executa schema.sql, dividindo em lotes por linhas 'GO' (separador de
    lote do sqlcmd/SSMS, que o driver ODBC não entende nativamente)."""
    with open(SCHEMA, "r", encoding="utf-8") as f:
        script = f.read()
    lotes = re.split(r"^\s*GO\s*$", script, flags=re.IGNORECASE | re.MULTILINE)
    cur = conn.cursor()
    for lote in lotes:
        lote = lote.strip()
        if lote:
            cur.execute(lote)
    conn.commit()


def hash_arquivo(caminho):
    h = hashlib.sha256()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(1024 * 1024), b""):
            h.update(bloco)
    return h.hexdigest()


def ja_importado(conn, hash_):
    row = conn.cursor().execute(
        "SELECT id FROM importacoes_log WHERE hash_arquivo = ?", (hash_,)
    ).fetchone()
    return row is not None


def para_data_iso(valor):
    """Aceita datetime, 'dd/mm/yyyy' ou None/NaT e devolve 'YYYY-MM-DD' ou None."""
    if valor is None or pd.isna(valor):
        return None
    if isinstance(valor, str):
        valor = pd.to_datetime(valor, format="%d/%m/%Y", errors="coerce")
        if pd.isna(valor):
            return None
    return pd.Timestamp(valor).strftime("%Y-%m-%d")


def parcela_para_int(valor):
    """Aceita int, float ou string tipo '4/6' e devolve o número da parcela atual."""
    if pd.isna(valor):
        return None
    if isinstance(valor, str) and "/" in valor:
        valor = valor.split("/")[0]
    try:
        return int(float(valor))
    except (TypeError, ValueError):
        return None


COLUNAS_INT_MKP_MAGALU = ["total_parcelas"]
COLUNAS_DECIMAL_MKP_MAGALU = [
    "valor_repasse", "valor_antecipacao", "perc_taxa_antecipacao",
    "comissao_por_parcela", "valor_bruto_parcela", "valor_bruto_pedido",
    "perc_comissao", "valor_comissao", "valor_bruto_seller",
    "valor_fixo_pedido_parcelada", "valor_fixo_pedido",
]
COLUNAS_TEXTO_MKP_MAGALU = [
    "id_pedido", "pedido_mkp", "id_transacao", "numero_nota_fiscal",
    "nome_cliente", "metodo_pagamento", "origem", "observacoes",
    "external_withdraw_id",
]


def _para_inteiro(valor):
    """Converte pra int nativo (ou None); tolera texto que não é número."""
    if pd.isna(valor):
        return None
    try:
        return int(float(valor))
    except (TypeError, ValueError):
        return None


def _para_decimal(valor):
    """Converte pra float nativo (ou None); tolera texto que não é número."""
    if pd.isna(valor):
        return None
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def _para_texto(valor):
    """Converte pra str nativa (ou None); evita '123.0' em IDs numéricos."""
    if pd.isna(valor):
        return None
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    valor = str(valor).strip()
    return valor or None


def _normalizar_tipos_mkp_magalu(df):
    """Garante um tipo Python homogêneo por coluna. Sem isso, uma planilha
    que mistura número e texto na mesma coluna (comum em dado digitado à
    mão) faz o executemany do mssql-python quebrar no meio do lote, porque
    ele fixa o tipo do parâmetro pela primeira linha."""
    for col in COLUNAS_INT_MKP_MAGALU:
        df[col] = df[col].apply(_para_inteiro)
    for col in COLUNAS_DECIMAL_MKP_MAGALU:
        df[col] = df[col].apply(_para_decimal)
    for col in COLUNAS_TEXTO_MKP_MAGALU:
        df[col] = df[col].apply(_para_texto)


def garantir_canal(conn, canal, loja):
    cur = conn.cursor()
    row = cur.execute(
        "SELECT id FROM canais WHERE canal = ? AND loja = ?", (canal, loja)
    ).fetchone()
    if row:
        return row[0]
    print(f"  aviso: canal '{canal}'/'{loja}' não estava cadastrado em canais — "
          f"criado com vertente 'SEM ID', ajuste depois se precisar.")
    row = cur.execute(
        "INSERT INTO canais (canal, loja, vertente) OUTPUT INSERTED.id VALUES (?, ?, 'SEM ID')",
        (canal, loja),
    ).fetchone()
    conn.commit()
    return row[0]


def atualizar_valor_ab(conn, canal, loja):
    """Recalcula conciliacao_diaria.valor_ab a partir da soma de valor_repasse
    em mkp_magalu, por data, para o canal+loja informado."""
    canal_id = garantir_canal(conn, canal, loja)
    cur = conn.cursor()
    linhas = cur.execute(
        """
        SELECT data_repasse, SUM(valor_repasse)
        FROM mkp_magalu
        WHERE loja = ? AND data_repasse IS NOT NULL
        GROUP BY data_repasse
        """,
        (loja,),
    ).fetchall()
    for data, soma in linhas:
        cur.execute(
            """
            MERGE conciliacao_diaria AS destino
            USING (SELECT ? AS canal_id, ? AS data, ? AS valor_ab) AS origem
            ON destino.canal_id = origem.canal_id AND destino.data = origem.data
            WHEN MATCHED THEN
                UPDATE SET valor_ab = origem.valor_ab, atualizado_em = SYSDATETIME()
            WHEN NOT MATCHED THEN
                INSERT (canal_id, data, valor_ab) VALUES (origem.canal_id, origem.data, origem.valor_ab);
            """,
            (canal_id, data, soma),
        )
    conn.commit()


MAPA_EMPILHADO_MAGALU = {
    "Data_Repasse": "data_repasse",
    "Data da transação": "data_transacao",
    "Data do pedido": "data_pedido",
    "ID do pedido Magazine Luiza": "id_pedido",
    "Pedido MKP": "pedido_mkp",
    "ID da transação": "id_transacao",
    "Número da nota fiscal": "numero_nota_fiscal",
    "Nome do cliente": "nome_cliente",
    "Método de pagamento": "metodo_pagamento",
    "Parcela atual": "parcela_atual",
    "Total de parcelas": "total_parcelas",
    "Valor Repasse": "valor_repasse",
    "Valor da antecipação": "valor_antecipacao",
    "% Taxa de antecipação": "perc_taxa_antecipacao",
    "Comissão ML por parcela": "comissao_por_parcela",
    "Valor bruto da parcela": "valor_bruto_parcela",
    "Valor bruto do pedido": "valor_bruto_pedido",
    "% Comissão": "perc_comissao",
    "Valor da comissão": "valor_comissao",
    "Valor bruto seller": "valor_bruto_seller",
    "Origem": "origem",
    "Observações": "observacoes",
    "external_withdraw_id": "external_withdraw_id",
    "Valor fixo por pedido parcelada": "valor_fixo_pedido_parcelada",
    "Valor fixo por pedido": "valor_fixo_pedido",
}

MAPA_RAW_MAGALU = {
    "ID da transação": "id_transacao",
    "Origem": "origem",
    "Data do repasse": "data_repasse",
    "Número do pedido": "pedido_mkp",
    "Data do pedido": "data_pedido",
    "Nome do cliente": "nome_cliente",
    "Número da nota fiscal": "numero_nota_fiscal",
    "Forma de pagamento": "metodo_pagamento",
    "Total de parcelas": "total_parcelas",
    "Valor do bruto do pedido para base calculo (R$)": "valor_bruto_pedido",
    "Valor de desconto da Tarifa fixa por pedido (R$)": "valor_fixo_pedido",
    "Parcela atual": "parcela_atual",
    "Valor bruto da parcela (R$)": "valor_bruto_parcela",
    "% de desconto dos Serviços de Marketplace por parcela": "perc_comissao",
    "Valor de desconto dos Serviços do marketplace por parcela (R$)": "valor_comissao",
    "Valor de desconto da Tarifa fixa por parcela (R$)": "valor_fixo_pedido_parcelada",
    "% de taxa de Antecipação por parcela": "perc_taxa_antecipacao",
    "Custo da Antecipação por parcela (R$)": "valor_antecipacao",
    "Valor do repasse financeiro da parcela (R$)": "valor_repasse",
    "Observações": "observacoes",
    "ID do repasse": "external_withdraw_id",
}

COLUNAS_MKP_MAGALU = [
    "loja", "data_repasse", "data_transacao", "data_pedido", "id_pedido",
    "pedido_mkp", "id_transacao", "numero_nota_fiscal", "nome_cliente",
    "metodo_pagamento", "parcela_atual", "total_parcelas", "valor_repasse",
    "valor_antecipacao", "perc_taxa_antecipacao", "comissao_por_parcela",
    "valor_bruto_parcela", "valor_bruto_pedido", "perc_comissao",
    "valor_comissao", "valor_bruto_seller", "origem", "observacoes",
    "external_withdraw_id", "valor_fixo_pedido_parcelada", "valor_fixo_pedido",
    "arquivo_origem",
]


def importar_historico_empilhado(conn, caminho, loja="CENTER"):
    """Import único de backfill a partir da aba EMPILHADO (histórico já
    consolidado em Excel). Assume que todo o arquivo é de uma loja só,
    já que hoje o EMPILHADO não guarda a loja como coluna."""
    hash_ = hash_arquivo(caminho)
    if ja_importado(conn, hash_):
        print(f"  já importado antes, pulando: {caminho}")
        return

    df = pd.read_excel(caminho, sheet_name="EMPILHADO")
    df = df.rename(columns=MAPA_EMPILHADO_MAGALU)
    df["loja"] = loja
    df["arquivo_origem"] = os.path.basename(caminho)

    df["data_repasse"] = df["data_repasse"].apply(para_data_iso)
    df["data_transacao"] = df["data_transacao"].apply(para_data_iso)
    df["data_pedido"] = df["data_pedido"].apply(para_data_iso)
    df["parcela_atual"] = df["parcela_atual"].apply(parcela_para_int)

    for col in COLUNAS_MKP_MAGALU:
        if col not in df.columns:
            df[col] = None
    _normalizar_tipos_mkp_magalu(df)

    _bulk_insert(conn, "mkp_magalu", COLUNAS_MKP_MAGALU, df)
    conn.cursor().execute(
        "INSERT INTO importacoes_log (canal, arquivo, hash_arquivo, linhas) VALUES (?, ?, ?, ?)",
        ("MAGALU", os.path.basename(caminho), hash_, len(df)),
    )
    conn.commit()
    print(f"  histórico importado: {caminho} ({len(df)} linhas)")
    atualizar_valor_ab(conn, "MAGALU", loja)


def importar_base_periodica_magalu(conn, caminho):
    """Import incremental de um export bruto periódico do Magalu (a
    'nova base' que seria colada manualmente no fim do empilhado)."""
    hash_ = hash_arquivo(caminho)
    if ja_importado(conn, hash_):
        print(f"  já importado antes, pulando: {caminho}")
        return

    df = pd.read_excel(caminho)
    if "MKP" not in df.columns:
        print(f"  aviso: {caminho} não tem coluna 'MKP', pulando (layout não reconhecido).")
        return

    partes = df["MKP"].iloc[0].split("_", 1)
    canal = partes[0].strip().upper()
    loja = partes[1].strip().upper() if len(partes) > 1 else "SEM ID"

    df = df.rename(columns=MAPA_RAW_MAGALU)
    df["loja"] = loja
    df["arquivo_origem"] = os.path.basename(caminho)

    df["data_repasse"] = df["data_repasse"].apply(para_data_iso)
    df["data_pedido"] = df["data_pedido"].apply(para_data_iso)
    df["parcela_atual"] = df["parcela_atual"].apply(parcela_para_int)

    for col in COLUNAS_MKP_MAGALU:
        if col not in df.columns:
            df[col] = None
    _normalizar_tipos_mkp_magalu(df)

    _bulk_insert(conn, "mkp_magalu", COLUNAS_MKP_MAGALU, df)
    conn.cursor().execute(
        "INSERT INTO importacoes_log (canal, arquivo, hash_arquivo, linhas) VALUES (?, ?, ?, ?)",
        (canal, os.path.basename(caminho), hash_, len(df)),
    )
    conn.commit()
    print(f"  base periódica importada: {caminho} ({len(df)} linhas, {canal}/{loja})")
    atualizar_valor_ab(conn, canal, loja)


if __name__ == "__main__":
    conn = conectar(autocommit=False)
    try:
        criar_schema(conn)

        historico = [
            c for c in glob.glob(os.path.join(PASTA_HISTORICO, "*.xlsx"))
            if not os.path.basename(c).startswith("~$")
        ]
        for caminho in historico:
            try:
                if "EMPILHADO" in pd.ExcelFile(caminho).sheet_names:
                    importar_historico_empilhado(conn, caminho)
            except Exception as e:
                print(f"  erro lendo {caminho}: {e}")

        os.makedirs(PASTA_NOVOS, exist_ok=True)
        novos = glob.glob(os.path.join(PASTA_NOVOS, "*.xlsx"))
        if not novos:
            print(f"Nenhuma base nova em '{PASTA_NOVOS}/'. Coloque os exports do Magalu lá.")
        for caminho in novos:
            importar_base_periodica_magalu(conn, caminho)
    finally:
        conn.close()
    print(f"\nPronto. Banco '{DATABASE}' atualizado em {SERVER}.")

"""Banco central do PTI (Fabric SQL database; SQLite so em teste/offline).

O PTI (export do TOTVS) e um ACUMULADO: cada arquivo traz ~6 meses de pedidos
ate o dia mais recente, por canal. Entao nao ha chave por decendio -- cada
importacao e um retrato completo e SUBSTITUI o anterior daquele canal+vendedor
(numa transacao so: se falhar no meio, o dado antigo continua). O FULL tem 2
vendedores (750/751) em arquivos separados; a substituicao e por vendedor, pra
importar um nao apagar o outro.

Conexao SEM pyodbc e SEM driver ODBC: python-tds (Python puro, fala TDS direto)
+ azure-identity (login Microsoft interativo no navegador; o token fica em cache
na maquina). Variavel de ambiente FABRIC_CONN com Server=host,porta e
Database=nome (a mesma string do ODBC serve; Driver/Authentication sao
ignorados). A string NAO vai no codigo nem no git. COMUM nao depende de nenhuma
frente: recebe o DataFrame ja lido (PARA_PROCESSAR carregar_pti).
"""

import hashlib
import os
import sqlite3
from datetime import datetime

import pandas as pd

ENV_CONN = "FABRIC_CONN"
ENV_SQLITE = "PTI_SQLITE"  # caminho de um .db local: teste/offline
ESCOPO_SQL = "https://database.windows.net/.default"

# Colunas do PTI (cabecalho do export TOTVS) -> nome SQL. Ordem = a do arquivo.
COLUNAS = {
    "COD CLIENTE": "cod_cliente", "LOJA": "loja", "NOME CLIENTE": "nome_cliente",
    "VENDEDOR": "vendedor", "NOME VENDEDOR": "nome_vendedor", "FRANQUIA": "franquia",
    "PREFIXO": "prefixo", "NÚMERO": "numero", "TIPO": "tipo", "PARCELA": "parcela",
    "FILIAL": "filial", "PEDIDO": "pedido", "PED. SITE": "ped_site", "PED. MKP": "ped_mkp",
    "FORMA PAGTO.": "forma_pagto", "VENCTO. REAL": "vencto_real", "VALOR REAIS": "valor_reais",
    "ATRASO DIAS": "atraso_dias", "NRO. BOLETO": "nro_boleto", "DT. EMISSÃO": "dt_emissao",
    "COD. PARCEIRO": "cod_parceiro", "NOME PARCEIRO": "nome_parceiro", "HISTÓRICO": "historico",
    "SITUAÇÃO HIST.": "situacao_hist", "VALOR": "valor", "PERC. JUROS": "perc_juros",
    "DT. BAIXA": "dt_baixa", "ULTIMA DT. CANC. BAIXA": "ultima_dt_canc_baixa",
    "EMISSÃO PED.": "emissao_ped", "FRETE": "frete", "ID TRANS": "id_trans",
    "ADM CART": "adm_cart",
}
# Ja numericos em carregar_pti (pt-BR -> float); o resto fica texto, como no CSV.
NUMERICAS = ("VALOR REAIS", "VALOR", "FRETE")

_LOTE_IN = 1000  # SQL Server aceita ate 2100 parametros por comando


def _sqlite(conn):
    return isinstance(conn, sqlite3.Connection)


def _marcador(conn):
    """sqlite3 usa '?', python-tds usa '%s'."""
    return "?" if _sqlite(conn) else "%s"


_LINHAS_POR_INSERT = 60  # x 34 colunas = 2040 < 2100 parametros do SQL Server


def _parse_conn(cadeia):
    """'Server=h,1433;Database={db};...' -> (host, porta, banco)."""
    partes = {}
    for trecho in cadeia.split(";"):
        if "=" in trecho:
            k, v = trecho.split("=", 1)
            partes[k.strip().lower()] = v.strip().strip("{}")
    try:
        host, _, porta = partes["server"].removeprefix("tcp:").partition(",")
        return host, int(porta or 1433), partes["database"]
    except KeyError as e:
        raise RuntimeError(f"{ENV_CONN} sem Server= ou Database=.") from e


def _token_fabric():
    """Login Microsoft interativo (abre o navegador; cache em disco quando
    possivel, pra nao pedir a cada execucao)."""
    from azure.identity import InteractiveBrowserCredential, TokenCachePersistenceOptions

    try:
        cred = InteractiveBrowserCredential(
            cache_persistence_options=TokenCachePersistenceOptions(name="mkp_conciliacao")
        )
    except Exception:  # sem suporte a cache persistente nessa maquina
        cred = InteractiveBrowserCredential()
    return cred.get_token(ESCOPO_SQL).token


def conectar():
    """Abre a conexao conforme o ambiente. Erro claro se faltar configuracao."""
    caminho = os.environ.get(ENV_SQLITE)
    if caminho:
        return sqlite3.connect(caminho)
    cadeia = os.environ.get(ENV_CONN)
    if not cadeia:
        raise RuntimeError(
            f"Defina a variavel de ambiente {ENV_CONN} com a connection string do Fabric "
            f"(ou {ENV_SQLITE} com um .db local pra teste)."
        )
    host, porta, banco = _parse_conn(cadeia)
    try:
        import certifi
        import pytds
    except ImportError as e:
        raise RuntimeError("Faltam dependencias: pip install python-tds azure-identity certifi") from e
    try:
        return pytds.connect(
            server=host, port=porta, database=banco,
            access_token_callable=_token_fabric,
            cafile=certifi.where(), validate_host=True,  # Fabric exige TLS
            login_timeout=120, autocommit=False,
        )
    except Exception as e:
        raise RuntimeError(f"Falha ao conectar no Fabric ({host}). Detalhe: {e}") from e


def criar_tabelas(conn):
    """Idempotente."""
    texto = "TEXT" if _sqlite(conn) else "NVARCHAR(400)"
    real = "REAL" if _sqlite(conn) else "FLOAT"
    cols = ",\n  ".join(
        f"{sql} {real if orig in NUMERICAS else texto} NULL" for orig, sql in COLUNAS.items()
    )
    if _sqlite(conn):
        ddl = [
            "CREATE TABLE IF NOT EXISTS pti_carga (carga_id INTEGER PRIMARY KEY AUTOINCREMENT, "
            "canal TEXT NOT NULL, vendedores TEXT NOT NULL, arquivo TEXT, importado_por TEXT, "
            "importado_em TEXT NOT NULL, n_linhas INTEGER NOT NULL, hash TEXT NOT NULL)",
            f"CREATE TABLE IF NOT EXISTS pti_linha (canal TEXT NOT NULL, carga_id INTEGER NOT NULL,\n  {cols})",
            "CREATE INDEX IF NOT EXISTS ix_pti_linha_ped ON pti_linha (canal, ped_mkp)",
        ]
    else:
        ddl = [
            "IF OBJECT_ID('pti_carga','U') IS NULL CREATE TABLE pti_carga (carga_id INT IDENTITY PRIMARY KEY, "
            "canal NVARCHAR(40) NOT NULL, vendedores NVARCHAR(200) NOT NULL, arquivo NVARCHAR(400) NULL, "
            "importado_por NVARCHAR(200) NULL, importado_em DATETIME2 NOT NULL, n_linhas INT NOT NULL, "
            "hash CHAR(64) NOT NULL)",
            f"IF OBJECT_ID('pti_linha','U') IS NULL CREATE TABLE pti_linha (canal NVARCHAR(40) NOT NULL, "
            f"carga_id INT NOT NULL,\n  {cols})",
            "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='ix_pti_linha_ped') "
            "CREATE INDEX ix_pti_linha_ped ON pti_linha (canal, ped_mkp)",
        ]
    cur = conn.cursor()
    for comando in ddl:
        cur.execute(comando)
    conn.commit()


def _valor(v):
    return None if pd.isna(v) else v


def importar_pti(conn, canal, pti, arquivo=None, usuario=None):
    """Grava o retrato completo `pti` (DataFrame de carregar_pti) do `canal`,
    substituindo o anterior dos vendedores presentes. Devolve um resumo.

    Atomico: delete + insert + registro da carga numa transacao."""
    faltando = [c for c in COLUNAS if c not in pti.columns]
    extras = [c for c in pti.columns if c not in COLUNAS]
    if faltando or extras:
        raise ValueError(f"Cabecalho do PTI diferente do esperado. Faltando: {faltando}. Extras: {extras}.")
    if pti.empty:
        raise ValueError("PTI vazio: nada a importar (nao apago o que ja esta no banco).")

    pti = pti[list(COLUNAS)]
    vendedores = sorted(pti["VENDEDOR"].dropna().astype(str).str.strip().unique())
    # hash do conteudo: detecta reimportacao do mesmo arquivo
    h = hashlib.sha256(pd.util.hash_pandas_object(pti.fillna(""), index=False).values.tobytes()).hexdigest()
    ph = _marcador(conn)
    cur = conn.cursor()
    try:
        # delete casa pelo valor cru gravado (pode ter espaco), nao pelo normalizado
        crus = sorted(pti["VENDEDOR"].dropna().astype(str).unique())
        cur.execute(
            f"DELETE FROM pti_linha WHERE canal = {ph} AND vendedor IN ({','.join([ph] * len(crus))})",
            [canal, *crus],
        )
        removidas = cur.rowcount
        cur.execute(
            f"INSERT INTO pti_carga (canal, vendedores, arquivo, importado_por, importado_em, n_linhas, hash) "
            f"VALUES ({','.join([ph] * 7)})",
            [canal, ",".join(vendedores), arquivo, usuario, datetime.now().isoformat(timespec="seconds"),
             len(pti), h],
        )
        if _sqlite(conn):
            carga_id = cur.lastrowid
        else:  # escopo da sessao: nao confunde com carga de outra maquina
            cur.execute("SELECT CAST(SCOPE_IDENTITY() AS INT)")
            carga_id = cur.fetchone()[0]

        nomes = ["canal", "carga_id", *COLUNAS.values()]
        linhas = [
            [canal, carga_id, *[_valor(v) for v in reg]]
            for reg in pti.itertuples(index=False, name=None)
        ]
        cols = ",".join(nomes)
        if _sqlite(conn):
            cur.executemany(f"INSERT INTO pti_linha ({cols}) VALUES ({','.join([ph] * len(nomes))})", linhas)
        else:  # python-tds nao tem bulk: INSERT multi-linha em lotes
            um = "(" + ",".join([ph] * len(nomes)) + ")"
            for i in range(0, len(linhas), _LINHAS_POR_INSERT):
                lote = linhas[i:i + _LINHAS_POR_INSERT]
                cur.execute(
                    f"INSERT INTO pti_linha ({cols}) VALUES {','.join([um] * len(lote))}",
                    [v for reg in lote for v in reg],
                )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    return {"canal": canal, "vendedores": vendedores, "linhas": len(pti),
            "substituidas": removidas, "carga_id": carga_id, "hash": h}


def buscar_pti(conn, canal, pedidos=None):
    """PTI do canal como DataFrame no mesmo formato de carregar_pti (nomes de
    coluna originais, 3 colunas float, resto texto). `pedidos` (PED. MKP)
    restringe a consulta; None traz tudo do canal."""
    sql_cols = ",".join(COLUNAS.values())
    ph = _marcador(conn)
    if pedidos is None:
        lotes = [None]
    else:
        pedidos = sorted({str(p).strip() for p in pedidos if pd.notna(p)})
        lotes = [pedidos[i:i + _LOTE_IN] for i in range(0, len(pedidos), _LOTE_IN)]
    partes = []
    for lote in lotes:
        sql = f"SELECT {sql_cols} FROM pti_linha WHERE canal = {ph}"
        params = [canal]
        if lote is not None:
            sql += f" AND ped_mkp IN ({','.join([ph] * len(lote))})"
            params += lote
        cur = conn.cursor()
        cur.execute(sql, params)
        partes.append(pd.DataFrame.from_records(cur.fetchall(), columns=list(COLUNAS)))
    pti = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame(columns=list(COLUNAS))
    for orig in COLUNAS:
        if orig in NUMERICAS:
            pti[orig] = pd.to_numeric(pti[orig], errors="coerce")
        else:
            pti[orig] = pti[orig].astype(object).where(pti[orig].notna(), None)
    return pti


def resumo_cargas(conn, canal=None):
    """Ultima importacao por canal (pra o app mostrar 'PTI ate <dia>')."""
    cur = conn.cursor()
    sql = "SELECT canal, vendedores, arquivo, importado_por, importado_em, n_linhas FROM pti_carga"
    params = []
    if canal:
        sql += f" WHERE canal = {_marcador(conn)}"
        params.append(canal)
    cur.execute(sql + " ORDER BY carga_id DESC", params)
    return cur.fetchall()

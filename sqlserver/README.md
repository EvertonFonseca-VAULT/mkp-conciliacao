# Banco Marketplace (SQL Server / Microsoft Fabric)

Implementação em T-SQL do banco `Marketplace`. O alvo é um **banco SQL do
Microsoft Fabric** já provisionado (não um SQL Server local).

## Conexão

```
Server=2syk64xvbs6etmnq4gvqaftbom-lxq4dihtl5nuzpponrs3hni6cm.database.fabric.microsoft.com,1433
Database={Marketplace-76c1ddc3-472f-4bf3-9445-ea7d2537f608}
Encrypt=yes;TrustServerCertificate=no
```

A autenticação **não** vai na connection string — é feita à parte via
`token_provider` (ver seção Autenticação abaixo).

Esses valores já estão hardcoded como default em `conexao.py`.
Para apontar pra outro servidor/banco sem editar o código, defina
`MSSQL_SERVER` / `MSSQL_DATABASE` antes de rodar.

Como o banco já existe no Fabric, os scripts **não** têm um passo de
`CREATE DATABASE` — só conectam direto nele e rodam o `schema.sql`.

## Pré-requisitos

Os scripts usam o driver **[`mssql-python`](https://github.com/microsoft/mssql-python)**
(da própria Microsoft) em vez de `pyodbc`. Ele embute o driver ODBC dentro
do próprio pacote pip — **não precisa instalar nada no Windows nem ter
direitos de administrador** (diferente do "ODBC Driver 18/17 for SQL
Server" tradicional, cujo instalador exige elevação UAC).

```bash
pip install mssql-python azure-identity pandas openpyxl
```

(`azure-identity` já vem como dependência do `mssql-python`, mas é usado
diretamente em `conexao.py` pro cache de login — ver seção Autenticação.)

> Se preferir o `pyodbc` tradicional em vez do `mssql-python`, é possível,
> mas exige instalar o "ODBC Driver 18 for SQL Server" da Microsoft
> (instalador `.msi`, precisa de admin):
> https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server

## Autenticação

Login interativo via Azure AD/Entra ID, implementado em `conexao.py` com
`InteractiveBrowserCredential` (pacote `azure-identity`) e cache de token
persistido em disco (`TokenCachePersistenceOptions`, via `msal-extensions`,
criptografado com DPAPI no Windows). Na prática:

- Só abre o navegador pra login **na primeira vez** (ou quando o token
  expira/é revogado) — as próximas execuções, mesmo em processos novos,
  reaproveitam o token salvo silenciosamente.
- Conta esperada: `everton.fonseca@webcontinental.com.br` (ou outra com
  acesso ao workspace do Fabric).
- Precisa de rede liberada até `*.fabric.microsoft.com` e
  `login.microsoftonline.com`.
- Pra forçar novo login (ex.: trocar de conta), apague a pasta:
  `%LOCALAPPDATA%\mkp-fabric-auth` (guarda o `AuthenticationRecord`, que diz
  qual conta procurar no cache — sem ela, o login silencioso não funciona
  mesmo com o cache de token intacto).

## Passo a passo

1. Instale as dependências (pré-requisito acima).
2. Rode a importação (roda o `schema.sql` no banco do Fabric e importa o
   histórico + bases novas do Magalu):
   ```bash
   python importar.py
   ```
   Vai abrir o navegador pra login na primeira conexão. Ele lê:
   - Histórico: `../Projeto Imports/*.xlsx` (aba `EMPILHADO`)
   - Bases novas: `../Dados/magalu/*.xlsx`
   Arquivos já importados (mesmo hash SHA-256) são ignorados — pode rodar
   quantas vezes quiser.
3. Consulte pendências e resumos prontos:
   ```bash
   python funcoes.py
   ```

## Conectando pelo VSCode (extensão MSSQL)

1. Ícone do SQL Server na barra lateral → *Add Connection*
2. Server: `2syk64xvbs6etmnq4gvqaftbom-lxq4dihtl5nuzpponrs3hni6cm.database.fabric.microsoft.com,1433`
3. Authentication: **Microsoft Entra Id - Universal w/ MFA support** (equivalente ao ActiveDirectoryInteractive)
4. Database: `Marketplace-76c1ddc3-472f-4bf3-9445-ea7d2537f608`
5. Encrypt: `Mandatory` (Encrypt=yes)
6. Explore tabelas/rode queries direto no editor `.sql` do VSCode.

## Observações conhecidas

- `valor_eb` e `valor_rb` ainda são preenchidos manualmente em
  `conciliacao_diaria` — cada um vai ganhar seu próprio processo de import
  automático depois.
- O export bruto periódico do Magalu não traz `comissao_por_parcela`,
  `valor_bruto_seller`, `id_pedido` nem `data_transacao` separadamente —
  ficam `NULL` em linhas importadas a partir de agora.

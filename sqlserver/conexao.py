"""Conexão compartilhada com o banco Marketplace no Fabric.

Usa um credential do azure-identity com cache de token persistido em disco
(via msal-extensions, no Windows guardado criptografado com DPAPI). Isso faz
o login no navegador acontecer só na primeira vez — nas próximas execuções
(mesmo em processos novos) o token salvo é reaproveitado/renovado
silenciosamente, sem abrir navegador de novo, até ele ser revogado ou expirar
de vez.

Detalhe importante: só `cache_persistence_options` não é suficiente pra
reaproveitar login entre processos — sem saber qual conta procurar no cache,
uma instância nova do credential não consegue tentar login silencioso e
sempre abre o navegador. Por isso aqui também salvamos um
`AuthenticationRecord` (identifica a conta autenticada) num arquivo local;
nas próximas execuções ele é recarregado e passado pro credential, permitindo
o lookup silencioso no cache.

Se quiser forçar um novo login (ex.: trocou de conta), apague a pasta:
    %LOCALAPPDATA%\\mkp-fabric-auth
"""

import os

import mssql_python
from azure.identity import (
    AuthenticationRecord,
    InteractiveBrowserCredential,
    TokenCachePersistenceOptions,
)

SERVER = os.environ.get(
    "MSSQL_SERVER",
    "2syk64xvbs6etmnq4gvqaftbom-lxq4dihtl5nuzpponrs3hni6cm.database.fabric.microsoft.com,1433",
)
DATABASE = os.environ.get(
    "MSSQL_DATABASE",
    "Marketplace-76c1ddc3-472f-4bf3-9445-ea7d2537f608",
)

_SCOPE = "https://database.windows.net/.default"

_CACHE_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "mkp-fabric-auth"
)
_RECORD_PATH = os.path.join(_CACHE_DIR, "auth_record.json")

_credential = None


def _carregar_authentication_record():
    if os.path.exists(_RECORD_PATH):
        with open(_RECORD_PATH, "r", encoding="utf-8") as f:
            conteudo = f.read().strip()
        if conteudo:
            return AuthenticationRecord.deserialize(conteudo)
    return None


def _salvar_authentication_record(record):
    os.makedirs(_CACHE_DIR, exist_ok=True)
    with open(_RECORD_PATH, "w", encoding="utf-8") as f:
        f.write(record.serialize())


def _obter_credential():
    global _credential
    if _credential is None:
        record = _carregar_authentication_record()
        _credential = InteractiveBrowserCredential(
            cache_persistence_options=TokenCachePersistenceOptions(name="mkp-fabric"),
            authentication_record=record,
        )
        if record is None:
            # Primeiro login (ou cache apagado): autentica explicitamente pra
            # capturar o AuthenticationRecord e salvar pra próxima execução.
            novo_record = _credential.authenticate(scopes=[_SCOPE])
            _salvar_authentication_record(novo_record)
    return _credential


def conectar(autocommit=False):
    """Abre uma conexão com o banco Marketplace no Fabric.

    token_provider substitui o `Authentication=ActiveDirectoryInteractive`
    da connection string (não podem ser usados juntos) — o InteractiveBrowserCredential
    já cobre o mesmo fluxo de login, só que com cache.
    """
    return mssql_python.connect(
        f"Server={SERVER};Database={{{DATABASE}}};Encrypt=yes;TrustServerCertificate=no;",
        token_provider=_obter_credential(),
        autocommit=autocommit,
    )

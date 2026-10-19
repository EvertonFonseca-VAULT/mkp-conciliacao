"""Guarda usuario/senha do Koncili e do SRI no Gerenciador de Credenciais do
Windows (biblioteca keyring), para nao precisar digitar `$env:...` toda vez
nem deixar senha escrita em arquivo nenhum.

O proprio Windows guarda a senha de forma criptografada, do mesmo jeito que
guarda a senha do Wi-Fi ou de um site salvo no Chrome. Nem o codigo deste
projeto, nem este script, nem ninguem olhando a tela consegue ler a senha de
volta -- so o koncili.py/Pagina.py pedem a senha pro Windows na hora de usar.

Uso (duplo clique em cofre_senhas.bat, ou no terminal):
    python cofre_senhas.py            # assistente: pergunta e grava
    python cofre_senhas.py --testar   # mostra o que ja esta gravado (sem revelar senha)
    python cofre_senhas.py --apagar koncili   # remove o que foi gravado (koncili ou sri)
    python cofre_senhas.py --teste    # autoteste offline (confere se o cofre funciona nesta maquina)
"""

import sys
from getpass import getpass

SERVICOS = {
    "koncili": "MARKETPLACE_KONCILI",
    "sri": "MARKETPLACE_SRI",
}


def salvar(servico, usuario, senha):
    import keyring
    keyring.set_password(servico, "usuario", usuario)
    keyring.set_password(servico, "senha", senha)


def consultar(servico):
    """Devolve (usuario, tem_senha) sem nunca devolver a senha em si."""
    import keyring
    usuario = keyring.get_password(servico, "usuario")
    tem_senha = keyring.get_password(servico, "senha") is not None
    return usuario, tem_senha


def apagar(servico):
    import keyring
    from keyring.errors import PasswordDeleteError
    apagado = False
    for chave in ("usuario", "senha"):
        try:
            keyring.delete_password(servico, chave)
            apagado = True
        except PasswordDeleteError:
            pass
    return apagado


def assistente():
    print("=== Cofre de senhas do MARKETPLACE ===")
    print("A senha digitada aqui NAO aparece na tela e fica guardada no Windows,")
    print("criptografada -- nunca em texto puro em arquivo nenhum.\n")
    print("Qual sistema voce quer cadastrar?")
    print("  1) Koncili")
    print("  2) SRI")
    print("  3) Os dois")
    escolha = input("Escolha (1, 2 ou 3): ").strip()
    nomes = {"1": ["koncili"], "2": ["sri"], "3": ["koncili", "sri"]}.get(escolha)
    if not nomes:
        sys.exit("Opcao invalida.")

    for nome in nomes:
        servico = SERVICOS[nome]
        print(f"\n--- {nome.upper()} ---")
        usuario = input("Usuario (email/login): ").strip()
        while True:
            senha = getpass("Senha (não aparece na tela): ")
            confirmacao = getpass("Digite a senha de novo pra confirmar: ")
            if senha == confirmacao and senha:
                break
            print("As senhas nao batem (ou vieram em branco). Tente de novo.")
        salvar(servico, usuario, senha)
        print(f"Gravado no cofre do Windows. Usuario salvo: {usuario}")

    print("\nPronto. Os scripts (koncili.py, Pagina.py) vao usar essas credenciais")
    print("sozinhos a partir de agora -- nao precisa mais do $env:... toda vez.")


def mostrar_status():
    print("=== O que esta gravado no cofre ===")
    for nome, servico in SERVICOS.items():
        usuario, tem_senha = consultar(servico)
        if usuario and tem_senha:
            print(f"{nome.upper()}: usuario '{usuario}', senha gravada.")
        elif usuario or tem_senha:
            print(f"{nome.upper()}: INCOMPLETO (falta usuario ou senha) -- rode o assistente de novo.")
        else:
            print(f"{nome.upper()}: nada gravado.")


def _teste():
    """Offline: grava, le e apaga um usuario/senha de teste -- confere que o
    cofre do Windows funciona nesta maquina, sem tocar nas credenciais reais."""
    servico_teste = "MARKETPLACE_TESTE_AUTOTESTE"
    salvar(servico_teste, "usuario_teste", "senha_teste_123")
    usuario, tem_senha = consultar(servico_teste)
    assert usuario == "usuario_teste" and tem_senha, "nao conseguiu ler de volta o que acabou de gravar"
    assert apagar(servico_teste), "nao conseguiu apagar o teste"
    usuario, tem_senha = consultar(servico_teste)
    assert usuario is None and not tem_senha, "apagar nao limpou de verdade"
    print("cofre_senhas.py: ok (o cofre do Windows funciona nesta maquina)")


if __name__ == "__main__":
    if "--teste" in sys.argv:
        _teste()
    elif "--testar" in sys.argv:
        mostrar_status()
    elif "--apagar" in sys.argv:
        idx = sys.argv.index("--apagar")
        nome = sys.argv[idx + 1].lower() if len(sys.argv) > idx + 1 else ""
        servico = SERVICOS.get(nome)
        if not servico:
            sys.exit(f"uso: python cofre_senhas.py --apagar <{'|'.join(SERVICOS)}>")
        print("Removido." if apagar(servico) else "Nao havia nada gravado.")
    else:
        assistente()

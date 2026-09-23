# Aula básica: ambiente virtual (venv)

## 1. Para que serve

Um **venv** é uma pasta (`.venv`) com um Python e bibliotecas só deste projeto.

- O projeto usa `pandas`, `openpyxl`, `playwright` etc. em versões específicas (`requirements.txt`). Instaladas no venv, não brigam com as de outros projetos nem com o Python do Windows.
- Se o venv quebrar, basta apagar a pasta `.venv` e criar de novo — nada do sistema é afetado.
- Todo mundo do time roda com as mesmas bibliotecas, então "na minha máquina funciona" vale para todos.

A pasta `.venv` **não vai para o git** (está no `.gitignore`): ela é recriada em cada máquina a partir do `requirements.txt`.

## 2. Criar (uma vez só, na pasta `Mkp`)

```bash
python -m venv .venv
```

## 3. Ativar (toda vez que abrir um terminal novo)

| Terminal | Comando |
|---|---|
| PowerShell | `.venv\Scripts\Activate.ps1` |
| Prompt de Comando (cmd) | `.venv\Scripts\activate.bat` |
| Git Bash / Linux / Mac | `source .venv/Scripts/activate` (Linux/Mac: `.venv/bin/activate`) |

Deu certo quando aparece `(.venv)` no começo da linha:

```
(.venv) PS C:\Users\...\Mkp>
```

Para conferir qual Python está em uso:

```bash
python -c "import sys; print(sys.executable)"   # deve apontar para ...\Mkp\.venv\Scripts\python.exe
```

**Erro no PowerShell** "a execução de scripts foi desabilitada neste sistema": rode uma vez

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

e ative de novo.

## 4. Instalar as dependências (com o venv ativo)

```bash
pip install -r requirements.txt
playwright install chromium
```

Instalou uma biblioteca nova que o projeto passou a usar? Adicione no `requirements.txt` para o resto do time.

## 5. Usar

Com `(.venv)` na linha, é só rodar os scripts normalmente (`python AB\interativo.py` etc.) — o `python` e o `pip` passam a ser os do venv.

Sem ativar, dá para chamar o Python do venv direto:

```bash
.venv\Scripts\python.exe AB\interativo.py
```

## 6. Desativar

```bash
deactivate
```

O `(.venv)` some e o terminal volta ao Python do sistema. Fechar o terminal também desativa.

## 7. Usuário e senha (credenciais do Koncili)

O `koncili.py` (e o `koncili_shopee.py`) faz login no Koncili com o **seu** usuário e senha, lidos de duas variáveis de ambiente:

- `KONCILI_USUARIO` — o e-mail de login
- `KONCILI_SENHA` — a senha

Se faltar alguma, ele para logo no início com:

```
ERRO: defina a(s) variavel(is) de ambiente: KONCILI_USUARIO, KONCILI_SENHA
```

### Por que é necessário (e por que não fica no código)

- **O repositório está no GitHub.** Uma senha escrita num `.py` fica no histórico do git **para sempre** — mesmo apagando depois, qualquer um com acesso ao repositório encontra nos commits antigos.
- **Cada pessoa usa o próprio login.** O mesmo código roda em várias máquinas; a variável de ambiente faz cada um usar a sua conta sem mexer em nenhum arquivo versionado.
- **Trocar a senha não exige commit.** Mudou a senha no Koncili? Só atualiza a variável.
- A pasta `.venv` **está no `.gitignore`**, então o que for guardado nela (passo B abaixo) nunca sobe para o GitHub.

### A) Só para o terminal atual (some ao fechar)

```powershell
$env:KONCILI_USUARIO = "seu.email@empresa.com"
$env:KONCILI_SENHA   = "sua-senha"
python AB\koncili.py ES 18.09.2026
```

No cmd: `set KONCILI_USUARIO=seu.email@empresa.com` e `set KONCILI_SENHA=sua-senha`.

### B) Carregar sozinho ao ativar o venv (recomendado)

Adicione as linhas **no final** do script de ativação do venv que você usa para o Koncili (o `koncili.py` roda com o venv do **Screaper Mercado**, que tem o playwright; se usar o venv do `Mkp`, é o `Mkp\.venv`):

- PowerShell — final de `.venv\Scripts\Activate.ps1`:

  ```powershell
  $env:KONCILI_USUARIO = "seu.email@empresa.com"
  $env:KONCILI_SENHA   = "sua-senha"
  ```

- cmd — final de `.venv\Scripts\activate.bat`:

  ```bat
  set "KONCILI_USUARIO=seu.email@empresa.com"
  set "KONCILI_SENHA=sua-senha"
  ```

Pronto: toda vez que ativar o venv, o login já está definido. Para conferir (sem mostrar a senha):

```powershell
python -c "import os; print(os.environ.get('KONCILI_USUARIO'), 'senha OK' if os.environ.get('KONCILI_SENHA') else 'SEM SENHA')"
```

Cuidados:

- O `deactivate` **não** apaga essas duas variáveis — elas ficam até fechar o terminal. Para limpar na hora: `Remove-Item Env:KONCILI_USUARIO, Env:KONCILI_SENHA`.
- Apagou e recriou a `.venv`? O `Activate.ps1` é gerado de novo — adicione as linhas outra vez.
- A senha fica em texto puro nesse arquivo, dentro do seu usuário do Windows. Não copie a pasta `.venv` para outra pessoa nem para a unidade `X:`.
- **Nunca** escreva a senha em arquivo `.py`, `.bat` versionado, README ou mensagem de commit.

## Resumo

```bash
python -m venv .venv            # criar (1x)
.venv\Scripts\Activate.ps1      # ativar (PowerShell)
pip install -r requirements.txt # instalar (1x, ou quando o requirements mudar)
deactivate                      # desativar
$env:KONCILI_USUARIO = "..."    # login do Koncili (ou no final do Activate.ps1)
$env:KONCILI_SENHA   = "..."
```

Observação: o scraper usa o venv do projeto **Screaper Mercado** (ver Setup no README); o `interativo.py --scraper` troca para ele sozinho.

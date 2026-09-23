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

## Resumo

```bash
python -m venv .venv            # criar (1x)
.venv\Scripts\Activate.ps1      # ativar (PowerShell)
pip install -r requirements.txt # instalar (1x, ou quando o requirements mudar)
deactivate                      # desativar
```

Observação: o scraper usa o venv do projeto **Screaper Mercado** (ver Setup no README); o `interativo.py --scraper` troca para ele sozinho.

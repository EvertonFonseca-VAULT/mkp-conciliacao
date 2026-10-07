# App Desktop — Conciliação Mercado

**Data:** 2026-10-07  
**Projeto:** mkp-conciliacao  
**Autor:** Everton Fonseca

---

## Intenção

Substituir os `.bat` e scripts avulsos por um programa desktop com até 3 telas que guia o operador pelo fluxo diário completo: extrair AB do Koncili → gerar Para Processar → empilhar → conciliar com dashboard. O objetivo é que um colega novo consiga operar todos os canais do Mercado sem saber qual script chamar, sem navegar por pastas ou abrir terminal.

**Sucesso:** um colega recebe o `.exe`, configura credenciais, e conclui o processo completo do dia sem precisar do README.

---

## Escopo

**Inclui:**
- Tela 1: hub operacional (canal, data, 4 botões, log em tempo real)
- Tela 2: conciliação (upload RBs, toggle scraper, dashboard HTML embutido)
- Tela 3: configurações (credenciais Koncili, caminhos editáveis)
- Distribuição como `.exe` via PyInstaller

**Não inclui (por agora):**
- Reescrita do dashboard HTML (continua sendo gerado pelos scripts existentes)
- Gestão de histórico / banco de dados próprio
- Autenticação de usuários do app
- Suporte a múltiplos ambientes simultâneos

---

## As 3 Telas

### Tela 1 — Operação (tela inicial)

Abre ao iniciar o programa. Contém:

- **Seletor de canal** (dropdown): FULL, ES, MOVEIS, ELETRO, CENTER, CO05, PRIME, WEB
- **Seletor de data** (campo dd.mm.aaaa com validação)
- **4 botões de ação** em grade 2×2:
  1. `Extrair AB` — chama `koncili.py conferir` para o canal/data; salva o AB nas pastas corretas
  2. `Para Processar` — chama `rodar_canal.py processar` para o canal/data
  3. `Empilhar` — chama o empilhador correto para o canal (ver roteamento abaixo)
  4. `Conciliar →` — abre a Tela 2 com canal/data já preenchidos
- **Área de log** (somente leitura, scroll automático): output dos scripts linha a linha em tempo real
- **Indicadores de status** por etapa: cinza (não rodado) → amarelo (rodando) → verde (ok) → vermelho (erro)
- **Ícone de engrenagem** no canto superior direito: abre Tela 3

**Roteamento do empilhador por canal:**

| Canal | Módulo chamado |
|---|---|
| FULL, ES | `empilhar.main` (com `--publicar`) |
| CENTER | `empilhar_center.main` |
| WEB | `empilhar_web.main` |
| PRIME | `empilhar_prime.main` |
| MOVEIS | `empilhar_moveis.main` |
| CO05 | `empilhar_coletas05.main` |
| ELETRO | avisa que empilhador não existe ainda |

---

### Tela 2 — Conciliação

Abre ao clicar em `Conciliar →`. Exibe canal e data no título. Contém:

- **Upload de arquivos RB** (dois campos com botão "Selecionar arquivo"):
  - `rb.xlsx` → copiado para `<pasta_saida>\AUXILIAR\rb.xlsx`
  - `rbx.xlsx` → copiado para `<pasta_saida>\AUXILIAR\rbx.xlsx`
  - Indicador visual: ✓ ok / ✗ faltando por arquivo
- **Checkbox "Usar scraper ML/MP"** (desmarcado por padrão)
- **Botão `▶ Iniciar Conciliação`** — habilitado apenas quando os dois arquivos RB estiverem presentes; chama `interativo.py`
- **Área do dashboard** (QWebEngineView, ocupa a metade inferior da tela): carrega o HTML gerado após a conciliação rodar; recarrega automaticamente quando o arquivo é atualizado (mesmo comportamento atual do `visibilitychange`)
- **Botão `← Voltar`** no canto superior esquerdo: volta para a Tela 1 sem perder a seleção de canal/data

---

### Tela 3 — Configurações

Abre pelo ícone de engrenagem. Contém:

**Bloco Koncili:**
- Campo `Usuário` (texto)
- Campo `Senha` (texto mascarado)
- Botão `Salvar` — persiste via `keyring` (cofre do Windows); nunca em arquivo de texto

**Bloco Caminhos** (expansível, colapsado por padrão — maioria não precisa mudar):
- `Unidade X:` (padrão `X:\`)
- `RODAR TOTVS` (padrão `C:\RODAR TOTVS`)
- `Pasta de saída da conciliação` (padrão `\\<servidor>\Conciliacoes` ou o padrão atual do `interativo.py`)
- Botão `Restaurar padrões`

**Botão `← Voltar`** no canto superior esquerdo.

---

## Arquitetura Técnica

### Estrutura de arquivos

```
Mkp/
├── app/
│   ├── main.py               ← QApplication, carrega MainWindow
│   ├── main_window.py        ← QMainWindow + QStackedWidget (3 telas)
│   ├── tela_operacao.py      ← Tela 1
│   ├── tela_conciliacao.py   ← Tela 2 + QWebEngineView
│   ├── tela_config.py        ← Tela 3
│   ├── worker.py             ← WorkerThread (QThread) genérico
│   └── config.py             ← leitura/escrita de caminhos (QSettings) e credenciais (keyring)
├── AB/                       ← código existente, sem modificações
├── docs/
└── requirements.txt
```

### Threading

Toda operação que chama um script roda em um `WorkerThread(QThread)`. A UI nunca trava.

```
WorkerThread
  ├── sinal: line_emitted(str)   → log da Tela 1 (append em tempo real)
  ├── sinal: finished(bool)      → atualiza indicador de status
  └── sinal: error(str)          → exibe mensagem de erro
```

Dois modos de execução dentro do worker:

**Import direto** (scripts Python puros sem Playwright):
- `rodar_canal`, `empilhar*`, `interativo` sem `--scraper`
- Redireciona `sys.stdout` para capturar output linha a linha via `io.StringIO` + thread de leitura
- Chama `main(argv)` do módulo importado

**Subprocess** (scripts com Playwright):
- `koncili.py` (login Playwright no Koncili)
- `interativo.py` com `--scraper` (abre Chrome via debug remoto)
- `subprocess.Popen(stdout=PIPE)` + leitura linha a linha em thread separada
- Usa o mesmo Python do venv atual (`sys.executable`)

### Credenciais

`keyring.set_password("mkp-conciliacao", "KONCILI_USUARIO", valor)` e idem para senha.  
Antes de chamar qualquer script do Koncili, o app injeta as credenciais como variáveis de ambiente do processo filho:

```python
env = os.environ.copy()
env["KONCILI_USUARIO"] = keyring.get_password("mkp-conciliacao", "KONCILI_USUARIO")
env["KONCILI_SENHA"]   = keyring.get_password("mkp-conciliacao", "KONCILI_SENHA")
```

### Persistência de caminhos

`QSettings("WebContinental", "MkpConciliacao")` — salvo no registro do Windows. Valores padrão definidos em `config.py`; editáveis na Tela 3.

### Distribuição

```bash
pyinstaller --onefile --windowed --name "Conciliacao Mercado" app/main.py
```

O colega recebe `Conciliacao Mercado.exe`. Sem instalar Python, sem ativar venv, sem terminal. Na primeira abertura, o app detecta credenciais ausentes e abre a Tela 3 automaticamente.

### Novas dependências (additions ao requirements.txt)

```
PyQt6
PyQt6-WebEngine
keyring
pyinstaller  # só build, não runtime
```

---

## Fluxo completo do dia (perspectiva do operador)

```
1. Abre o .exe
2. Seleciona canal e data na Tela 1
3. Clica "Extrair AB"   → log mostra progresso → indicador fica verde
4. Clica "Para Processar" → idem
5. Clica "Empilhar"       → idem
6. Clica "Conciliar →"    → Tela 2 abre
7. Seleciona rb.xlsx e rbx.xlsx
8. (Opcional) marca "Usar scraper"
9. Clica "▶ Iniciar Conciliação"
10. Dashboard aparece na metade inferior da tela
```

---

## Critérios de sucesso

- Colega sem conhecimento dos scripts opera o fluxo completo do dia sem abrir terminal
- Credenciais salvas no cofre do Windows (nenhuma senha em texto claro)
- Dashboard HTML existente aparece intacto dentro do app (sem reescrita)
- `.exe` distribuível sem dependências externas
- Nenhum arquivo em `AB/` modificado (o app só importa/chama o que existe)

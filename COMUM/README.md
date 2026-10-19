# COMUM — peças compartilhadas e cofre de senhas

Esta pasta não roda sozinha no dia a dia. Ela guarda duas coisas usadas pelas outras pastas:

- `calendario_alvo.py` — sabe, pra cada canal e cada dia, qual é o valor que **deveria** ter vindo do Mercado (o "alvo" do calendário de planejamento). O Koncili, o Empilhar e a Conciliação usam essa informação pra saber se o dia fechou certo.
- `cofre_senhas.py` / `cofre_senhas.bat` — onde você guarda, **uma única vez**, o usuário e a senha do Koncili e do SRI.

Esta segunda parte é a que toda pessoa nova precisa fazer antes de usar o resto do projeto. Veja abaixo.

## Como guardar as senhas (primeira vez, ou quando a senha mudar)

Os programas deste projeto (`KONCILI/koncili.py` e o scraper do SRI) precisam de um usuário e senha pra funcionar. Em vez de digitar isso toda vez que for rodar algo, você guarda a senha **uma vez só**, de forma segura, no **Gerenciador de Credenciais do Windows** — o mesmo cofre onde o Windows guarda, por exemplo, a senha do Wi-Fi ou uma senha de rede salva. A senha fica criptografada pelo próprio Windows; nem este projeto, nem quem olhar o código, consegue ler ela de volta.

### Passo a passo

1. Abra a pasta `COMUM` no Explorador de Arquivos do Windows.
2. Dê **duplo clique** em `cofre_senhas.bat`.
3. Vai abrir uma janela preta (terminal) perguntando qual sistema você quer cadastrar:
   - `1` para Koncili
   - `2` para SRI
   - `3` para os dois de uma vez (recomendado, se você usa os dois)
4. Digite o número e aperte Enter.
5. Para cada sistema, ele vai pedir:
   - **Usuário** (o email/login que você usa pra entrar no Koncili ou no SRI) — aparece na tela normalmente, pode digitar e conferir.
   - **Senha** — ao digitar, **nada aparece na tela** (nem ponto, nem asterisco). Isso é normal e proposital, é assim que se evita que alguém "olhando por cima do ombro" veja a senha. Digite e aperte Enter.
   - Ele pede a senha **de novo**, só pra confirmar que você não digitou errado.
6. Se aparecer "Gravado no cofre do Windows", deu certo. Pode fechar a janela.

A partir daí, `koncili.py` (pasta `KONCILI`) e `Pagina.py` (pasta `Screaper Mercado`) vão achar a senha sozinhos, sem você precisar fazer mais nada.

### Como saber o que já está gravado

Duplo clique não serve aqui (ia abrir o assistente de novo). Abra um terminal na pasta `COMUM` e rode:

```
..\.venv\Scripts\python.exe cofre_senhas.py --testar
```

Isso mostra, para Koncili e para SRI, se tem usuário/senha gravados — **sem nunca mostrar a senha em si**, só se ela existe ou não.

### Se a senha mudou (trocou a senha do Koncili, por exemplo)

Rode o assistente de novo (`cofre_senhas.bat`) e cadastre de novo — ele substitui o que estava gravado, não precisa apagar antes.

### Se quiser apagar o que está gravado

```
..\.venv\Scripts\python.exe cofre_senhas.py --apagar koncili
..\.venv\Scripts\python.exe cofre_senhas.py --apagar sri
```

### Alternativa (não recomendada para quem não mexe com código): variável de ambiente

Quem prefere, pode digitar a senha a cada vez, numa janela do PowerShell, sem gravar em lugar nenhum:

```powershell
$env:KONCILI_USUARIO="seu-email"
$env:KONCILI_SENHA="sua-senha"
```

Essa variável só vale **naquela janela aberta** — fecha o PowerShell, esquece a senha. Os scripts sempre dão prioridade à variável de ambiente, se ela estiver definida; só vão olhar no cofre se ela não estiver.

### Testar se o cofre está funcionando neste computador

```
..\.venv\Scripts\python.exe cofre_senhas.py --teste
```

Isso grava, lê e apaga uma credencial de teste (não mexe nas suas senhas reais) e confirma que o Gerenciador de Credenciais do Windows está respondendo direito nesta máquina. Deve terminar com `cofre_senhas.py: ok`.

## Por que é seguro

- A senha nunca fica escrita em nenhum arquivo deste projeto, nem em texto, nem em planilha.
- Quem olhar o código-fonte não encontra nenhuma senha nele — os scripts só perguntam ao Windows "me dá a senha salva pro serviço tal", na hora de usar.
- A pasta de dados (`.xlsx`, `.xls`, `.db`) também nunca é enviada pra lugar nenhum fora deste computador — veja `.gitignore` na raiz do projeto.

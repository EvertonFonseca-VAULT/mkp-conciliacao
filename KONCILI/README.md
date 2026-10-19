# KONCILI — baixar o repasse do dia

Esta é a **etapa 1** do dia. O programa entra no site do Koncili (que recebe os repasses do Mercado Livre), acha o repasse do canal e da data que você pedir, confere o valor (RESULTADO) contra o calendário de planejamento, e baixa o arquivo original (o "AB") pra pasta certa.

## Antes de rodar pela primeira vez

1. Configure o cofre de senhas — veja `../COMUM/README.md`. Sem isso, o programa para com um erro pedindo usuário/senha.
2. Confirme que você tem acesso à unidade de rede `X:` (é lá que o AB é gravado quando você usa `--salvar`).
3. Precisa de internet e acesso ao site do Koncili.

## Como rodar

Abra um terminal (PowerShell) dentro da pasta `KONCILI` e rode um dos comandos abaixo, trocando `ES` pelo canal que você quer (`ES`, `ELETRO`, `MOVEIS`, `PRIME`, `FULL`, `CENTER`, `WEB`, `COLETAS_CO05`) e a data pelo dia desejado, no formato `DD.MM.AAAA`:

```
..\.venv\Scripts\python.exe koncili.py ES 22.09.2026
```

Isso **só confere** o valor — não grava nada na pasta `X:`. É o jeito seguro de "só dar uma olhada".

Para **gravar de verdade** o AB na pasta do `X:` (sem sobrescrever o que já existe):

```
..\.venv\Scripts\python.exe koncili.py ELETRO 22.09.2026 --salvar
```

Para conferir **vários dias de uma vez** (lista cada data, não é "de tal dia até tal dia"):

```
..\.venv\Scripts\python.exe koncili.py FULL 11.09.2026 14.09.2026 18.09.2026 --salvar
```

Para conferir um **intervalo** de datas sem digitar uma por uma:

```
..\.venv\Scripts\python.exe koncili.py FULL 11.09.2026 20.09.2026 --intervalo --salvar
```

## O que a tela mostra

Para cada data processada, o programa avisa se o valor do AB bate com o RESULTADO do Koncili e com o alvo do calendário. Se não bater, aparece um aviso — é sinal de olhar com mais calma antes de seguir pra próxima etapa (`PARA_PROCESSAR`).

## Erros comuns

| Mensagem / sintoma | O que fazer |
|---|---|
| `ERRO: defina a(s) variavel(is) de ambiente ou grave no cofre` | Configure o cofre de senhas em `../COMUM/README.md`, ou defina `$env:KONCILI_USUARIO` / `$env:KONCILI_SENHA` naquela janela. |
| `login nao confirmado: continua em /auth` | Usuário ou senha errados — grave de novo no cofre com `cofre_senhas.bat`. |
| Demora muito ou trava | Confira sua internet. O Koncili às vezes demora a responder; não é bug. |
| Não encontra a pasta `X:` ao usar `--salvar` | A VPN pode estar caída, ou o acesso à unidade de rede não está montado. Confira com o time de TI. |
| `python` não é reconhecido / comando não funciona | Você precisa usar o Python de dentro do projeto, não o do Windows: `..\.venv\Scripts\python.exe`, exatamente como nos exemplos acima. |

## Testar sem mexer em nada real

```
..\.venv\Scripts\python.exe koncili.py --teste
```

Esse comando **não acessa a internet nem grava nada** — ele só confere se a lógica interna do programa (como montar o nome do arquivo, como calcular o valor, como ler as credenciais) continua funcionando. Deve terminar com `koncili.py: ok`. É um bom primeiro teste depois de qualquer atualização do projeto, antes de usar pra valer.

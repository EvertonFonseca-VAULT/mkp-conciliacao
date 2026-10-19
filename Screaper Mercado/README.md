# Screaper Mercado — busca de pedido no SRI, ML e MP

Esta pasta é uma ferramenta auxiliar, usada por dentro das outras etapas (`PARA_PROCESSAR` e `CONCILIACAO`) e também pode ser usada sozinha para procurar um pedido específico.

- `Pagina.py` — entra no SRI (sistema interno) e busca pedidos listados numa planilha, preenchendo os dados encontrados de volta na planilha.
- `Screaper.py` — usa uma janela do Chrome já aberta (ver `abrir_chrome_debug.bat`, na pasta `CONCILIACAO`) pra buscar detalhes de pedidos no Mercado Livre e no Mercado Pago.

## Antes de rodar: senha do SRI

`Pagina.py` precisa de usuário e senha do SRI. Configure isso **uma vez** no cofre de senhas — veja o passo a passo em `../COMUM/README.md` (escolha a opção "SRI" no assistente).

## Como rodar `Pagina.py`

1. Prepare a planilha de entrada em `C:\Users\<seu usuário>\Desktop\Projetos\PROCURAR.xlsx`, com os pedidos a buscar na coluna B.
2. Rode:
   ```
   ..\.venv\Scripts\python.exe Pagina.py
   ```
3. Ao final, aparece uma janela avisando quantos pedidos foram processados e quantos foram encontrados. A própria planilha `PROCURAR.xlsx` é atualizada com o resultado.

## Como rodar `Screaper.py`

1. Abra o Chrome especial com `abrir_chrome_debug.bat` (está na pasta `CONCILIACAO`) e faça login no Mercado Livre e no Mercado Pago, se pedir.
2. Prepare a planilha de entrada em `C:\Users\<seu usuário>\Desktop\Projetos\AJUSTES.xlsx`, com o número do pedido na coluna A.
3. Rode:
   ```
   ..\.venv\Scripts\python.exe Screaper.py
   ```
4. O resultado é salvo em `CONSULTAS_ML.xlsx` (pedidos do Mercado Livre) ou `CONSULTAS_MP.xlsx` (Mercado Pago), na mesma pasta `Desktop\Projetos`.

## Erros comuns

| Mensagem / sintoma | O que fazer |
|---|---|
| `ERRO: defina a(s) variavel(is) de ambiente ou grave no cofre (MARKETPLACE_SRI)` | Configure a senha do SRI no cofre — `../COMUM/README.md`. |
| `PULADO (nenhuma aba do Mercado Pago aberta)` | Abra uma aba do Mercado Pago no Chrome aberto por `abrir_chrome_debug.bat` antes de rodar `Screaper.py`. |
| `Nada encontrado` para um pedido | O pedido pode não existir naquele sistema, ou o número foi digitado errado na planilha de entrada. |

Esta pasta não tem autoteste próprio (ela só automatiza cliques e leitura de tela — não tem lógica de cálculo pra testar offline). A forma de validar é rodar com um pedido que você já sabe o resultado esperado.

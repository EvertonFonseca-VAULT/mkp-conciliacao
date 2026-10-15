"""Gera a lista de lancamentos de REBATE prontos pra colar na aba
Movimento_Bancario da planilha CMM do canal (ex.: CMM_MELI_ELETROD...xlsx).
Mesma ordem de colunas dessa aba -- ver estrutura real conferida em
CMM_MELI_ELETROD_Controle de Movimentos Manual 2026.xlsx.

So cobre o caso REBATE (o unico com convencao 100% fixa confirmada ate
agora -- Manuais/Passo a passo - Resolucao manual de pendencias (Baixas
no TOTVS).md). LOJA/VENDEDOR/HISTORICO vem de CANAIS[canal] (loja_cmm/
vendedor/prefixo_mov) -- confirmados com dado real do FULL (LOJA=FULL,
VENDEDOR=751) e do ES (LOJA=CO, VENDEDOR=763, HISTORICO=
"MKP_MELI_CO_RECEBIMENTO DE REBATE" = prefixo_mov + "_" + MOTIVO_REBATE).
Pedido do Everton, 18/09/2026.
"""

import pandas as pd

from canais import CANAIS

NATUREZA_REBATE = 1107
MOTIVO_REBATE = "RECEBIMENTO DE REBATE"

COLUNAS = [
    "RESPONSAVEL", "DATA_LACTO", "ORIGEM", "PAGADOR", "LOJA",
    "MOTIVO DO MOVIMENTO", "NATUREZA", "PEDIDO MKP / CV", "VENDEDOR",
    "BANCO", "AGENCIA", "CONTA", "HISTORICO", "VALOR R$", "OBSERVACAO",
    "PEDIDO_ERP", "PEDIDO_SITE", "N. FISCAL", "PROCESSO",
]


def _banco_agencia_conta_por_pedido(rb):
    """Conta e sempre a transitoria (com o prefixo "X") -- pega direto do RB,
    que ja traz BANCO/AGENCIA/CONTA por lancamento (rastreado por pagina do
    extrato em parser_rb.py). Usa a moda do arquivo inteiro como fallback
    pros pedidos que por algum motivo nao tem lancamento no RB ainda."""
    banco_cod = rb["BANCO"].astype(str).str.split(" - ").str[0].str.strip()
    referencia = pd.DataFrame({
        "PED_MKP": rb["PED_MKP"], "BANCO": banco_cod,
        "AGENCIA": rb["AGENCIA"], "CONTA": rb["CONTA"],
    })
    por_pedido = referencia[referencia["PED_MKP"] != ""].drop_duplicates("PED_MKP").set_index("PED_MKP")
    fallback = referencia.mode().iloc[0] if len(referencia) else pd.Series({"BANCO": None, "AGENCIA": None, "CONTA": None})
    return por_pedido, fallback


def gerar_rebate_cmm(canal, resultado, raiz, data_ref, responsavel, rb):
    cfg = CANAIS[canal]
    loja = cfg["loja_cmm"]
    vendedor = cfg["vendedor"]
    historico = f"{cfg['prefixo_mov']}_{MOTIVO_REBATE}"
    por_pedido, fallback = _banco_agencia_conta_por_pedido(rb)

    rebates = resultado[resultado["STATUS"] == "REBATE"]

    linhas = []
    for _, row in rebates.iterrows():
        pedido = row["PEDIDO_MKP"]
        conta_info = por_pedido.loc[pedido] if pedido in por_pedido.index else fallback
        linhas.append({
            "RESPONSAVEL": responsavel,
            "DATA_LACTO": data_ref,
            "ORIGEM": "MKP",
            "PAGADOR": "MELI",
            "LOJA": loja,
            "MOTIVO DO MOVIMENTO": MOTIVO_REBATE,
            "NATUREZA": NATUREZA_REBATE,
            "PEDIDO MKP / CV": pedido,
            "VENDEDOR": vendedor,
            "BANCO": conta_info["BANCO"],
            "AGENCIA": conta_info["AGENCIA"],
            "CONTA": conta_info["CONTA"],
            "HISTORICO": historico,
            "VALOR R$": row["REBATE"],
            "OBSERVACAO": None,
            "PEDIDO_ERP": None,
            "PEDIDO_SITE": None,
            "N. FISCAL": None,
            "PROCESSO": None,
        })

    return pd.DataFrame(linhas, columns=COLUNAS)


if __name__ == "__main__":
    resultado = pd.DataFrame({
        "PEDIDO_MKP": ["1"], "STATUS": ["REBATE"], "REBATE": [30.0],
    })
    rb = pd.DataFrame({
        "PED_MKP": ["1"], "BANCO": ["237 - BRADESCO SA"], "AGENCIA": ["3708"], "CONTA": ["X103925"],
    })
    df = gerar_rebate_cmm("FULL", resultado, "MKP_MELI_FU", "18/09/2026", "EVERTON", rb)
    assert (df["LOJA"] == "FULL").all()  # nao "FU" (chave_rbx, campo errado)
    assert (df["VENDEDOR"] == "751").all()
    assert (df["HISTORICO"] == "MKP_ME_FULL_RECEBIMENTO DE REBATE").all()

    df_es = gerar_rebate_cmm("ES", resultado, "MKP_MELI_CO_ES", "18/09/2026", "EVERTON", rb)
    assert (df_es["LOJA"] == "CO").all()
    assert (df_es["VENDEDOR"] == "763").all()
    assert (df_es["HISTORICO"] == "MKP_MELI_CO_RECEBIMENTO DE REBATE").all()  # confirmado com dado real do ES
    print("rebate_cmm.py: ok")

"""Le o arquivo de RESULTADO que o TOTVS gera depois de importar o
PROCESSAR_REBATE (ex. marketplace20260915174306.xls -- baixado pelo Everton)
e monta o CMM dos pedidos que NAO baixaram sozinhos ("Fora da Regra"), pra
lancar manualmente igual Frete/Tarifa. Pedido do Everton, 16/09/2026.

Apesar da extensao .xls o conteudo e SpreadsheetML/XML -- openpyxl e
pandas.read_excel nao leem esse formato (erro "cannot be determined"), por
isso o parse e via xml.etree.ElementTree. Estrutura conferida num arquivo
real: sheets "1-credito"/"2-debito" (a "3-Banco" so confirma Banco/Agencia/
Conta, ignorada aqui), linha 1 e so o nome da sheet solta, cabecalho na
linha 2, dados a partir da linha 3. O numero de colunas varia entre as 2
sheets, mas em ambas a 1a celula e sempre o pedido (External), a 3a e
sempre o Valor, e a ULTIMA e sempre a Acao.
"""

import xml.etree.ElementTree as ET

import pandas as pd

from rebate_cmm import COLUNAS, NATUREZA_REBATE, MOTIVO_REBATE

_NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
_SHEETS_DADOS = ("1-credito", "2-debito")
_ACAO_FORA_DA_REGRA = "Fora da Regra"

# Conta transitoria confirmada na sheet "3-Banco" do arquivo real -- mesma
# conta de sempre (nao varia por pedido, diferente do fluxo do RB em
# rebate_cmm.py, que tem o extrato pra descobrir por lancamento).
BANCO, AGENCIA, CONTA = "237", "3708", "X103925"

# Arquivo e 100% hardcoded pro canal FULL (mesmo espirito do "MERCADO FULL"
# fixo em rebate_processar.py). HISTORICO usa prefixo_mov do FULL
# (CANAIS["FULL"]["prefixo_mov"] = "MKP_ME_FULL", ver canais.py e
# dashboard.py:450) -- ja confirmado com rotulo real do rbx.xlsx do FULL,
# diferente de rebate_cmm.py (generico pra qualquer canal, la fica em
# branco por nao estar confirmado).
HISTORICO_REBATE = "MKP_ME_FULL_" + MOTIVO_REBATE


def ler_fora_da_regra(caminho_xls):
    """[(pedido, valor)] das linhas com Acao == "Fora da Regra" nas sheets
    de dados. Assume sem ss:Index esparso nas Cell (confirmado no arquivo
    real) -- se algum export vier com celulas puladas isso desalinha."""
    root = ET.parse(caminho_xls).getroot()
    saida = []
    for ws in root.findall("ss:Worksheet", _NS):
        if ws.get("{urn:schemas-microsoft-com:office:spreadsheet}Name") not in _SHEETS_DADOS:
            continue
        for linha in ws.find("ss:Table", _NS).findall("ss:Row", _NS)[2:]:
            celulas = [c.find("ss:Data", _NS) for c in linha.findall("ss:Cell", _NS)]
            valores = [d.text if d is not None else None for d in celulas]
            if not valores or valores[-1] != _ACAO_FORA_DA_REGRA:
                continue
            saida.append((valores[0], float(valores[2])))
    return saida


def gerar_rebate_fora_da_regra_cmm(caminho_xls, raiz, data_ref, responsavel):
    loja = raiz.split("_")[-1]
    linhas = [{
        "RESPONSAVEL": responsavel,
        "DATA_LACTO": data_ref,
        "ORIGEM": "MKP",
        "PAGADOR": "MELI",
        "LOJA": loja,
        "MOTIVO DO MOVIMENTO": MOTIVO_REBATE,
        "NATUREZA": NATUREZA_REBATE,
        "PEDIDO MKP / CV": pedido,
        "VENDEDOR": None,
        "BANCO": BANCO,
        "AGENCIA": AGENCIA,
        "CONTA": CONTA,
        "HISTORICO": HISTORICO_REBATE,
        "VALOR R$": valor,
        "OBSERVACAO": None,
        "PEDIDO_ERP": None,
        "PEDIDO_SITE": None,
        "N. FISCAL": None,
        "PROCESSO": None,
    } for pedido, valor in ler_fora_da_regra(caminho_xls)]
    return pd.DataFrame(linhas, columns=COLUNAS)


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        caminho_xls, raiz, data_ref, responsavel = sys.argv[1:5]
        df = gerar_rebate_fora_da_regra_cmm(caminho_xls, raiz, data_ref, responsavel)
        caminho_saida = caminho_xls.rsplit(".", 1)[0] + "_cmm.xlsx"
        df.to_excel(caminho_saida, index=False)
        print(f"{len(df)} pedido(s) fora da regra -> {caminho_saida}")
        sys.exit(0)

    import os
    import tempfile

    def _cel(valor, tipo="String"):
        return f'<Cell><Data ss:Type="{tipo}">{valor}</Data></Cell>'

    def _linha(*valores):
        return "<Row>" + "".join(_cel(v, "Number" if isinstance(v, float) else "String") for v in valores) + "</Row>"

    # Reproduz a estrutura real: titulo solto (linha 1), cabecalho (linha 2),
    # dados a partir da linha 3. Credito com 5 colunas, debito com 4 (testa
    # que a diferenca de contagem de colunas entre sheets nao quebra o parse
    # -- so index 0/2/-1 importam). Pedidos e valores dos "Fora da Regra"
    # reais que o Everton conferiu no arquivo marketplace20260915174306.xls.
    _XML = f"""<?xml version="1.0"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
          xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
  <Worksheet ss:Name="1-credito">
    <Table>
      <Row>{_cel("1-credito")}</Row>
      {_linha("External", "Tipo", "Valor", "ColD", "Acao")}
      {_linha("2000017771286522", "x", 79.79, "y", "Fora da Regra")}
      {_linha("2000017944118140", "x", 57.70, "y", "Fora da Regra")}
      {_linha("2000018001112988", "x", 75.84, "y", "Fora da Regra")}
      {_linha("2000017887346980", "x", 70.00, "y", "Fora da Regra")}
      {_linha("2000018006603830", "x", 60.00, "y", "Fora da Regra")}
      {_linha("2000018051194114", "x", 60.00, "y", "Fora da Regra")}
      {_linha("2000018025209720", "x", 119.00, "y", "Realizo a baixa do Titulo")}
    </Table>
  </Worksheet>
  <Worksheet ss:Name="2-debito">
    <Table>
      <Row>{_cel("2-debito")}</Row>
      {_linha("External", "Tipo", "Valor", "Acao")}
      {_linha("2000017834924604", "x", -30.50, "Fora da Regra")}
      {_linha("2000017999272824", "x", -15.04, "Realizo a baixa do Titulo")}
    </Table>
  </Worksheet>
  <Worksheet ss:Name="3-Banco">
    <Table>
      <Row>{_cel("3-Banco")}</Row>
      {_linha("Banco", "Agencia", "Conta")}
      {_linha("237", "3708", "X103925")}
    </Table>
  </Worksheet>
</Workbook>"""

    with tempfile.TemporaryDirectory() as tmp:
        caminho = os.path.join(tmp, "marketplace_teste.xls")
        with open(caminho, "w", encoding="utf-8") as f:
            f.write(_XML)

        fora = ler_fora_da_regra(caminho)
        assert len(fora) == 7
        assert round(sum(v for _, v in fora), 2) == 372.83
        assert ("2000017834924604", -30.5) in fora
        assert ("2000018025209720", 119.0) not in [(p, v) for p, v in fora]  # baixou sozinho, fora

        df = gerar_rebate_fora_da_regra_cmm(caminho, "AB_ELETROD", "16/09/2026", "Everton")
        assert list(df.columns) == COLUNAS
        assert len(df) == 7
        assert (df["BANCO"] == BANCO).all() and (df["AGENCIA"] == AGENCIA).all() and (df["CONTA"] == CONTA).all()
        assert (df["LOJA"] == "ELETROD").all()
        assert (df["NATUREZA"] == NATUREZA_REBATE).all()
        assert (df["HISTORICO"] == "MKP_ME_FULL_RECEBIMENTO DE REBATE").all()
        assert df["VENDEDOR"].isna().all()

    print("rebate_fora_da_regra.py: ok")

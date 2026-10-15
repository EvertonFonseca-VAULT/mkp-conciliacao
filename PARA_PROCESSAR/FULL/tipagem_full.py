"""Tipagem dos ABs do Mercado FULL (`MERCADO FULL DD.MM.AAAA.xlsx`, aba
'Relatorio'). O export vem com layout fixo de 40 colunas (RELEASE_DATE até
BUSINESS_UNIT) e as colunas de valor (F a N) às vezes chegam como texto com
ponto decimal em vez de número — aqui garante que virem número de verdade,
com separador decimal de vírgula na exibição do Excel (locale pt-BR).
"""

import glob
import os

import openpyxl

LAYOUT_FULL = [
    "RELEASE_DATE", "SOURCE_ID", "EXTERNAL_REFERENCE", "RECORD_TYPE", "DESCRIPTION",
    "NET_CREDIT_AMOUNT", "NET_DEBIT_AMOUNT", "GROSS_AMOUNT", "SELLER_AMOUNT",
    "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT", "SHIPPING_FEE_AMOUNT", "TAXES_AMOUNT",
    "COUPON_AMOUNT", "INSTALLMENTS", "PAYMENT_METHOD", "TAX_DETAIL",
    "TAX_AMOUNT_TELCO", "APPROVAL_DATE", "POS_ID", "POS_NAME", "EXTERNAL_POS_ID",
    "STORE_ID", "STORE_NAME", "EXTERNAL_STORE_ID", "CURRENCY", "TAXES_DISAGGREGATED",
    "SHIPPING_ID", "SHIPMENT_MODE", "ORDER_ID", "PACK_ID", "METADATA",
    "EFFECTIVE_COUPON_AMOUNT", "POI_ID", "CARD_INITIAL_NUMBER", "TRANSACTION_LABELS",
    "TRANSACTION_INTENT_ID", "PAYOUT_BANK_ACCOUNT_NUMBER", "DIGITAL_WALLET",
    "BUSINESS_UNIT",
]

COLUNAS_VALOR = ["F", "G", "H", "I", "J", "K", "L", "M", "N", "AG"]  # NET_CREDIT_AMOUNT..COUPON_AMOUNT e EFFECTIVE_COUPON_AMOUNT


FORMATO_VALOR = "#,##0.00"  # exibido com vírgula decimal no Excel pt-BR


def validar_layout(ws):
    header = [ws.cell(row=1, column=c).value for c in range(1, len(LAYOUT_FULL) + 1)]
    if header != LAYOUT_FULL:
        raise ValueError(f"Layout do Mercado FULL mudou. Esperado: {LAYOUT_FULL}\nEncontrado: {header}")


def tipar_valores(caminho):
    """Garante que as colunas F-N sejam número (float) com formatação de
    vírgula decimal. Salva o arquivo em cima do original."""
    wb = openpyxl.load_workbook(caminho)
    ws = wb["Relatorio"]
    validar_layout(ws)

    for row in range(2, ws.max_row + 1):
        for col in COLUNAS_VALOR:
            cell = ws[f"{col}{row}"]
            if cell.value is None:
                continue
            cell.value = float(str(cell.value).replace(",", "."))
            cell.number_format = FORMATO_VALOR

    wb.save(caminho)


def main():
    pasta = os.path.dirname(os.path.abspath(__file__))
    for caminho in sorted(glob.glob(os.path.join(pasta, "MERCADO FULL *.xlsx"))):
        tipar_valores(caminho)
        print(f"Tipado: {os.path.basename(caminho)}")


if __name__ == "__main__":
    main()

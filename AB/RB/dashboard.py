"""Gera o dashboard HTML (autocontido, sem CDN/JS externo -- roda offline
numa pasta de rede) da conciliacao do dia: faixa de status (o dia fecha com
o calendario?), fluxo Original x Transitoria, historico TOTVS, a fila
VERIFICAR (onde a analista trabalha) e a gaveta com as linhas prontas pra
colar no TOTVS (CMM e MOV).

Nova cara (21/09/2026): recriada a partir do handoff do design
(design_handoff_conciliacao_full/README.md) -- tokens claro/escuro
(prefers-color-scheme + botao auto/claro/escuro), IBM Plex quando instalada
(sem baixar fonte: precisa abrir offline), valores com sinal colorido
(negativo vermelho, positivo verde). A logica de negocio NAO mudou: mesmas
naturezas/contas/historicos do CMM e do MOV (validados com o CMM real),
mesma persistencia em localStorage, mesma conferencia RB/RBX.

1 arquivo por dia (dashboard_<data>.html -- pedido do Everton, so quer 1
link): cada `gerar_dashboard()` SOBRESCREVE esse mesmo arquivo com o canal
que acabou de rodar (FULL, depois ES, etc. -- so o ultimo fica na tela; o
historico de cada canal continua nos xlsx salvos do lado). A pagina se
recarrega sozinha quando o Everton volta pra aba (visibilitychange), entao
ele so roda o processo de novo e alterna pra aba que ja esta aberta.

O template e texto puro com marcadores @@NOME@@ (nada de f-string: CSS e JS
cheios de chaves ja doeram aqui antes).
"""

import datetime
import html
import json
import os

from canais import CANAIS
from rebate_cmm import COLUNAS as COLUNAS_CMM
from rebate_cmm import MOTIVO_REBATE, NATUREZA_REBATE

TOLERANCIA_ALVO = 0.01

# Marcacao de NATUREZA na fila VERIFICAR (Frete/Comissao/Rebate) -- gera as
# linhas prontas pra colar no CMM (Movimento_Bancario) e no MOV, evitando o
# Everton digitar isso a mao toda vez. So essas 3 naturezas confirmadas ate
# agora (com dado real do CMM_MELI_FULL); Desc/Bonus e Repasse ficam de fora.
# NATUREZA_REBATE/MOTIVO_REBATE reaproveitados de rebate_cmm.py (mesmo
# valor ja confirmado la). Pedido do Everton, 17/09/2026.
NATUREZA_FRETE = 2209
MOTIVO_FRETE = "DESPESAS DE FRETE ENVIO"
NATUREZA_COMISSAO = 2804
MOTIVO_COMISSAO = "TARIFA DE COBRANÇA"
# Frete/Tarifa/Rebate sao SEMPRE lancados na Transitoria (X103925), mesmo
# quando o CMM real mostrado ao Everton exibe "103925" na coluna CONTA
# daquelas linhas -- nao e pra copiar isso literalmente. Correcao do
# Everton, 17/09/2026.
CONTA_TRANSITORIA = "X103925"
BANCO_PADRAO = "237"
AGENCIA_PADRAO = "3708"

# Marcacoes extras selecionaveis na fila VERIFICAR (Mediacao/Cashback/etc) --
# ao contrario de Frete/Comissao/Rebate (calculadas a partir de colunas
# especificas do ML/MP), essas marcam o DIF INTEIRO da linha: pro Everton
# escolher direto no dash o que aquela diferenca e, em vez de lancar tudo a
# mao no TOTVS depois. Mesma conta/banco/agencia do resto (Transitoria
# X103925/237/3708, confirmado com o Everton 18/09/2026) -- so troca
# NATUREZA/MOTIVO conforme a opcao escolhida. HISTORICO segue o mesmo padrao
# (prefixo_mov + "_" + motivo, calculado no JS -- ver PREFIXO_MOV).
# Comeca vazia DE PROPOSITO: nenhum codigo de NATUREZA chutado (pedido do
# Everton, 18/09/2026 -- ainda sem os codigos reais de Mediacao/Cashback
# confirmados no CMM). O seletor so aparece na tabela quando esse dict tem
# pelo menos 1 opcao. Preencher assim pra ativar (rotulo -> (natureza, motivo)):
#   NATUREZAS_EXTRAS = {
#       "Mediacao": (NNNN, "TEXTO MOTIVO DO CMM"),
#       "Cashback": (NNNN, "TEXTO MOTIVO DO CMM"),
#   }
NATUREZAS_EXTRAS = {}

# Mapa coluna-do-CMM -> expressao JS que preenche ela, na MESMA ordem de
# rebate_cmm.COLUNAS -- gerado a partir dessa lista (nao duplicado a mao)
# pra nunca dessincronizar se a ordem real do CMM mudar. Campos fixos por
# linha (RESPONSAVEL/DATA_LACTO/ORIGEM/PAGADOR/LOJA) sao constantes JS
# geradas por dashboard; os demais vem do proprio checkbox marcado
# (chk.dataset.*, ver _conteudo_natureza).
_CAMPOS_CMM_JS = {
    "RESPONSAVEL": "RESPONSAVEL", "DATA_LACTO": "DATA_LACTO", "ORIGEM": "ORIGEM",
    "PAGADOR": "PAGADOR", "LOJA": "LOJA",
    "MOTIVO DO MOVIMENTO": "chk.dataset.motivo", "NATUREZA": "chk.dataset.natureza",
    # apostrofo na frente, igual o botao do pedido -- sem isso o pedido (ate
    # 16 digitos) vira notacao cientifica ao colar no Excel.
    "PEDIDO MKP / CV": "\"'\" + chk.dataset.pedido", "VENDEDOR": "VENDEDOR",
    "BANCO": "chk.dataset.banco", "AGENCIA": "chk.dataset.agencia",
    "CONTA": "chk.dataset.conta", "HISTORICO": "chk.dataset.historico",
    "VALOR R$": "chk.dataset.valor", "OBSERVACAO": '""',
    "PEDIDO_ERP": '""', "PEDIDO_SITE": '""', "N. FISCAL": '""', "PROCESSO": '""',
}


# ---------------------------------------------------------------------------
# Formatacao
# ---------------------------------------------------------------------------

def _nulo(v):
    return v is None or (isinstance(v, float) and v != v)  # None ou NaN


def _numero_br(valor):
    """1.234,56 -- negativo com o sinal de menos tipografico (U+2212)."""
    texto = f"{abs(valor):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"{chr(0x2212) if valor < 0 else ''}{texto}"


def _moeda(valor):
    neg = valor < 0
    texto = f"{abs(valor):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"{chr(0x2212) if neg else ''}R$ {texto}"


def _formatar_numero(valor):
    """Numero pro CMM/MOV (data-valor): virgula decimal, sem milhar, hifen ASCII."""
    neg = valor < 0
    numero = f"{abs(valor):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"{'-' if neg else ''}{numero}"


def _valor_html(v, extra=""):
    """Valor monetario da tabela: mono/tabular, negativo (debito) vermelho,
    positivo (credito) verde, zero neutro, ausente "—". Cor nunca e o unico
    sinal: o sinal de menos vem escrito."""
    if _nulo(v):
        return '<span class="vazio">—</span>'
    if not isinstance(v, (int, float)):
        return html.escape(str(v))
    classe = "neg" if v < 0 else ("pos" if v > 0 else "zero")
    return f'<span class="v {classe}{extra}">{_numero_br(v)}</span>'


def _valor_ou_vazio(v):
    """Celula em branco (sem "—") quando nao tem valor -- usado onde o "—" so
    poluiria (RBX/RB: a celula vazia ja diz "nao achei nada la")."""
    return "" if _nulo(v) else _valor_html(v)


# ---------------------------------------------------------------------------
# Pedaços de HTML
# ---------------------------------------------------------------------------

def _painel_historico(raiz):
    """RAIZ + codigos _REPASSE/_COMISSAO/_FRETE do historico TOTVS pra
    colar direto ao resolver um VERIFICAR. Sufixos seguem sempre o padrao
    RAIZ + "_REPASSE" etc (confirmado nas 93 linhas da tabela HISTORICOS
    do CONTROLE_V2.xlsx -- ver extrair_historicos.py/historicos_totvs.json)
    -- deriva direto, nao depende dessa raiz estar na tabela (ex. ELETRO
    usa uma raiz propria desse pipeline, fora do CONTROLE_V2)."""
    linhas = [("Raiz", raiz), ("Repasse", f"{raiz}_REPASSE"),
              ("Comissão", f"{raiz}_COMISSAO"), ("Frete", f"{raiz}_FRETE")]
    return "".join(
        f'<button type="button" class="hist" data-copiar="{html.escape(valor)}" title="Copiar">'
        f'<span class="hist-rot">{rotulo}</span><span class="hist-val">{html.escape(valor)}</span></button>'
        for rotulo, valor in linhas
    )


def _td(conteudo, classe="num", detalhe=None):
    """<td> com tooltip nativo (title) listando as sublinhas que compoem esse
    total, com o mesmo rotulo que o Mercado usou -- pedido do Everton,
    16/09/2026 ("regua de calculo" ao passar o mouse)."""
    if isinstance(detalhe, str) and detalhe:  # NaN (float) nao e tooltip
        return f'<td class="{classe} tem-detalhe" title="{html.escape(detalhe)}">{conteudo}</td>'
    return f'<td class="{classe}">{conteudo}</td>'


def _conteudo_natureza(valor, pedido, motivo, natureza, conta, historico):
    """Checkbox + valor de 1 marcacao de NATUREZA, sem o <td> em volta --
    normalmente 1 por celula, mas a coluna Frete pode precisar de 2 na MESMA
    celula quando o pedido tem custo de envio e tarifa de cobranca separados
    (ver _linhas_tabela). Checkbox so quando o valor existe e e != 0; cada
    valor e um <label> (clicar no numero marca). Pedido do Everton,
    17/09/2026."""
    valor_html = _valor_ou_vazio(valor)
    if not valor_html or not valor:  # "" (None/NaN) e 0 nao geram checkbox
        return valor_html
    return (
        f'<label class="nat" title="Marcar: {html.escape(motivo)} (vai pro CMM/MOV)">'
        '<input type="checkbox" class="chk-natureza" '
        f'data-pedido="{pedido}" data-motivo="{html.escape(motivo)}" '
        f'data-natureza="{natureza}" data-conta="{conta}" '
        f'data-historico="{html.escape(historico)}" '
        f'data-banco="{BANCO_PADRAO}" data-agencia="{AGENCIA_PADRAO}" '
        f'data-valor="{_formatar_numero(valor)}">{valor_html}</label>'
    )


def _cel_marcacao_extra(dif, pedido, conta, banco, agencia):
    """<select> pra marcar o DIF INTEIRO da linha VERIFICAR com uma
    natureza extra (Mediacao/Cashback/etc, ver NATUREZAS_EXTRAS) -- no
    maximo 1 escolha por pedido. Celula em branco quando DIF e None/0. So
    chamada quando NATUREZAS_EXTRAS ja tem pelo menos 1 opcao."""
    if dif is None or not dif:
        return "<td></td>"
    opcoes = "".join(
        f'<option value="{html.escape(rotulo)}" data-natureza="{natureza}" '
        f'data-motivo="{html.escape(motivo)}">{html.escape(rotulo)}</option>'
        for rotulo, (natureza, motivo) in NATUREZAS_EXTRAS.items()
    )
    select = (
        '<select class="sel-natureza-extra" '
        f'data-pedido="{pedido}" data-conta="{conta}" data-banco="{banco}" '
        f'data-agencia="{agencia}" data-valor="{_formatar_numero(dif)}">'
        f'<option value="">--</option>{opcoes}</select>'
    )
    return f"<td>{select}</td>"


def _repasse_calculado(dif, comissao, custo_envio):
    """Repasse = |DIF| + |Comissao| + |Custo de envio| -- soma dos
    absolutos, nao o absoluto da soma: o Frete sempre vem negativo, entao
    somar os valores com sinal nunca fecharia com a nota. Usa so o CUSTO
    DE ENVIO (nao o grupo Envios inteiro) -- a Tarifa de cobranca, quando
    existe, fica de fora daqui (continua contando na coluna Frete, que
    mostra o grupo Envios inteiro). Pedido do Everton, 16/09/2026. None se
    faltar algum dos 3 (pedido ainda sem dado do ML/MP)."""
    if dif is None or comissao is None or custo_envio is None:
        return None
    return round(abs(dif) + abs(comissao) + abs(custo_envio), 2)


def _cel_valor_conta(pedido, valores):
    """Valor liquido lancado nesse pedido numa conta (RB ou RBX) -- celula em
    branco quando nao achou nada la (ou quando nao ha rb.xlsx/rbx.xlsx pra
    conferir esse dia). Pedido do Everton, 15-17/09/2026."""
    return f'<td class="num">{_valor_ou_vazio((valores or {}).get(pedido))}</td>'


def _bolinha(v_rb, v_rbx, diferenca_extrato):
    """Bolinha da Dif. extrato: ambar quando o pedido nao foi encontrado nem
    no RB nem no RBX (nada pra comparar ainda); verde quando achou e a
    diferenca fecha (dentro de TOLERANCIA_ALVO); vermelha quando achou mas
    sobra diferenca. Pedido do Everton, 16/09/2026."""
    if v_rb is None and v_rbx is None:
        return "warn", "Não encontrado no RB nem no RBX — nada para comparar ainda"
    if abs(diferenca_extrato) < TOLERANCIA_ALVO:
        return "ok", "Fecha: o extrato tem o que o AB diz que deveria ter"
    return "bad", "Sobra diferença entre o extrato (RB/RBX) e o AB"


def _linhas_tabela(fila_verificar, conferencia_rbx=None, historicos=None, valores_pedidos=None):
    # valores_pedidos: fallback pedido a pedido sem chave_rbx (ver
    # conferir_rbx.conferir_pedidos) -- so usado quando conferencia_rbx e
    # None (canal sem chave_rbx confirmada, ex. ES). Pedido do Everton,
    # 18/09/2026.
    fonte_valores = conferencia_rbx or valores_pedidos
    valores_rb = fonte_valores["valores_rb"] if fonte_valores else None
    valores_rbx = fonte_valores["valores_rbx"] if fonte_valores else None
    historicos = historicos or {}
    saida = []
    for _, row in fila_verificar.iterrows():
        pedido = html.escape(str(row["PEDIDO_MKP"]))
        dif = row["DIF"]
        comissao = row.get('ML_COMISSAO_LIQUIDA', row.get('ML_COMISSAO'))
        frete = row.get('ML_ENVIO')
        custo_envio = row.get('ML_CUSTO_ENVIO')
        # Everton confirmou (17/09/2026): a coluna Comissao (ML_COMISSAO_LIQUIDA)
        # fica de fora da marcacao de NATUREZA -- a Tarifa de Cobranca (2804)
        # nao mora ali, mora dentro do proprio grupo Envios (coluna Frete),
        # como a diferenca entre o grupo inteiro (ML_ENVIO) e a sublinha
        # "custo de envio" (ML_CUSTO_ENVIO) -- confirma o que
        # _repasse_calculado ja dizia. Quando o pedido tem as 2 sublinhas
        # separadas, a celula Frete ganha 2 valores empilhados, cada um com
        # seu checkbox (custo de envio + tarifa); quando so tem 1 (ou
        # nenhuma), continua com 1 so.
        cel_comissao = _td(_valor_html(comissao), detalhe=row.get('ML_COMISSAO_DETALHE'))
        cel_desconto = _td(_valor_html(row.get('ML_DESCONTO_BONUS')), detalhe=row.get('ML_DESCONTO_BONUS_DETALHE'))
        tarifa_cobranca = None
        if not _nulo(frete) and not _nulo(custo_envio) and round(frete - custo_envio, 2) != 0:
            tarifa_cobranca = round(frete - custo_envio, 2)
        if tarifa_cobranca is not None:
            conteudo_envio = (
                '<div class="pilha">'
                + _conteudo_natureza(custo_envio, pedido, MOTIVO_FRETE, NATUREZA_FRETE,
                                     CONTA_TRANSITORIA, historicos.get("frete", ""))
                + _conteudo_natureza(tarifa_cobranca, pedido, MOTIVO_COMISSAO, NATUREZA_COMISSAO,
                                     CONTA_TRANSITORIA, historicos.get("comissao", ""))
                + "</div>"
            )
        else:
            conteudo_envio = _conteudo_natureza(frete, pedido, MOTIVO_FRETE, NATUREZA_FRETE,
                                                CONTA_TRANSITORIA, historicos.get("frete", ""))
        cel_envio = _td(conteudo_envio or _valor_html(None), detalhe=row.get('ML_ENVIO_DETALHE'))
        conteudo_rebate = _conteudo_natureza(row.get('ML_REBATE'), pedido, MOTIVO_REBATE, NATUREZA_REBATE,
                                             CONTA_TRANSITORIA, historicos.get("rebate", ""))
        cel_rebate = _td(conteudo_rebate or _valor_html(None), detalhe=row.get('ML_REBATE_DETALHE'))

        repasse = _repasse_calculado(dif, comissao, custo_envio)
        detalhe_repasse = "Repasse líquido do pedido (histórico MKP_MELI_..._REPASSE)"
        if repasse is not None:
            detalhe_repasse = "\n".join([
                f"DIF (banco, absoluto): {abs(dif):.2f}".replace(".", ","),
                f"Comissao (absoluto): {abs(comissao):.2f}".replace(".", ","),
                f"Custo de envio (absoluto, sem tarifa de cobranca): {abs(custo_envio):.2f}".replace(".", ","),
                f"= Repasse: {repasse:.2f}".replace(".", ","),
            ])
        cel_repasse = _td(_valor_html(repasse), detalhe=detalhe_repasse)
        pedido_mkp = str(row["PEDIDO_MKP"]).strip()
        cel_rbx = _cel_valor_conta(pedido_mkp, valores_rbx)
        cel_rb = _cel_valor_conta(pedido_mkp, valores_rb)

        # Diferenca extrato = DIF - valor RB - valor RBX + TOTAL_RB -- None
        # quando nao achou em nenhuma das duas contas, pra nao sugerir uma
        # diferenca que na verdade e so a DIF original (a bolinha fica
        # ambar nesse caso). Nao muda o STATUS (continua VERIFICAR), so
        # informa o analista. Pedido do Everton, 16/09/2026.
        #
        # + TOTAL_RB confirmado pelo Everton, 18/09/2026: TOTAL_RB (achado
        # no cruzamento inicial, contra CONCILIADOR_RB.xlsx -- ver
        # cruzamento.py) e um SUBCONJUNTO do que RB/RBX (AUXILIAR/rb.xlsx e
        # rbx.xlsx) mostram pra esse pedido -- a mesma transacao que gerou
        # o DIF continua aparecendo em RB/RBX, entao subtrair RB/RBX
        # sozinho descontava essa parte 2x. Somar TOTAL_RB de volta cancela
        # so a parte duplicada, sobrando so o que e realmente NOVO em
        # RB/RBX (lancamento feito depois, cancelamento+reemissao, etc.).
        # Ex. real (10/09/2026): DIF=371.65, TOTAL_RB=1911.43, RBX=2283.08
        # (cancelou o lancamento antigo de 1911.43 e reemitiu um agregado de
        # 2283.08) -> 371.65 - 2283.08 + 1911.43 = 0 (fechou, verde).
        # TOTAL_RB=0 (pedido nao achado no cruzamento inicial) reduz pra
        # formula de sempre (DIF - RB - RBX), sem precisar de caso especial.
        v_rb = (valores_rb or {}).get(pedido_mkp)
        v_rbx = (valores_rbx or {}).get(pedido_mkp)
        diferenca_extrato = (
            round(dif - (v_rb or 0) - (v_rbx or 0) + row.get("TOTAL_RB", 0), 2)
            if v_rb is not None or v_rbx is not None else None
        )
        if diferenca_extrato is None:
            cor, dica = _bolinha(v_rb, v_rbx, 0)
            cel_diferenca = f'<td class="num" title="{dica}"><span class="v nao-achado">n/e</span><span class="dot {cor}"></span></td>'
        else:
            cor, dica = _bolinha(v_rb, v_rbx, diferenca_extrato)
            cel_diferenca = f'<td class="num" title="{dica}">{_valor_html(diferenca_extrato)}<span class="dot {cor}"></span></td>'
        # AB total (TOTAL_PED = REPASSE+COMISSAO+FRETE, ver cruzamento.py)
        # -- do lado de RBX/RB/Diferenca extrato, pro Everton comparar
        # direto o que o banco tem (RBX/RB) com o que o AB/ML diz que
        # deveria ter, sem abrir outra planilha. Pedido do Everton,
        # 18/09/2026.
        cel_ab_total = f'<td class="num">{_valor_html(row.get("TOTAL_PED", 0))}</td>'
        # DESCRICAO (coluna E do AB, ver cruzamento.py) -- pro Everton ver
        # "mediation"/"refund"/etc na fila sem abrir o AB pra procurar o
        # pedido. Vazio quando o pedido nao tem nenhuma linha no AB do dia
        # (so no RB/RBX). Pedido do Everton, 18/09/2026.
        cel_descricao = f'<td class="desc">{html.escape(str(row.get("DESCRICAO") or ""))}</td>'
        cel_marcacao = (
            _cel_marcacao_extra(dif, pedido, CONTA_TRANSITORIA, BANCO_PADRAO, AGENCIA_PADRAO)
            if NATUREZAS_EXTRAS else ""
        )
        dica_dif = "DIF = repasse (AB) − total já lançado no extrato (RB)"

        saida.append(f"""
            <tr data-pedido="{pedido}" data-dif="{float(dif):.2f}">
              <td class="c-chk"><input type="checkbox" class="chk-ajustado" data-pedido="{pedido}" aria-label="Ajustado"></td>
              {cel_rbx}
              {cel_rb}
              {cel_diferenca}
              {cel_ab_total}
              <td class="ped"><button type="button" class="pedido" data-copiar="'{pedido}" title="Copiar pedido">{pedido}</button></td>
              {cel_descricao}
              <td class="num dif" title="{dica_dif}">{_valor_html(dif)}</td>
              {cel_marcacao}
              {cel_comissao}
              {cel_desconto}
              {cel_envio}
              {cel_repasse}
              {cel_rebate}
            </tr>""")
    return "".join(saida)


def _faixa_status(alvo, rb_total, rebate_total, ajuste_pendente, conferencia_rbx=None):
    """Faixa de status do dia (destaque maximo): veredito grande + a conta
    RB atual + Rebate (cupom) + Ajuste pendente = Calculado x Alvo, mais a
    conferencia do Rebate (formula x RBX) quando ha rb.xlsx/rbx.xlsx."""
    calculado = rb_total + rebate_total + ajuste_pendente
    if alvo is None:
        estado, icone, rotulo = "neutro", "?", "Status do dia"
        titulo = "Sem alvo no calendário"
        texto = "Alvo não cadastrado no calendário para este canal/dia — não dá para conferir se o dia fecha."
        alvo_txt, dif_txt = "—", "sem alvo"
    else:
        diferenca = round(alvo - calculado, 2) + 0.0  # sem "−R$ 0,00" por ruido de ponto flutuante
        bate = abs(diferenca) < TOLERANCIA_ALVO
        estado = "ok" if bate else "bad"
        icone, rotulo = ("✓", "Status do dia") if bate else ("!", "Status do dia")
        titulo = "Dia fechado ✓" if bate else f"Diferença de {_moeda(abs(diferenca))}"
        texto = ("RB atual, rebate e ajuste pendente fecham com o calendário do dia." if bate else
                 "RB atual, rebate e ajuste pendente não fecham com o alvo do calendário. Confira a fila VERIFICAR e o rebate.")
        alvo_txt, dif_txt = _moeda(alvo), _moeda(diferenca)

    rebate_html = ""
    if conferencia_rbx:
        confere = conferencia_rbx["rebate_bate"]
        rebate_html = (
            '<div class="faixa-rebate">'
            '<span class="rot">Conferência do rebate</span>'
            f'<span class="mono forte" title="Fórmula: cupons aplicados no período + lançamentos CMM">fórmula {_moeda(rebate_total)}</span>'
            '<span class="mono muda">×</span>'
            f'<span class="mono forte">RBX {_moeda(conferencia_rbx["liquido_rebate_rbx"])}</span>'
            f'<span class="apoio">cupom {conferencia_rbx["qtd_cupom_rebate"]} lançamentos · '
            f'CMM {conferencia_rbx["qtd_cmm_rebate"]} lançamentos</span>'
            f'<span class="pill {"ok" if confere else "bad"}"><span class="dot"></span>{"confere" if confere else "diverge"}</span>'
            '</div>'
        )
    return f"""
    <section class="card faixa {estado}">
      <div class="faixa-veredito">
        <div class="selo-linha"><span class="selo">{icone}</span><span class="rot-cor">{rotulo}</span></div>
        <div class="veredito">{html.escape(titulo)}</div>
        <div class="veredito-texto">{html.escape(texto)}</div>
      </div>
      <div class="faixa-conta">
        <div class="operandos">
          <div class="op"><span class="rot">RB atual</span><span class="mono val">{_moeda(rb_total)}</span></div>
          <span class="mono sinal">+</span>
          <div class="op"><span class="rot">Rebate (cupom)</span><span class="mono val">{_moeda(rebate_total)}</span></div>
          <span class="mono sinal">+</span>
          <div class="op"><span class="rot">Ajuste pendente</span><span class="mono val" title="Soma dos DIF da fila VERIFICAR">{_moeda(ajuste_pendente)}</span></div>
          <span class="mono sinal">=</span>
          <div class="op"><span class="rot">Calculado</span><span class="mono val grande">{_moeda(calculado)}</span></div>
          <div class="divisor"></div>
          <div class="op"><span class="rot">Alvo (calendário)</span><span class="mono val grande">{alvo_txt}</span></div>
          <div class="dif-badge"><span class="dot"></span><span>dif. {dif_txt}</span></div>
        </div>
        {rebate_html}
      </div>
    </section>"""


def _fluxo_original_transitoria(conf, alvo):
    """Conta ORIGINAL (103925) x conta TRANSITORIA (X103925) -- ver
    conferir_rbx.py -- como fluxo de 3 passos. So chamado quando conf nao e
    None (arquivos rb.xlsx/rbx.xlsx existem em AUXILIAR/).

    Correcao de regra, 17/09/2026: a Transitoria NAO zera durante o dia --
    ela acumula o movimento. Quando o liquido dela (chave do canal + RA)
    bate com o MESMO alvo usado na linha Original, e sinal de que ja pode
    fazer a transferencia manual pra Original (que ai sim zera a
    Transitoria). Tem 2 momentos CERTOS -- bate com o alvo (pronta pra
    transferir) OU voltou a zero (transferencia ja feita); so no meio do dia,
    acumulando sem bater com nenhum dos dois, e que e "normal, nao e erro"
    (ambar, nao vermelho)."""
    # 1) Original
    if conf["qtd_original"] == 0:
        c1, n1, v1 = "neutro", "sem lançamento de fechamento ainda", "—"
    elif conf["original_bate"]:
        c1, n1, v1 = "ok", f'bate ✓ · {conf["qtd_original"]} lançamentos', _moeda(conf["liquido_original"])
    else:
        c1, n1, v1 = "bad", f'não bate · {conf["qtd_original"]} lançamentos', _moeda(conf["liquido_original"])

    # 2) Cancelado hoje: pedido lancado (e ja transferido) num dia ANTERIOR e
    # cancelado hoje -- so o debito residual fica na Original, sem entrada
    # correspondente na Transitoria de hoje, o que "infla" a Transitoria em
    # relacao ao alvo. Pedido do Everton, 18/09/2026.
    if conf["qtd_cancelado_original"]:
        lista = "".join(
            f'<div class="canc"><span>{html.escape(str(pedido))}</span><span class="v neg">{_numero_br(valor)}</span></div>'
            for pedido, valor in conf["cancelados_original"]
        )
        passo2 = (
            '<div class="passo bad"><div class="passo-topo"><span class="num-passo">2</span>'
            '<span class="rot-cor">Cancelado hoje</span></div>'
            f'<div class="mono passo-valor bad-txt" title="Cancelado na Original hoje">{_moeda(conf["liquido_cancelado_original"])}</div>'
            f'<div class="canc-lista">{lista}</div></div>'
        )
    else:
        passo2 = (
            '<div class="passo neutro"><div class="passo-topo"><span class="num-passo">2</span>'
            '<span class="rot-cor">Cancelado hoje</span></div>'
            '<div class="mono passo-valor">—</div><div class="passo-nota">nenhum cancelamento na Original hoje</div></div>'
        )

    # 3) Transitoria
    if conf["qtd_transitoria_simulado"] == 0:
        v3, pill3 = "—", '<span class="pill neutro"><span class="dot"></span>sem lançamentos encontrados</span>'
    else:
        v3 = _moeda(conf["liquido_transitoria_simulado"])
        if conf["zerada"]:
            pill3 = '<span class="pill ok"><span class="dot"></span>zerada — transferência já feita</span>'
        elif conf["pronta_para_transferencia"]:
            pill3 = '<span class="pill accent"><span class="dot"></span>pronta para transferência</span>'
        elif alvo is not None:
            faltam = _moeda(alvo - conf["valor_transferencia"])
            pill3 = f'<span class="pill warn"><span class="dot"></span>acumulando (faltam {faltam})</span>'
        else:
            pill3 = '<span class="pill neutro"><span class="dot"></span>sem alvo do calendário pra comparar</span>'

    return f"""
    <section class="card fluxo">
      <div class="card-topo"><h2>Original × Transitória</h2><span class="apoio">ciclo do dia</span></div>
      <div class="passos">
        <div class="passo {c1}"><div class="passo-topo"><span class="num-passo">1</span><span class="rot-cor">Original (calendário)</span></div>
          <div class="mono passo-valor">{v1}</div><div class="passo-nota">{n1}</div></div>
        {passo2}
        <div class="passo transitoria"><div class="passo-topo"><span class="num-passo">3</span><span class="rot-cor">Transitória</span></div>
          <div class="mono passo-valor">{v3}</div>{pill3}</div>
      </div>
      <div class="fluxo-rodape">
        <div class="op"><span class="rot">Alvo da transitória (calendário)</span><span class="mono val">{_moeda(alvo) if alvo is not None else "—"}</span></div>
        <div class="op"><span class="rot">Valor de transferência (ajustado)</span><span class="mono val">{_moeda(conf["valor_transferencia"])}</span></div>
      </div>
    </section>"""


_STATUS_LOTE = {"rodado": ("processado", "ok"), "pendente": ("pendente", "warn"), "erro": ("erro", "bad")}


def _tabela_lotes(lotes_status):
    """Guia 2: status dos lotes do pipeline AB bruto -> Para Processar (ver
    pipeline.py/conciliacao.db) -- vazio quando o banco ainda nao existe ou
    nao tem lote desse canal (feature opcional, mesmo padrao de
    conferencia_rbx=None). Pedido do Everton, 18/09/2026."""
    if not lotes_status:
        return '<p class="apoio vazio-msg">Nenhum lote registrado (pipeline.py sync).</p>'
    linhas = ""
    for lote, status, pedidos, acao_manual, rodado_em in lotes_status:
        rotulo, cor = _STATUS_LOTE.get(status, (status, "neutro"))
        linhas += (
            f'<tr><td class="esq mono">{html.escape(lote)}</td>'
            f'<td class="esq"><span class="pill {cor}"><span class="dot"></span>{html.escape(rotulo)}</span></td>'
            f'<td class="num">{pedidos if pedidos is not None else "—"}</td>'
            f'<td class="esq">{"nenhuma" if acao_manual == 0 else (acao_manual if acao_manual is not None else "—")}</td>'
            f'<td class="esq mono">{html.escape(rodado_em or "—")}</td></tr>'
        )
    return (
        '<div class="rolagem"><table class="tab"><thead><tr><th class="esq">Lote</th><th class="esq">Status</th>'
        '<th>Pedidos</th><th class="esq">Ação manual</th><th class="esq">Rodado em</th></tr></thead>'
        f"<tbody>{linhas}</tbody></table></div>"
    )


# Colunas de valor de verdade da aba "Para Processar" (ver
# processar_full.COLUNAS_PARA_PROCESSAR) -- so essas usam _valor_html
# (formatacao de moeda). As demais (SOURCE_ID/EXTERNAL_REFERENCE/etc) sao
# ID/texto: o pandas as carrega como float (NaN em alguma linha forca a
# coluna inteira pra float64) e formatava esses IDs como dinheiro
# ("2.000.017,00" em vez de "2000017591558558") -- bug real reportado pelo
# Everton, 18/09/2026.
_COLUNAS_VALOR_PARA_PROCESSAR = {
    "NET_CREDIT_AMOUNT", "NET_DEBIT_AMOUNT", "GROSS_AMOUNT", "MP_FEE_AMOUNT",
    "FINANCING_FEE_AMOUNT", "SHIPPING_FEE_AMOUNT", "TAXES_AMOUNT", "COUPON_AMOUNT",
}


def _valor_texto(v):
    """Celula de ID/texto (SOURCE_ID, EXTERNAL_REFERENCE, DATE, etc.) --
    numero inteiro sem o '.0' que o pandas cola em coluna float com alguma
    celula vazia (mesmo problema que _limpar_id resolve no cruzamento),
    sem formatacao de moeda (nao e valor monetario). Vazio quando
    None/NaN."""
    if _nulo(v):
        return '<span class="vazio">—</span>'
    if isinstance(v, float) and v.is_integer():
        return html.escape(str(int(v)))
    return html.escape(str(v))


def _tabela_para_processar(para_processar):
    """Guia 2: linhas cruas da aba "Para Processar" do AB do dia (ver
    processar_full.gerar_para_processar) -- colunas de valor (_AMOUNT)
    formatadas como moeda, o resto como ID/texto puro. Tolera N colunas.
    Pedido do Everton, 18/09/2026."""
    if para_processar is None or not len(para_processar):
        return '<p class="apoio vazio-msg">Sem dados do AB (Para Processar) pra esse lote.</p>'
    colunas = list(para_processar.columns)
    cabecalho = "".join(
        f'<th class="{"" if c in _COLUNAS_VALOR_PARA_PROCESSAR else "esq"}">{html.escape(str(c))}</th>' for c in colunas)
    linhas = "".join(
        "<tr>" + "".join(
            f'<td class="{"num" if c in _COLUNAS_VALOR_PARA_PROCESSAR else "esq"}">'
            f'{_valor_html(row[c]) if c in _COLUNAS_VALOR_PARA_PROCESSAR else _valor_texto(row[c])}</td>'
            for c in colunas
        ) + "</tr>"
        for _, row in para_processar.iterrows()
    )
    return f'<div class="rolagem ab-cru"><table class="tab mono"><thead><tr>{cabecalho}</tr></thead><tbody>{linhas}</tbody></table></div>'


def _lista_valores(valores):
    """Guia 2: 1 linha por pedido com o valor liquido lancado numa conta
    (RB ou RBX) -- mesmos dados de _cel_valor_conta (valores_rb/valores_rbx
    de conferir_rbx.py), so que numa lista dedicada em vez de coluna da
    fila VERIFICAR. Ordenado do mais negativo pro mais positivo. Pedido do
    Everton, 18/09/2026."""
    if not valores:
        return '<p class="apoio vazio-msg">Sem rb.xlsx/rbx.xlsx pra conferir ainda.</p>'
    linhas = "".join(
        f'<div class="li"><button type="button" class="pedido" data-copiar="\'{html.escape(str(pedido))}" title="Copiar pedido">'
        f'{html.escape(str(pedido))}</button>{_valor_html(valor)}</div>'
        for pedido, valor in sorted(valores.items(), key=lambda kv: kv[1])
    )
    return f'<div class="lista">{linhas}</div>'


def _cartao_conta(titulo, valores):
    total = sum(valores.values()) if valores else 0.0
    return f"""
        <section class="card">
          <div class="card-topo"><h2>{titulo}</h2>
            <span class="apoio mono">{len(valores)} lançamentos · {_moeda(total)}</span></div>
          {_lista_valores(valores)}
        </section>"""


# ---------------------------------------------------------------------------
# Template (texto puro; marcadores @@NOME@@)
# ---------------------------------------------------------------------------

_CSS = r"""
:root{
  --bg:#f6f6f4;--surface:#ffffff;--surface2:#fbfbfa;--ink:#16181a;--ink2:#575d63;--ink3:#8b9197;
  --line:#e5e5e1;--line2:#efefec;--ok:#0d7a4d;--okbg:#e9f5ee;--okline:#bfe0cd;
  --bad:#b0261c;--badbg:#fdecea;--badline:#f0c8c3;--warn:#8a5a00;--warnbg:#fcf2df;--warnline:#ecd6a8;
  --accent:#2a56c6;--accentbg:#eaf0fd;--accentline:#c7d6f5;
  --shadow:0 1px 2px rgba(20,22,25,.05),0 6px 16px -8px rgba(20,22,25,.12);
}
@media (prefers-color-scheme:dark){
  :root{
    --bg:#0f1114;--surface:#161a1e;--surface2:#1b2024;--ink:#e8ebee;--ink2:#a2aab2;--ink3:#6d767e;
    --line:#262c32;--line2:#1f252a;--ok:#4fc98d;--okbg:#122a20;--okline:#1f4b38;
    --bad:#ff8078;--badbg:#2c1614;--badline:#5a2b26;--warn:#e0aa53;--warnbg:#2a2010;--warnline:#544021;
    --accent:#84a7f7;--accentbg:#17203a;--accentline:#2c3b66;
    --shadow:0 1px 2px rgba(0,0,0,.4),0 8px 22px -10px rgba(0,0,0,.6);
  }
}
html[data-theme="light"]{
  --bg:#f6f6f4;--surface:#ffffff;--surface2:#fbfbfa;--ink:#16181a;--ink2:#575d63;--ink3:#8b9197;
  --line:#e5e5e1;--line2:#efefec;--ok:#0d7a4d;--okbg:#e9f5ee;--okline:#bfe0cd;
  --bad:#b0261c;--badbg:#fdecea;--badline:#f0c8c3;--warn:#8a5a00;--warnbg:#fcf2df;--warnline:#ecd6a8;
  --accent:#2a56c6;--accentbg:#eaf0fd;--accentline:#c7d6f5;
  --shadow:0 1px 2px rgba(20,22,25,.05),0 6px 16px -8px rgba(20,22,25,.12);
}
html[data-theme="dark"]{
  --bg:#0f1114;--surface:#161a1e;--surface2:#1b2024;--ink:#e8ebee;--ink2:#a2aab2;--ink3:#6d767e;
  --line:#262c32;--line2:#1f252a;--ok:#4fc98d;--okbg:#122a20;--okline:#1f4b38;
  --bad:#ff8078;--badbg:#2c1614;--badline:#5a2b26;--warn:#e0aa53;--warnbg:#2a2010;--warnline:#544021;
  --accent:#84a7f7;--accentbg:#17203a;--accentline:#2c3b66;
  --shadow:0 1px 2px rgba(0,0,0,.4),0 8px 22px -10px rgba(0,0,0,.6);
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);-webkit-font-smoothing:antialiased;
  font-family:'IBM Plex Sans','Segoe UI',system-ui,-apple-system,sans-serif}
.mono,.v,.pedido,.hist-val,textarea{font-family:'IBM Plex Mono','Cascadia Mono',Consolas,ui-monospace,monospace}
.pagina{padding:20px 24px 24px}
.conteudo{max-width:1560px;margin:0 auto;display:flex;flex-direction:column;gap:16px;padding-bottom:0}
body.drawer-aberto .conteudo{padding-bottom:210px}
body.drawer-min .conteudo{padding-bottom:88px}
button{font-family:inherit}
input[type=checkbox]{accent-color:var(--accent);width:15px;height:15px;cursor:pointer;margin:0;vertical-align:middle}
::-webkit-scrollbar{height:10px;width:10px}
::-webkit-scrollbar-thumb{background:var(--line);border-radius:8px}
::-webkit-scrollbar-track{background:transparent}
.apoio{font-size:12px;color:var(--ink3)}
.rot{font:500 10.5px/1 'IBM Plex Sans',sans-serif;letter-spacing:.07em;text-transform:uppercase;color:var(--ink3)}
.rot-cor{font:600 11px/1 'IBM Plex Sans',sans-serif;letter-spacing:.07em;text-transform:uppercase;color:var(--c,var(--ink2))}

/* cabecalho */
.topo{display:flex;flex-wrap:wrap;gap:16px;align-items:flex-start;justify-content:space-between;margin-bottom:2px}
.titulo-linha{display:flex;align-items:baseline;gap:12px;flex-wrap:wrap}
h1{margin:0;font:600 24px/1.1 'IBM Plex Sans',sans-serif;letter-spacing:-.015em}
.chip{font:500 11px/1 'IBM Plex Mono',monospace;color:var(--ink2);background:var(--surface);border:1px solid var(--line);padding:5px 8px;border-radius:6px}
.sub{font:400 12.5px/1.4 'IBM Plex Sans',sans-serif;color:var(--ink3);margin-top:6px}
.acoes{display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.btn{font:500 12.5px/1 'IBM Plex Sans',sans-serif;color:var(--ink2);background:var(--surface);border:1px solid var(--line);
  border-radius:8px;padding:9px 12px;cursor:pointer;text-decoration:none;display:inline-block}
.btn:hover{color:var(--ink);border-color:var(--ink3)}
/* texto na cor da superficie (nao #fff fixo): branco sobre o azul claro do tema escuro nao passa de contraste AA */
.btn.primario{font-weight:600;color:var(--surface);background:var(--accent);border-color:transparent;padding:9px 14px}
.btn.primario:hover{filter:brightness(1.08);color:var(--surface)}
.btn.tema{font-family:'IBM Plex Mono',monospace;padding:9px 11px}
.avisos{display:flex;flex-direction:column;gap:6px}
.aviso{display:flex;align-items:center;gap:8px;font-size:12.5px;color:var(--ink2)}
.aviso.ok{color:var(--ok);font-weight:600}
.aviso .dot{background:var(--ok)}

/* abas */
.abas{display:flex;gap:4px;border-bottom:1px solid var(--line);margin-bottom:4px}
.aba-btn{font:600 13px/1 'IBM Plex Sans',sans-serif;background:none;border:none;border-bottom:2px solid transparent;color:var(--ink3);
  padding:10px 4px 11px;margin-right:16px;cursor:pointer}
.aba-btn.ativo{color:var(--ink);border-bottom-color:var(--ink)}
.guia{display:flex;flex-direction:column;gap:16px}
.guia[hidden]{display:none}

/* cartoes e estados semanticos */
.card{background:var(--surface);border:1px solid var(--line);border-radius:14px;box-shadow:var(--shadow);padding:18px 20px 20px;min-width:0}
.card-topo{display:flex;align-items:baseline;justify-content:space-between;gap:12px;margin-bottom:14px}
.card h2{margin:0;font:600 14px/1 'IBM Plex Sans',sans-serif}
.ok{--c:var(--ok);--cbg:var(--okbg);--cline:var(--okline)}
.bad{--c:var(--bad);--cbg:var(--badbg);--cline:var(--badline)}
.warn{--c:var(--warn);--cbg:var(--warnbg);--cline:var(--warnline)}
.accent{--c:var(--accent);--cbg:var(--accentbg);--cline:var(--accentline)}
.neutro{--c:var(--ink2);--cbg:var(--surface2);--cline:var(--line)}
.dot{width:8px;height:8px;border-radius:50%;display:inline-block;flex-shrink:0;background:var(--c,var(--ink3));margin-left:6px}
.dot.ok{background:var(--ok)}.dot.bad{background:var(--bad)}.dot.warn{background:var(--warn)}
.pill{display:inline-flex;align-items:center;gap:6px;padding:5px 9px;border-radius:999px;background:var(--cbg);border:1px solid var(--cline);
  font:600 11.5px/1 'IBM Plex Sans',sans-serif;color:var(--c)}
.pill .dot{margin-left:0;width:7px;height:7px;background:var(--c)}

/* faixa de status */
.faixa{padding:0;overflow:hidden;display:grid;grid-template-columns:minmax(300px,1fr) minmax(420px,1.6fr);border-color:var(--cline)}
.faixa-veredito{padding:22px 24px;background:var(--cbg);border-right:1px solid var(--line);display:flex;flex-direction:column;gap:10px;justify-content:center}
.selo-linha{display:flex;align-items:center;gap:9px}
.selo{width:22px;height:22px;border-radius:50%;background:var(--c);color:var(--surface);display:flex;align-items:center;justify-content:center;font:700 13px/1 'IBM Plex Sans',sans-serif}
.veredito{font:600 38px/1.05 'IBM Plex Sans',sans-serif;letter-spacing:-.025em;color:var(--c)}
.veredito-texto{font:400 13px/1.45 'IBM Plex Sans',sans-serif;color:var(--ink2);max-width:42ch}
.faixa-conta{padding:20px 24px;display:flex;flex-direction:column;gap:16px;background:var(--surface)}
.operandos{display:flex;flex-wrap:wrap;align-items:flex-end;gap:10px 14px}
.op{display:flex;flex-direction:column;gap:4px}
.op .val{font-size:17px;font-weight:500;line-height:1}
.op .val.grande{font-size:19px;font-weight:600}
.sinal{font:400 16px/1.6 'IBM Plex Mono',monospace;color:var(--ink3)}
.divisor{width:1px;height:38px;background:var(--line);margin:0 4px}
.dif-badge{display:flex;align-items:center;gap:7px;padding:7px 10px;border-radius:8px;background:var(--cbg);border:1px solid var(--cline);font:500 12px/1 'IBM Plex Sans',sans-serif;color:var(--c)}
.dif-badge .dot{margin:0;background:var(--c)}
.faixa-rebate{display:flex;flex-wrap:wrap;align-items:center;gap:10px;padding-top:14px;border-top:1px solid var(--line2)}
.faixa-rebate .rot{font-weight:600}
.forte{font-size:13px;font-weight:500;color:var(--ink)}
.muda{font-size:13px;color:var(--ink3)}

/* fluxo Original x Transitoria + historico */
.duas{display:grid;grid-template-columns:minmax(560px,2.2fr) minmax(260px,1fr);gap:16px}
.duas.uma{grid-template-columns:1fr}
.passos{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;align-items:stretch}
.passo{border:1px solid var(--cline);background:var(--cbg);border-radius:11px;padding:14px;display:flex;flex-direction:column;gap:8px}
.passo.transitoria{--c:var(--ink2);--cbg:var(--surface2);--cline:var(--line);align-items:flex-start}
.passo-topo{display:flex;align-items:center;gap:7px;margin-bottom:2px}
.num-passo{width:18px;height:18px;border-radius:50%;background:var(--c);color:var(--surface);display:flex;align-items:center;justify-content:center;font:600 10px/1 'IBM Plex Mono',monospace}
.passo-valor{font-size:21px;font-weight:600;line-height:1;color:var(--ink)}
.passo-valor.bad-txt{color:var(--bad)}
.passo-nota{font-size:12px;line-height:1.4;color:var(--ink2)}
.canc-lista{display:flex;flex-direction:column;gap:5px}
.canc{display:flex;justify-content:space-between;gap:8px;font:400 11.5px/1 'IBM Plex Mono',monospace;color:var(--ink2)}
.fluxo-rodape{display:flex;flex-wrap:wrap;gap:24px;margin-top:16px;padding-top:14px;border-top:1px solid var(--line2)}
.fluxo-rodape .val{font-size:14px;font-weight:500;line-height:1}
.historico h2{color:var(--ink2)}
.hist-lista{display:flex;flex-direction:column;gap:6px;margin-top:12px}
.hist{display:flex;align-items:center;justify-content:space-between;gap:10px;width:100%;background:none;border:1px solid var(--line2);
  border-radius:8px;padding:9px 10px;cursor:pointer;text-align:left}
.hist:hover{border-color:var(--ink3);background:var(--surface2)}
.hist-rot{font-size:11px;color:var(--ink3)}
.hist-val{font-size:11.5px;font-weight:500;color:var(--ink)}
.hist.copiado .hist-val{color:var(--ok)}

/* fila VERIFICAR */
.fila{padding:0;overflow:hidden}
.fila-topo{display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:14px;padding:16px 20px;border-bottom:1px solid var(--line)}
.fila-titulo{display:flex;align-items:baseline;gap:12px;flex-wrap:wrap}
.fila-titulo h2{font-size:16px}
.fila-resumo{font-size:12.5px;color:var(--ink2)}
.fila-resumo b{font-weight:600;color:var(--ink)}
.fila-dir{display:flex;flex-wrap:wrap;align-items:center;gap:16px}
.em-jogo{display:flex;align-items:baseline;gap:8px;background:var(--surface2);border:1px solid var(--line2);border-radius:8px;padding:6px 10px}
.em-jogo .rot-e{font-size:11px;color:var(--ink3)}
.em-jogo .mono{font-size:13px;font-weight:600}
.legenda{display:flex;flex-wrap:wrap;gap:12px;font-size:11.5px;color:var(--ink3)}
.legenda span{display:inline-flex;align-items:center}
.legenda .dot{margin:0 6px 0 0}
.tabela-scroll{overflow:auto;max-height:calc(100vh - 220px)}
table{border-collapse:collapse;width:100%}
.tabela-scroll table{min-width:1180px}
thead th{position:sticky;top:0;z-index:1;background:var(--surface2);border-bottom:1px solid var(--line);font:600 10px/1 'IBM Plex Sans',sans-serif;
  letter-spacing:.07em;text-transform:uppercase;color:var(--ink3);padding:10px 12px;text-align:right;white-space:nowrap}
thead th.esq,td.esq{text-align:left}
tbody td{padding:9px 12px;border-bottom:1px solid var(--line2);white-space:nowrap;text-align:right;vertical-align:middle}
tbody tr:hover{background:var(--surface2)}
td.num{font-variant-numeric:tabular-nums}
td.tem-detalhe{cursor:help}
td.c-chk,th.c-chk{width:36px;text-align:center;padding-left:14px}
td.desc{text-align:left;font-size:12px;color:var(--ink2)}
td.ped{text-align:left}
.v{font-size:12.5px;font-variant-numeric:tabular-nums}
.v.pos{color:var(--ok)}.v.neg{color:var(--bad)}.v.zero{color:var(--ink2)}
.v.nao-achado{color:var(--warn)}
td.dif .v{font-size:13px;font-weight:600}
.vazio{color:var(--ink3)}
.pedido{background:none;border:none;padding:0;cursor:pointer;font-size:12.5px;font-weight:500;color:var(--ink)}
.pedido:hover{color:var(--accent)}
.nat{display:inline-flex;align-items:center;gap:7px;cursor:pointer}
.pilha{display:flex;flex-direction:column;gap:5px;align-items:flex-end}
tr.linha-feita{opacity:.45;background:var(--surface2)}
tr.linha-feita .pedido{text-decoration:line-through}
.vazio-fila{padding:64px 20px;text-align:center}
.vazio-fila .grande{font:600 18px/1.3 'IBM Plex Sans',sans-serif;margin-bottom:6px}
.vazio-fila .apoio{font-size:13px;color:var(--ink2)}
.sel-natureza-extra{background:var(--surface2);color:var(--ink);border:1px solid var(--line);border-radius:6px;font-size:11.5px;padding:3px 5px;max-width:130px}

/* aba AB / RBX / RB */
.lado-a-lado{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.lista{max-height:420px;overflow:auto;display:flex;flex-direction:column}
.li{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:8px 0;border-bottom:1px solid var(--line2)}
.rolagem{overflow:auto}
.tab{width:100%}
.tab td,.tab th{padding:9px 12px}
.tab td{font-size:12.5px}
.ab-cru{max-height:420px}
.ab-cru table{min-width:1600px}
.ab-cru td,.ab-cru th{font-size:12px}
.vazio-msg{margin:8px 0 0}

/* drawer de copia */
.drawer{position:fixed;left:0;right:0;bottom:0;background:var(--surface);border-top:1px solid var(--line);
  box-shadow:0 -10px 30px -16px rgba(0,0,0,.35);display:none;z-index:20}
.drawer.on{display:block}
.drawer-in{max-width:1560px;margin:0 auto;padding:12px 24px 16px}
.drawer-topo{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:10px}
.drawer-topo b{font:600 13px/1 'IBM Plex Sans',sans-serif}
.drawer-topo .apoio{margin-left:12px}
.drawer-paineis{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.drawer.min .drawer-paineis{display:none}
.copia{border:1px solid var(--line);border-radius:10px;background:var(--surface2);padding:10px 12px}
.copia-topo{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px}
.copia-topo b{font:600 11.5px/1 'IBM Plex Sans',sans-serif}
.btn-copiar{font:600 11.5px/1 'IBM Plex Sans',sans-serif;color:var(--surface);background:var(--accent);border:none;border-radius:6px;padding:7px 11px;cursor:pointer}
.btn-copiar.feito{background:var(--ok)}
.copia textarea{width:100%;height:104px;resize:vertical;white-space:pre;border:none;background:transparent;color:var(--ink);font-size:11.5px;outline:none}

/* toast */
.toast{position:fixed;top:16px;right:16px;background:var(--ink);color:var(--bg);font:600 12px/1 'IBM Plex Sans',sans-serif;
  padding:11px 14px;border-radius:9px;opacity:0;pointer-events:none;transition:opacity .12s;z-index:40}
.toast.on{opacity:1}

@media (max-width:1100px){  /* notebook de 1280 continua em colunas; so quebra abaixo disso */
  .faixa{grid-template-columns:1fr}
  .faixa-veredito{border-right:none;border-bottom:1px solid var(--line)}
  .duas{grid-template-columns:1fr}
  .lado-a-lado{grid-template-columns:1fr}
}
"""

_HTML = r"""<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>@@TITULO_ESC@@</title>
<script>try{var t=localStorage.getItem("concfull:tema");if(t==="light"||t==="dark")document.documentElement.setAttribute("data-theme",t);}catch(e){}</script>
<style>@@CSS@@</style>
</head>
<body>
<div class="pagina"><div class="conteudo">

  <div class="topo">
    <div>
      <div class="titulo-linha"><h1>@@TITULO_ESC@@</h1><span class="chip">RB ref: @@RB_REF@@</span></div>
      <div class="sub">Mercado Livre · canal @@CANAL@@ · Contas a Receber · última atualização @@HORA@@</div>
    </div>
    <div class="acoes">@@BOTOES@@<button type="button" class="btn tema" id="btn-tema" title="Alternar tema">auto</button></div>
  </div>

  <div class="abas">
    <button class="aba-btn ativo" data-aba="guia-1">Conciliação</button>
    <button class="aba-btn" data-aba="guia-2">AB / RBX / RB</button>
  </div>

  <div class="guia" id="guia-1">
    @@AVISOS@@
    @@FAIXA@@
    <div class="duas@@DUAS_CLASSE@@">@@FLUXO@@@@HISTORICO@@</div>

    <section class="card fila">
      <div class="fila-topo">
        <div class="fila-titulo"><h2>Fila VERIFICAR</h2>
          <span class="fila-resumo">@@FILA_RESUMO@@</span></div>
        <div class="fila-dir">
          <div class="em-jogo"><span class="rot-e">R$ em jogo</span><span class="mono">@@EM_JOGO@@</span></div>
          <div class="legenda"><span><i class="dot ok"></i>fecha</span><span><i class="dot bad"></i>sobra diferença</span><span><i class="dot warn"></i>não encontrado</span></div>
        </div>
      </div>
      @@FILA_CORPO@@
    </section>
  </div>

  <div class="guia" id="guia-2" hidden>@@GUIA2@@</div>

</div></div>

<div class="drawer" id="drawer">
  <div class="drawer-in">
    <div class="drawer-topo">
      <div><b>Pronto para colar no TOTVS</b><span class="apoio" id="drawer-resumo"></span></div>
      <button type="button" class="btn" id="btn-drawer">Minimizar</button>
    </div>
    <div class="drawer-paineis">
      <div class="copia">
        <div class="copia-topo"><b>Colar no CMM (Movimento_Bancario)</b>
          <button type="button" class="btn-copiar" onclick="copiarNatureza('ta-natureza-cmm', this)">Copiar CMM</button></div>
        <textarea id="ta-natureza-cmm" readonly></textarea>
      </div>
      <div class="copia">
        <div class="copia-topo"><b>Colar no MOV</b>
          <button type="button" class="btn-copiar" onclick="copiarNatureza('ta-natureza-mov', this)">Copiar MOV</button></div>
        <textarea id="ta-natureza-mov" readonly></textarea>
      </div>
    </div>
  </div>
</div>
<div class="toast" id="toast"></div>
<script>@@JS@@</script>
</body>
</html>"""

_JS = r"""
(function () {
  var CHAVE = @@CHAVE@@;
  var marcados;
  try { marcados = new Set(JSON.parse(localStorage.getItem(CHAVE) || "[]")); }
  catch (e) { marcados = new Set(); }

  function salvar() {
    try { localStorage.setItem(CHAVE, JSON.stringify(Array.from(marcados))); } catch (e) {}
  }

  // ---- toast + copiar (clipboard com fallback pra textarea+execCommand) ----
  var toastEl = document.getElementById("toast");
  var toastTimer = null;
  function toast(texto) {
    toastEl.textContent = texto;
    toastEl.classList.add("on");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { toastEl.classList.remove("on"); }, 1400);
  }
  function copiar(texto, aviso) {
    function fallback() {
      var ta = document.createElement("textarea");
      ta.value = texto; ta.style.position = "fixed"; ta.style.opacity = "0";
      document.body.appendChild(ta); ta.select();
      try { document.execCommand("copy"); } catch (e) {}
      document.body.removeChild(ta);
    }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(texto).catch(fallback);
    } else { fallback(); }
    toast(aviso || "Copiado!");
  }
  window.copiar = copiar;

  // pedido / rotulos do historico TOTVS: 1 clique copia
  document.querySelectorAll("[data-copiar]").forEach(function (b) {
    b.addEventListener("click", function () {
      copiar(b.dataset.copiar);
      if (b.classList.contains("hist")) {
        b.classList.add("copiado");
        setTimeout(function () { b.classList.remove("copiado"); }, 1600);
      }
    });
  });

  // ---- formatacao e contadores da fila (restam / ajuste pendente) ----
  function fmtMoeda(v) {
    var neg = v < 0;
    var s = Math.abs(v).toFixed(2).replace(".", ",").replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    return (neg ? "−" : "") + "R$ " + s;
  }
  function atualizarResumo() {
    var restam = 0, ajuste = 0;
    document.querySelectorAll("#tabela-verificar tbody tr[data-pedido]").forEach(function (tr) {
      if (!tr.classList.contains("linha-feita")) { restam++; ajuste += parseFloat(tr.dataset.dif || "0"); }
    });
    var a = document.getElementById("fila-restam"), b = document.getElementById("fila-ajuste");
    if (a) a.textContent = restam;
    if (b) b.textContent = fmtMoeda(ajuste);
  }

  document.querySelectorAll(".chk-ajustado").forEach(function (chk) {
    var linha = chk.closest("tr");
    if (marcados.has(chk.dataset.pedido)) {
      chk.checked = true;
      linha.classList.add("linha-feita");
    }
    chk.addEventListener("change", function () {
      if (chk.checked) { marcados.add(chk.dataset.pedido); linha.classList.add("linha-feita"); }
      else { marcados.delete(chk.dataset.pedido); linha.classList.remove("linha-feita"); }
      salvar();
      atualizarResumo();
    });
  });
  atualizarResumo();

  // Copiar da tabela pro Excel: no tema escuro o texto copiado vinha claro
  // (invisivel no Excel branco) -- entrega TSV puro, com apostrofo no pedido
  // (senao 16 digitos viram notacao cientifica).
  document.addEventListener("copy", function (e) {
    var tabela = document.getElementById("tabela-verificar");
    var sel = window.getSelection();
    if (!tabela || !sel || sel.isCollapsed || !sel.anchorNode || !tabela.contains(sel.anchorNode)) return;
    var linhas = [];
    tabela.querySelectorAll("tbody tr").forEach(function (tr) {
      if (!sel.containsNode(tr, true)) return;
      linhas.push(Array.prototype.map.call(tr.cells, function (td) {
        var t = td.innerText.replace(/\s+/g, " ").trim();
        return td.querySelector(".pedido") ? "'" + t : t;
      }).join("\t"));
    });
    if (!linhas.length) return;
    e.clipboardData.setData("text/plain", linhas.join("\n"));
    e.preventDefault();
  });

  // Abas (Conciliacao / AB-RBX-RB) -- guia ativa persiste em localStorage
  // (chave separada de CHAVE). O reload do visibilitychange acontece toda
  // vez que o Everton volta pra aba (nao so entre dias, ver mais abaixo) --
  // sem persistir, ele perdia a guia 2 toda vez que saia pra colar algo no
  // TOTVS e voltava. Pedido do Everton, 18/09/2026.
  var CHAVE_ABA = CHAVE + "_aba";

  function ativarAba(nomeAba) {
    document.querySelectorAll(".aba-btn").forEach(function (b) {
      b.classList.toggle("ativo", b.dataset.aba === nomeAba);
    });
    document.querySelectorAll(".guia").forEach(function (g) {
      g.hidden = g.id !== nomeAba;
    });
  }

  document.querySelectorAll(".aba-btn").forEach(function (botao) {
    botao.addEventListener("click", function () {
      ativarAba(botao.dataset.aba);
      try { localStorage.setItem(CHAVE_ABA, botao.dataset.aba); } catch (e) {}
    });
  });

  try {
    var abaSalva = localStorage.getItem(CHAVE_ABA);
    if (abaSalva && document.getElementById(abaSalva)) { ativarAba(abaSalva); }
  } catch (e) {}

  // Volta pra essa aba do navegador (depois de rodar o processo de novo,
  // outro canal ou nao) -> recarrega sozinho, sem precisar abrir outro link.
  document.addEventListener("visibilitychange", function () {
    if (document.visibilityState === "visible") { location.reload(); }
  });

  // Tema: auto (segue o sistema) -> claro -> escuro
  var TEMAS = ["auto", "light", "dark"], ROTULOS = {auto: "auto", light: "claro", dark: "escuro"};
  var temaAtual = "auto";
  try { var ts = localStorage.getItem("concfull:tema"); if (ts === "light" || ts === "dark") temaAtual = ts; } catch (e) {}
  var btnTema = document.getElementById("btn-tema");
  function aplicarTema() {
    if (temaAtual === "auto") document.documentElement.removeAttribute("data-theme");
    else document.documentElement.setAttribute("data-theme", temaAtual);
    btnTema.textContent = ROTULOS[temaAtual];
  }
  btnTema.addEventListener("click", function () {
    temaAtual = TEMAS[(TEMAS.indexOf(temaAtual) + 1) % TEMAS.length];
    try { localStorage.setItem("concfull:tema", temaAtual); } catch (e) {}
    aplicarTema();
  });
  aplicarTema();

  // Atualizar (RB/RBX/AB): a pagina nao consegue rodar o .bat sozinha --
  // copia o caminho dele e lembra de dar 2 cliques (uma pasta acima).
  var btnAtualizar = document.getElementById("btn-atualizar");
  if (btnAtualizar) {
    btnAtualizar.addEventListener("click", function () {
      var bat = btnAtualizar.dataset.bat;
      if (bat) { copiar(bat, "Caminho copiado — dê 2 cliques no .bat (uma pasta acima)"); }
      else { location.reload(); }
    });
  }

  // Marcacao de NATUREZA (Frete/Rebate) -- persiste em localStorage tambem
  // (chave separada de CHAVE), porque o visibilitychange reload (acima)
  // acontece a cada troca de aba, nao so entre sessoes -- sem isso o
  // Everton perdia a marcacao so de alternar pro TOTVS colar e voltar.
  // Chave e pedido+motivo (nao so pedido -- o Frete pode ter 2 checkboxes
  // no mesmo pedido, custo de envio + tarifa, ver _conteudo_natureza).
  // Pedido do Everton, 17/09/2026.
  var RESPONSAVEL = @@RESPONSAVEL@@;
  var DATA_LACTO = @@DATA_LACTO@@;
  var ORIGEM = "MKP";
  var PAGADOR = "MELI";
  var LOJA = @@LOJA@@;
  // Vem de CANAIS[canal]["vendedor"] (FULL 751, ES 763), o mesmo do
  // rebate_cmm.py -- antes era "751" fixo pra qualquer canal.
  var VENDEDOR = @@VENDEDOR@@;
  // Usado pra montar o HISTORICO das marcacoes extras (Mediacao/Cashback/
  // etc, ver NATUREZAS_EXTRAS): prefixo_mov + "_" + motivo da opcao
  // escolhida, mesmo padrao do Frete/Comissao/Rebate (calculado em Python
  // la, calculado aqui em JS porque o motivo so se sabe depois da escolha).
  var PREFIXO_MOV = @@PREFIXO_MOV@@;

  var CHAVE_NATUREZA = CHAVE + "_natureza";
  var marcadosNatureza;
  try { marcadosNatureza = new Set(JSON.parse(localStorage.getItem(CHAVE_NATUREZA) || "[]")); }
  catch (e) { marcadosNatureza = new Set(); }

  function chaveNatureza(chk) { return chk.dataset.pedido + "|" + chk.dataset.motivo; }

  function salvarNatureza() {
    try { localStorage.setItem(CHAVE_NATUREZA, JSON.stringify(Array.from(marcadosNatureza))); } catch (e) {}
  }

  // Gaveta de copia: minimizada persiste (mesmo motivo do resto)
  var CHAVE_DRAWER = CHAVE + "_drawer";
  var drawerMin = false;
  try { drawerMin = localStorage.getItem(CHAVE_DRAWER) === "min"; } catch (e) {}
  var drawer = document.getElementById("drawer");
  var btnDrawer = document.getElementById("btn-drawer");
  function aplicarDrawer(aberta) {
    drawer.classList.toggle("on", aberta);
    drawer.classList.toggle("min", drawerMin);
    document.body.classList.toggle("drawer-aberto", aberta && !drawerMin);
    document.body.classList.toggle("drawer-min", aberta && drawerMin);
    btnDrawer.textContent = drawerMin ? "Abrir painéis" : "Minimizar";
  }
  btnDrawer.addEventListener("click", function () {
    drawerMin = !drawerMin;
    try { localStorage.setItem(CHAVE_DRAWER, drawerMin ? "min" : "aberto"); } catch (e) {}
    aplicarDrawer(true);
  });

  function atualizarNatureza() {
    var marcados = Array.prototype.slice.call(document.querySelectorAll(".chk-natureza:checked"));
    // Marcacoes extras (Mediacao/Cashback/etc, ver NATUREZAS_EXTRAS): o
    // <select> nao e um checkbox, entao monta um objeto com o mesmo
    // formato (.dataset.*) esperado pelas expressoes de @@CAMPOS_CMM_JS@@
    // (chk.dataset.motivo etc.) -- assim entra na mesma lista sem duplicar
    // a logica de montar linha do CMM/MOV.
    document.querySelectorAll(".sel-natureza-extra").forEach(function (sel) {
      if (!sel.value) return;
      var opt = sel.options[sel.selectedIndex];
      marcados.push({dataset: {
        pedido: sel.dataset.pedido, motivo: opt.dataset.motivo, natureza: opt.dataset.natureza,
        conta: sel.dataset.conta, banco: sel.dataset.banco, agencia: sel.dataset.agencia,
        historico: PREFIXO_MOV ? (PREFIXO_MOV + "_" + opt.dataset.motivo) : "",
        valor: sel.dataset.valor
      }});
    });
    if (!marcados.length) { aplicarDrawer(false); return; }

    var linhasCmm = marcados.map(function (chk) {
      return [@@CAMPOS_CMM_JS@@].join("\t");
    });
    document.getElementById("ta-natureza-cmm").value = linhasCmm.join("\n");

    var linhasMov = marcados.map(function (chk) {
      return [
        DATA_LACTO, chk.dataset.natureza, "'" + chk.dataset.pedido, "",
        chk.dataset.banco, chk.dataset.agencia, chk.dataset.conta,
        chk.dataset.historico, chk.dataset.valor
      ].join("\t");
    });
    document.getElementById("ta-natureza-mov").value = linhasMov.join("\n");

    var pedidos = {};
    marcados.forEach(function (c) { pedidos[c.dataset.pedido] = true; });
    document.getElementById("drawer-resumo").textContent =
      Object.keys(pedidos).length + " linha(s) marcada(s) · CMM " + linhasCmm.length + " · MOV " + linhasMov.length;
    aplicarDrawer(true);
  }

  document.querySelectorAll(".chk-natureza").forEach(function (chk) {
    if (marcadosNatureza.has(chaveNatureza(chk))) { chk.checked = true; }
    chk.addEventListener("change", function () {
      if (chk.checked) { marcadosNatureza.add(chaveNatureza(chk)); }
      else { marcadosNatureza.delete(chaveNatureza(chk)); }
      salvarNatureza();
      atualizarNatureza();
    });
  });

  // Persistencia das marcacoes extras (select), mesmo motivo do resto: o
  // reload do visibilitychange (toda vez que o Everton troca de aba)
  // perderia a escolha sem isso. Chave por pedido (nao por pedido+motivo
  // como o Set dos checkboxes) -- so 1 escolha possivel por pedido aqui.
  var CHAVE_NATUREZA_EXTRA = CHAVE + "_natureza_extra";
  var marcadosNaturezaExtra;
  try { marcadosNaturezaExtra = JSON.parse(localStorage.getItem(CHAVE_NATUREZA_EXTRA) || "{}"); }
  catch (e) { marcadosNaturezaExtra = {}; }

  function salvarNaturezaExtra() {
    try { localStorage.setItem(CHAVE_NATUREZA_EXTRA, JSON.stringify(marcadosNaturezaExtra)); } catch (e) {}
  }

  document.querySelectorAll(".sel-natureza-extra").forEach(function (sel) {
    var salvo = marcadosNaturezaExtra[sel.dataset.pedido];
    if (salvo) {
      for (var i = 0; i < sel.options.length; i++) {
        if (sel.options[i].value === salvo) { sel.value = salvo; break; }
      }
    }
    sel.addEventListener("change", function () {
      if (sel.value) { marcadosNaturezaExtra[sel.dataset.pedido] = sel.value; }
      else { delete marcadosNaturezaExtra[sel.dataset.pedido]; }
      salvarNaturezaExtra();
      atualizarNatureza();
    });
  });

  atualizarNatureza();
})();

function copiarNatureza(id, botao) {
  var ta = document.getElementById(id);
  window.copiar(ta.value, "Copiado!");
  if (botao) {
    var textoOriginal = botao.textContent;
    botao.textContent = "Copiado!";
    botao.classList.add("feito");
    setTimeout(function () { botao.textContent = textoOriginal; botao.classList.remove("feito"); }, 1600);
  }
}
"""


def _preencher(template, valores):
    for chave, valor in valores.items():
        template = template.replace(f"@@{chave}@@", valor)
    return template


def gerar_dashboard(resultado, fila_verificar, rebate_cmm, *, canal, data, data_ref_rb,
                     caminho_xlsx_completo, caminho_xlsx_rebate, pasta_saida, alvo=None,
                     ml_atualizado=False, caminho_bat_scraper=None, raiz=None,
                     caminho_xlsx_descricoes=None, conferencia_rbx=None, caminho_bat_atualizar=None,
                     caminho_bat_arquivar=None, ml_de_cache=False, responsavel=None,
                     rebate_processar=None, caminho_xlsx_rebate_processar=None,
                     para_processar=None, lotes_status=None, valores_pedidos=None):
    # NATUREZA (Frete/Rebate marcados na fila VERIFICAR) -- HISTORICO/LOJA
    # derivados de CANAIS[canal] direto (nao do parametro `raiz`, que fica
    # None pros canais em CANAIS_SEM_RAIZ_CONFIRMADA -- essa marcacao nao
    # depende daquela confirmacao especifica). loja_calendario, nao
    # raiz.split("_")[-1] (dava "FU" pro FULL -- "FU" e o chave_rbx, usado
    # pra achar linha no extrato, nao o valor certo do campo LOJA do
    # CMM/MOV). Confirmado com o Everton, 18/09/2026.
    cfg_canal = CANAIS.get(canal, {})
    prefixo_mov = cfg_canal.get("prefixo_mov")
    loja = cfg_canal.get("loja_cmm", "")  # nao loja_calendario ("COLETAS ES" no ES; o CMM usa "CO")
    vendedor = cfg_canal.get("vendedor", "")  # FULL "751", ES "763" (confirmado, 19/09/2026)
    historicos = (
        {
            "frete": f"{prefixo_mov}_{MOTIVO_FRETE}",
            "comissao": f"{prefixo_mov}_{MOTIVO_COMISSAO}",
            "rebate": f"{prefixo_mov}_{MOTIVO_REBATE}",
        }
        if prefixo_mov else {}
    )
    campos_cmm_js = ", ".join(_CAMPOS_CMM_JS[c] for c in COLUNAS_CMM)

    rebate_total = resultado.loc[resultado["STATUS"] == "REBATE", "REBATE"].sum()
    rb_total = resultado["TOTAL_RB"].sum()
    ajuste_pendente = fila_verificar["DIF"].sum()
    verificar_qtd = len(fila_verificar)
    em_jogo = float(fila_verificar["DIF"].abs().sum()) if verificar_qtd else 0.0

    faixa_html = _faixa_status(alvo, rb_total, rebate_total, ajuste_pendente, conferencia_rbx)
    fluxo_html = _fluxo_original_transitoria(conferencia_rbx, alvo) if conferencia_rbx else ""
    historico_html = (
        '<section class="card historico"><div class="card-topo"><h2>Histórico TOTVS</h2>'
        f'<span class="apoio">{html.escape(canal)} · clique para copiar</span></div>'
        f'<div class="hist-lista">{_painel_historico(raiz)}</div></section>'
    ) if raiz else ""
    duas_classe = " uma" if not (fluxo_html and historico_html) else ""

    if verificar_qtd:
        colunas_extra = "<th>Marcação</th>" if NATUREZAS_EXTRAS else ""
        fila_corpo = f"""
      <div class="tabela-scroll">
        <table id="tabela-verificar">
          <thead><tr>
            <th class="c-chk" title="Ajustado">Aj.</th><th>RBX</th><th>RB</th><th>Dif. extrato</th><th>AB total</th>
            <th class="esq">Pedido</th><th class="esq">Descrição</th><th>DIF</th>{colunas_extra}
            <th>Comissão</th><th>Desc/Bônus</th><th>Frete</th><th>Repasse</th><th>Rebate</th>
          </tr></thead>
          <tbody>{_linhas_tabela(fila_verificar, conferencia_rbx, historicos, valores_pedidos)}</tbody>
        </table>
      </div>"""
        fila_resumo = (f'{verificar_qtd} pedidos · restam <b id="fila-restam">{verificar_qtd}</b> · '
                       f'ajuste pendente <b id="fila-ajuste" class="mono">{_moeda(ajuste_pendente)}</b>')
    else:
        fila_corpo = ('<div class="vazio-fila"><div class="grande">Nenhum pedido em VERIFICAR 🎉</div>'
                      '<div class="apoio">RB, RBX e AB fecharam com o calendário. Nada para ajustar manualmente hoje.</div></div>')
        fila_resumo = "nenhum pedido"

    botoes = [f'<a class="btn" href="{html.escape(caminho_xlsx_completo)}">Baixar planilha completa</a>']
    if len(rebate_cmm):
        botoes.append(f'<a class="btn" href="{html.escape(caminho_xlsx_rebate)}">XLSX Rebate CMM ({len(rebate_cmm)})</a>')
    if caminho_xlsx_rebate_processar and rebate_processar is not None and len(rebate_processar):
        botoes.append(
            f'<a class="btn" href="{html.escape(caminho_xlsx_rebate_processar)}">'
            f'PROCESSAR_REBATE ({len(rebate_processar)})</a>'
        )
    if caminho_xlsx_descricoes:
        botoes.append(f'<a class="btn" href="{html.escape(caminho_xlsx_descricoes)}">Descrições ML/MP</a>')
    botoes.append(
        '<button type="button" class="btn primario" id="btn-atualizar"'
        + (f' data-bat="{html.escape(caminho_bat_atualizar)}" title="Copia o caminho do .bat (uma pasta acima)"' if caminho_bat_atualizar else "")
        + '>Atualizar (RB/RBX/AB)</button>'
    )

    fonte_valores_guia2 = conferencia_rbx or valores_pedidos or {}
    valores_rbx = fonte_valores_guia2.get("valores_rbx") or {}
    valores_rb = fonte_valores_guia2.get("valores_rb") or {}
    guia2_html = f"""
    <section class="card"><div class="card-topo"><h2>Lotes do pipeline AB</h2></div>{_tabela_lotes(lotes_status)}</section>
    <div class="lado-a-lado">{_cartao_conta("RBX", valores_rbx)}{_cartao_conta("RB", valores_rb)}</div>
    <section class="card"><div class="card-topo"><h2>AB para processar</h2>
      <span class="apoio mono">{len(para_processar) if para_processar is not None else 0} linhas</span></div>
      {_tabela_para_processar(para_processar)}</section>"""

    titulo = f"Conciliação {canal} — {data}"
    chave_local_storage = f"ajustados_{canal.lower()}_{data.replace('.', '')}"
    # Avisos discretos do topo -- so informam de passagem, sem competir por
    # atencao (pedido do Everton, 17/09/2026). Os avisos "Scraper pronto" e
    # "ML de cache" foram removidos em 18/09/2026 (ruido sem uso);
    # ml_de_cache/caminho_bat_scraper continuam aceitos (interativo.py ainda
    # gera o .bat na RAIZ normalmente).
    avisos = []
    if caminho_bat_arquivar:
        # Destacado (verde) -- e o unico aviso que exige uma acao concreta do
        # Everton (arquivar a copia do RBX que motivou a transferencia). So
        # aparece quando a Transitoria ja bateu com o alvo (ver
        # executar()/conferir_rbx.pronta_para_transferencia). Pedido do
        # Everton, 17/09/2026.
        avisos.append(
            '<div class="aviso ok"><span class="dot"></span>'
            f'Transferência pronta! Arquivar RBX: {html.escape(caminho_bat_arquivar)} '
            '(uma pasta acima — dê 2 cliques nele)</div>'
        )
    if ml_atualizado:
        avisos.append('<div class="aviso"><span class="dot ok"></span>Valores do Mercado Livre atualizados</div>')
    avisos_html = f'<div class="avisos">{"".join(avisos)}</div>' if avisos else ""

    js = _preencher(_JS, {
        "CHAVE": json.dumps(chave_local_storage),
        "RESPONSAVEL": json.dumps(responsavel or ""),
        "DATA_LACTO": json.dumps(data_ref_rb or ""),
        "LOJA": json.dumps(loja),
        "VENDEDOR": json.dumps(vendedor),
        "PREFIXO_MOV": json.dumps(prefixo_mov or ""),
        "CAMPOS_CMM_JS": campos_cmm_js,
    })
    corpo = _preencher(_HTML, {
        "TITULO_ESC": html.escape(titulo), "CANAL": html.escape(canal),
        "RB_REF": html.escape(data_ref_rb or "?"), "HORA": datetime.datetime.now().strftime("%H:%M"),
        "BOTOES": "".join(botoes), "AVISOS": avisos_html, "FAIXA": faixa_html,
        "DUAS_CLASSE": duas_classe, "FLUXO": fluxo_html, "HISTORICO": historico_html,
        "FILA_RESUMO": fila_resumo, "EM_JOGO": _moeda(em_jogo), "FILA_CORPO": fila_corpo,
        "GUIA2": guia2_html, "CSS": _CSS,
    })
    corpo = corpo.replace("@@JS@@", js)  # por ultimo: o JS tem @@ proprios ja resolvidos acima

    os.makedirs(pasta_saida, exist_ok=True)
    caminho_html = os.path.join(pasta_saida, f"dashboard_{data.replace('.', '')}.html")
    with open(caminho_html, "w", encoding="utf-8") as f:
        f.write(corpo)
    return caminho_html


if __name__ == "__main__":
    import tempfile

    import pandas as pd

    # soma dos absolutos, nao o absoluto da soma: |-30|+|20|+|-10| = 60
    assert _repasse_calculado(-30.0, 20.0, -10.0) == 60.0
    assert _repasse_calculado(None, 20.0, -10.0) is None
    # formatacao: menos tipografico (U+2212), milhar com ponto, decimal com virgula
    assert _numero_br(-1234.5) == "−1.234,50" and _moeda(-921.84) == "−R$ 921,84"
    assert _formatar_numero(-25.5) == "-25,50"  # pro CMM/MOV o hifen e ASCII

    # Split Frete/Tarifa de cobranca (pedido isolado, sem passar pelo resto
    # do dashboard): ML_ENVIO=-30 x ML_CUSTO_ENVIO=-22 -> tarifa=-8, tem que
    # gerar 2 checkboxes na MESMA celula Frete (2209/X103925 pro custo de
    # envio, 2804/X103925 pra tarifa -- SEMPRE Transitoria, mesmo o CMM real
    # mostrando "103925" naquelas linhas, ver CONTA_TRANSITORIA), sem
    # quebrar a contagem de <td> da linha (13, igual as demais colunas).
    linhas_split = _linhas_tabela(
        pd.DataFrame({
            "PEDIDO_MKP": ["4"], "DIF": [-50.0],
            "ML_ENVIO": [-30.0], "ML_CUSTO_ENVIO": [-22.0], "ML_REBATE": [0.0],
        }),
        historicos={"frete": "MKP_ME_FULL_DESPESAS DE FRETE ENVIO",
                    "comissao": "MKP_ME_FULL_TARIFA DE COBRANÇA"},
    )
    assert linhas_split.count("<td") == 13
    assert ('data-motivo="DESPESAS DE FRETE ENVIO" data-natureza="2209" '
            'data-conta="X103925"') in linhas_split
    assert 'data-valor="-22,00"' in linhas_split
    assert ('data-motivo="TARIFA DE COBRANÇA" data-natureza="2804" '
            'data-conta="X103925"') in linhas_split
    assert 'data-historico="MKP_ME_FULL_TARIFA DE COBRANÇA"' in linhas_split
    assert 'data-valor="-8,00"' in linhas_split
    assert 'class="pilha"' in linhas_split  # 2 fretes empilhados na mesma celula
    assert 'data-copiar="\'4"' in linhas_split  # pedido copia com apostrofo

    # Diferenca extrato: quando o pedido JA tem TOTAL_RB != 0 (achado no
    # cruzamento inicial contra CONCILIADOR_RB.xlsx, ver cruzamento.py),
    # RB/RBX mostra a MESMA transacao -- Diferenca extrato tem que ser so o
    # DIF puro, sem descontar RB/RBX de novo (senao desconta a mesma grana
    # 2x). Cenario real (pedido 2000018137884794, 09/09/2026): DIF=371.65
    # (= REPASSE 2283.08 - TOTAL_RB 1911.43), RBX=1911.43 (mesma transacao)
    # -- tinha que mostrar 371.65 (o que ainda falta lancar), nao -1539.78
    # (desconto duplicado). Pedido do Everton, 18/09/2026.
    linhas_duplicado = _linhas_tabela(
        pd.DataFrame({
            "PEDIDO_MKP": ["9"], "DIF": [371.65], "TOTAL_RB": [1911.43],
            "ML_ENVIO": [0.0], "ML_REBATE": [0.0],
        }),
        conferencia_rbx={"valores_rb": {}, "valores_rbx": {"9": 1911.43}},
    )
    assert _valor_html(371.65) in linhas_duplicado  # DIF puro
    assert _valor_html(-1539.78) not in linhas_duplicado  # desconto duplicado (bug antigo)
    assert '<span class="dot bad"></span>' in linhas_duplicado  # vermelho: ainda falta lancar

    # Cenario real (pedido 2000018137884794, 10/09/2026): lancamento antigo
    # (1911.43, mesmo do TOTAL_RB) foi CANCELADO e REEMITIDO como agregado
    # de 2283.08 -- RBX no liquido ja fechou o pedido, mas e maior que
    # TOTAL_RB (nao e mais uma duplicata simples). A formula unificada
    # (DIF - RB - RBX + TOTAL_RB) tem que fechar em 0 (verde) sem precisar
    # de caso especial. Pedido do Everton, 18/09/2026.
    linhas_reemitido = _linhas_tabela(
        pd.DataFrame({
            "PEDIDO_MKP": ["9"], "DIF": [371.65], "TOTAL_RB": [1911.43],
            "ML_ENVIO": [0.0], "ML_REBATE": [0.0],
        }),
        conferencia_rbx={"valores_rb": {}, "valores_rbx": {"9": 2283.08}},
    )
    assert _valor_html(0.0) in linhas_reemitido
    assert '<span class="dot ok"></span>' in linhas_reemitido  # verde: fechou

    # TOTAL_RB == 0 (pedido NAO tinha sido achado no cruzamento inicial) --
    # continua subtraindo RB/RBX normalmente (RBX aqui e informacao nova,
    # nao duplicada). DIF=371.65, RBX=371.65 -> fecha em 0 (verde).
    linhas_novo = _linhas_tabela(
        pd.DataFrame({
            "PEDIDO_MKP": ["9"], "DIF": [371.65], "TOTAL_RB": [0.0],
            "ML_ENVIO": [0.0], "ML_REBATE": [0.0],
        }),
        conferencia_rbx={"valores_rb": {}, "valores_rbx": {"9": 371.65}},
    )
    assert _valor_html(0.0) in linhas_novo
    assert '<span class="dot ok"></span>' in linhas_novo  # verde: fechou

    # Pedido nao encontrado nem no RB nem no RBX: "n/e" + bolinha ambar.
    linhas_ne = _linhas_tabela(
        pd.DataFrame({"PEDIDO_MKP": ["7"], "DIF": [16.48], "ML_ENVIO": [0.0], "ML_REBATE": [0.0]}),
        conferencia_rbx={"valores_rb": {}, "valores_rbx": {}},
    )
    assert '<span class="v nao-achado">n/e</span><span class="dot warn"></span>' in linhas_ne

    resultado = pd.DataFrame({
        "PEDIDO_MKP": ["1", "2", "3"],
        "STATUS": ["SEM DIFF", "REBATE", "VERIFICAR"],
        "REBATE": [0.0, 50.0, 0.0],
        "TOTAL_RB": [100.0, 40.0, 0.0],
        "DIF": [0.0, 0.0, -921.84],
    })
    fila = pd.DataFrame({
        "PEDIDO_MKP": ["3"], "DIF": [-921.84],
        # ML_ENVIO/ML_REBATE pra testar a marcacao de NATUREZA (Frete/Rebate)
        "ML_ENVIO": [-25.5], "ML_REBATE": [30.0],
        # DESCRICAO (coluna E do AB, ver cruzamento.py) -- texto real
        # confirmado no pedido 2000017591518558 (08/09/2026).
        "DESCRICAO": ["mediation, refund"],
        # TOTAL_PED (REPASSE+COMISSAO+FRETE, ver cruzamento.py) -- coluna
        # "AB total", pro Everton comparar RBX/RB com o que o AB/ML diz que
        # deveria ter. Pedido do Everton, 18/09/2026.
        "TOTAL_PED": [1234.56],
    })
    rebate_cmm = pd.DataFrame({"VALOR R$": [50.0]})
    conferencia_rbx = {
        "liquido_original": 0.0, "qtd_original": 0, "original_bate": False,
        # RA (-7.0) somado ao liquido_transitoria (0.0) -> simulado -7.0.
        # alvo=None nessa chamada -> pronta_para_transferencia e sempre
        # False (regra: so compara quando tem alvo do calendario). -7.0
        # tambem nao e zero -> zerada=False.
        "liquido_transitoria": 0.0, "qtd_transitoria": 0,
        "liquido_transitoria_simulado": -7.0, "qtd_transitoria_simulado": 1,
        "zerada": False, "pronta_para_transferencia": False,
        "liquido_ra": -7.0, "qtd_ra": 1,
        # Sem cancelamento na Original nesse fixture base (qtd=0) ->
        # valor_transferencia == liquido_transitoria_simulado (sem ajuste).
        # Ver cenario dedicado de cancelamento mais abaixo.
        "liquido_cancelado_original": 0.0, "qtd_cancelado_original": 0,
        "cancelados_original": [], "valor_transferencia": -7.0,
        # pedido "3": achado no RB com o valor exato da DIF (debito) -> Diferenca
        # extrato deve fechar em 0. Nao achado no RBX.
        "valores_rb": {"3": -921.84}, "valores_rbx": {},
        # Rebate (formula) = 50.0 (pedido "2") -- bate com RBX: 50.0 via CMM,
        # nada via cupom direto do ML.
        "liquido_cupom_rebate": 0.0, "qtd_cupom_rebate": 0,
        "liquido_cmm_rebate": 50.0, "qtd_cmm_rebate": 1,
        "liquido_rebate_rbx": 50.0, "rebate_bate": True,
    }
    para_processar = pd.DataFrame({
        # SOURCE_ID/EXTERNAL_REFERENCE como float64 (None em alguma linha
        # forca a coluna inteira pra float, igual acontece de verdade no AB
        # real) -- tem que sair como ID puro ("2000017591558558"), nao
        # formatado como dinheiro ("2.000.017.591.558.558,00"). Bug real
        # reportado pelo Everton, 18/09/2026.
        "DATE": ["01/01/2026", "01/01/2026"],
        "SOURCE_ID": [2000017591558558.0, None],
        "EXTERNAL_REFERENCE": [2000017591558558.0, 3.0],
        "NET_CREDIT_AMOUNT": [None, None],
        "NET_DEBIT_AMOUNT": [-921.84, -10.0],
        "DESCRIPTION": ["shipping", "tax"],
    })
    lotes_status = [("01.01.2026", "rodado", 3, 1, "2026-01-01T10:00:00")]

    with tempfile.TemporaryDirectory() as tmp:
        caminho = gerar_dashboard(
            resultado, fila, rebate_cmm, canal="FULL", data="01.01",
            data_ref_rb="01/01/2026", caminho_xlsx_completo="a.xlsx",
            caminho_xlsx_rebate="b.xlsx", pasta_saida=tmp, alvo=None,
            conferencia_rbx=conferencia_rbx, responsavel="EVERTON",
            rebate_processar=pd.DataFrame({"NR PEDIDO MKP": ["1"]}),
            caminho_xlsx_rebate_processar="REBATE/c.xlsx",
            para_processar=para_processar, lotes_status=lotes_status,
            caminho_bat_atualizar="../atualizar_full_0101.bat",
        )
        assert os.path.exists(caminho)
        assert os.path.basename(caminho) == "dashboard_0101.html"  # 1 link, sem canal no nome
        conteudo = open(caminho, encoding="utf-8").read()
        assert "@@" not in conteudo, "sobrou marcador do template sem preencher"
        assert "<title>Conciliação FULL — 01.01</title>" in conteudo
        assert "Fila VERIFICAR" in conteudo and "−921,84" in conteudo
        assert "Status dos pedidos" not in conteudo and "Distancia do alvo" not in conteudo
        # tokens dos 2 temas + override manual (auto/claro/escuro)
        assert "#0f1114" in conteudo and "prefers-color-scheme:dark" in conteudo
        assert 'html[data-theme="light"]' in conteudo and 'id="btn-tema"' in conteudo
        assert "fonts.googleapis" not in conteudo  # offline: nada de fonte externa
        assert "Alvo não cadastrado" in conteudo and "Sem alvo no calendário" in conteudo  # alvo=None
        assert "visibilitychange" in conteudo  # recarrega sozinho ao voltar pra aba
        assert "Acao sugerida" not in conteudo and 'class="acao"' not in conteudo  # coluna removida
        assert ">RBX<" in conteudo and ">RB<" in conteudo and "Dif. extrato" in conteudo
        assert ">AB total<" in conteudo and _valor_html(1234.56) in conteudo
        assert 'REBATE/c.xlsx">PROCESSAR_REBATE (1)</a>' in conteudo
        assert 'id="btn-atualizar" data-bat="../atualizar_full_0101.bat"' in conteudo
        # Rebate (formula) x RBX na faixa de status: bate (cmm 50.0 + cupom
        # 0.0 = 50.0, rebate_bate=True no fixture) -> pilula "confere".
        assert "Conferência do rebate" in conteudo
        assert f"RBX {_moeda(50.0)}" in conteudo
        assert "cupom 0 lançamentos · CMM 1 lançamentos" in conteudo
        assert '<span class="pill ok"><span class="dot"></span>confere</span>' in conteudo
        # pedido "3": achado no RB com o valor exato da DIF (debito) -> RB
        # vermelho (com o menos escrito), Diferenca extrato fecha em 0,00
        # (neutro) com bolinha verde. Nao achado no RBX -> celula em branco.
        tabela = conteudo.split('id="tabela-verificar"')[1].split("</table>")[0]
        assert "R$" not in tabela  # sem "R$" dentro da tabela
        assert _valor_html(-921.84) in tabela  # DIF/RB: vermelho, sinal escrito
        assert _valor_html(0.0) in tabela  # Diferenca extrato zerada: neutra
        assert '<span class="dot ok"></span>' in tabela  # bolinha verde
        assert "n/e" not in tabela  # pedido achado no RB: nao e "nao encontrado"
        assert '<span class="vazio">—</span>' in tabela  # sem valor em Desc/Bonus, etc
        # "Outros (RA)" separada nao existe mais -- ja embutida no simulado
        assert "Outros (RA)" not in conteudo
        # fluxo Original x Transitoria: passo 1 sem lancamento de
        # fechamento, passo 3 mostra o SIMULADO (0.0 + -7.0 = -7.0); alvo=None
        # nessa chamada -> nao compara contra alvo nenhum, pilula neutra
        assert "Original × Transitória" in conteudo and "sem lançamento de fechamento ainda" in conteudo
        assert _moeda(-7.0) in conteudo
        assert "sem alvo do calendário pra comparar" in conteudo
        assert "pronta para transferência" not in conteudo
        assert "zerada — transferência já feita" not in conteudo
        # Historico TOTVS so aparece com raiz confirmada
        assert "Histórico TOTVS" not in conteudo  # raiz=None nessa chamada

        # Marcacao de NATUREZA: pedido "3", Frete=-25.5 e Rebate=30.0 tem
        # que gerar checkbox com motivo/natureza/conta/historico corretos
        # (Comissao fica de fora por enquanto, ver comentario em
        # _linhas_tabela). HISTORICO usa prefixo_mov="MKP_ME_FULL" (FULL).
        assert 'data-motivo="DESPESAS DE FRETE ENVIO" data-natureza="2209" data-conta="X103925"' in conteudo
        assert 'data-historico="MKP_ME_FULL_DESPESAS DE FRETE ENVIO"' in conteudo
        assert 'data-motivo="RECEBIMENTO DE REBATE" data-natureza="1107" data-conta="X103925"' in conteudo
        assert 'data-historico="MKP_ME_FULL_RECEBIMENTO DE REBATE"' in conteudo
        assert 'data-valor="-25,50"' in conteudo and 'data-valor="30,00"' in conteudo
        assert 'chk-natureza' in conteudo and 'id="drawer"' in conteudo
        assert 'var RESPONSAVEL = "EVERTON";' in conteudo
        assert 'var LOJA = "FULL";' in conteudo  # loja_calendario, nao raiz.split("_")[-1] ("FU")
        assert 'var VENDEDOR = "751";' in conteudo
        assert "onclick=\"copiarNatureza('ta-natureza-cmm', this)\"" in conteudo
        assert 'botao.textContent = "Copiado!";' in conteudo
        # gaveta de copia: resumo + minimizar persistente
        assert "linha(s) marcada(s)" in conteudo and 'var CHAVE_DRAWER = CHAVE + "_drawer";' in conteudo
        # contadores da fila recalculam ao marcar "ajustado"
        assert 'id="fila-restam"' in conteudo and "function atualizarResumo()" in conteudo
        assert 'data-dif="-921.84"' in conteudo

        # Mesmo dashboard pro ES: LOJA "CO" (loja_cmm, nao "COLETAS ES" do
        # calendario), VENDEDOR "763" (nao o 751 fixo de antes), HISTORICO com
        # prefixo_mov "MKP_MELI_CO". Confirmado com o CMM real do ES,
        # 19/09/2026.
        caminho_es = gerar_dashboard(
            resultado, fila, rebate_cmm, canal="ES", data="02.01",
            data_ref_rb="01/01/2026", caminho_xlsx_completo="a.xlsx",
            caminho_xlsx_rebate="b.xlsx", pasta_saida=tmp, alvo=None,
            responsavel="EVERTON",
        )
        conteudo_es = open(caminho_es, encoding="utf-8").read()
        assert 'var LOJA = "CO";' in conteudo_es
        assert 'var VENDEDOR = "763";' in conteudo_es
        assert 'data-historico="MKP_MELI_CO_DESPESAS DE FRETE ENVIO"' in conteudo_es
        assert 'data-historico="MKP_MELI_CO_RECEBIMENTO DE REBATE"' in conteudo_es
        assert 'id="btn-atualizar">' in conteudo_es  # sem .bat: o botao so recarrega a pagina

        # Guia 2 (AB / RBX / RB): abas presentes, guia-2 comeca escondida
        # (hidden), lotes/AB Para Processar/RB aparecem na tabela certa.
        assert 'data-aba="guia-1"' in conteudo and 'data-aba="guia-2"' in conteudo
        assert '<div class="guia" id="guia-2" hidden>' in conteudo
        guia2 = conteudo.split('id="guia-2"')[1]
        assert "01.01.2026" in guia2 and ">processado<" in guia2  # lotes_status (rodado -> processado)
        assert "EXTERNAL_REFERENCE" in guia2 and ">3<" in guia2  # para_processar cru
        assert _valor_html(-921.84) in guia2  # RB: pedido "3" (valores_rb do fixture)
        assert "Sem rb.xlsx/rbx.xlsx pra conferir ainda." in guia2  # RBX vazio no fixture
        assert "1 lançamentos" in guia2 and "0 lançamentos" in guia2  # contagem por conta
        # SOURCE_ID/EXTERNAL_REFERENCE (float64 por causa do None numa linha)
        # tem que sair como ID puro, NAO formatado como dinheiro.
        assert '<td class="esq">2000017591558558</td>' in guia2
        assert "2.000.017.591.558.558" not in guia2  # formatacao de moeda errada
        assert "2000017591558558,00" not in guia2

        # Guia ativa persiste em localStorage (chave separada de CHAVE) --
        # sem isso o reload do visibilitychange (toda vez que o Everton
        # volta pra aba) jogava de volta pra guia 1. Pedido do Everton,
        # 18/09/2026.
        assert 'var CHAVE_ABA = CHAVE + "_aba";' in conteudo
        assert "function ativarAba(nomeAba)" in conteudo
        assert 'localStorage.setItem(CHAVE_ABA, botao.dataset.aba)' in conteudo
        assert 'localStorage.getItem(CHAVE_ABA)' in conteudo

        # Marcacao de NATUREZA persiste em localStorage (chave separada de
        # CHAVE, pra sobreviver ao reload do visibilitychange ao trocar de
        # aba) -- confere so que o JS foi gerado sem quebrar escape de
        # chaves/tab/newline (risco real nesse arquivo, ja doeu antes).
        assert 'var CHAVE_NATUREZA = CHAVE + "_natureza";' in conteudo
        assert 'function chaveNatureza(chk)' in conteudo
        assert 'marcadosNatureza.add(chaveNatureza(chk))' in conteudo
        assert conteudo.count("atualizarNatureza()") >= 1  # chamada de bootstrap no fim da IIFE
        assert '].join("\\t")' in conteudo and 'join("\\n")' in conteudo  # tab e quebra de linha no JS

        # Pedido colado no CMM/MOV com apostrofo na frente -- sem isso o
        # pedido (ate 16 digitos) vira notacao cientifica ao colar no Excel.
        assert "\"'\" + chk.dataset.pedido" in conteudo
        assert '"\'" + chk.dataset.pedido, ""' in conteudo  # linhasMov

        # pedido "3" nao tem ML_CUSTO_ENVIO -> custo_envio e None, sem split
        # (tarifa_cobranca fica None) -- continua com 1 checkbox so no Frete.
        linha3 = conteudo.split('<tr data-pedido="3"')[1].split("</tr>")[0]
        assert linha3.count("<td") == 13
        assert linha3.count('class="chk-natureza"') == 2  # so Frete + Rebate
        assert '<td class="desc">mediation, refund</td>' in linha3  # coluna Descricao (AB)

        # Faixa de status com alvo: rb_total(140) + rebate_total(50) +
        # ajuste_pendente(-921.84) = -731.84 -> bate -> "Dia fechado ✓".
        conferencia_rbx_pronta = {**conferencia_rbx, "liquido_transitoria_simulado": -731.84,
                                   "qtd_transitoria_simulado": 1, "zerada": False,
                                   "pronta_para_transferencia": True}
        caminho2 = gerar_dashboard(
            resultado, fila, rebate_cmm, canal="FULL", data="01.01",
            data_ref_rb="01/01/2026", caminho_xlsx_completo="a.xlsx",
            caminho_xlsx_rebate="b.xlsx", pasta_saida=tmp, alvo=-731.84,
            conferencia_rbx=conferencia_rbx_pronta, raiz="MKP_MELI_FU",
        )
        assert caminho2 == caminho  # mesmo arquivo, sobrescrito (mesmo dia)
        conteudo2 = open(caminho2, encoding="utf-8").read()
        assert "Dia fechado ✓" in conteudo2 and "dif. R$ 0,00" in conteudo2
        assert "Diferença de" not in conteudo2
        assert "pronta para transferência" in conteudo2 and 'class="pill accent"' in conteudo2
        # Historico TOTVS: 4 rotulos copiaveis (Raiz/Repasse/Comissao/Frete)
        assert "Histórico TOTVS" in conteudo2
        for rotulo, valor in (("Raiz", "MKP_MELI_FU"), ("Repasse", "MKP_MELI_FU_REPASSE"),
                              ("Comissão", "MKP_MELI_FU_COMISSAO"), ("Frete", "MKP_MELI_FU_FRETE")):
            assert (f'<button type="button" class="hist" data-copiar="{valor}" title="Copiar">'
                    f'<span class="hist-rot">{rotulo}</span>') in conteudo2

        # 2o estado do ciclo: liquido_transitoria_simulado=0.0 (transferencia
        # manual ja feita) -> "zerada", verde, mesmo sem bater com o alvo
        conferencia_rbx_zerada = {**conferencia_rbx, "liquido_transitoria_simulado": 0.0,
                                   "qtd_transitoria_simulado": 1, "zerada": True,
                                   "pronta_para_transferencia": False}
        caminho_zerada = gerar_dashboard(
            resultado, fila, rebate_cmm, canal="FULL", data="01.01",
            data_ref_rb="01/01/2026", caminho_xlsx_completo="a.xlsx",
            caminho_xlsx_rebate="b.xlsx", pasta_saida=tmp, alvo=-731.84,
            conferencia_rbx=conferencia_rbx_zerada,
        )
        conteudo_zerada = open(caminho_zerada, encoding="utf-8").read()
        assert "zerada — transferência já feita" in conteudo_zerada
        assert "acumulando" not in conteudo_zerada
        assert "pronta para transferência" not in conteudo_zerada

        # alvo=0.0, liquido_transitoria_simulado=-7.0 -> nem zerou nem bateu,
        # so acumulando ainda -- ambar (nao e erro no meio do dia). Faixa de
        # status: calculado -731.84 x alvo 0.0 -> "Diferença de R$ 731,84".
        conferencia_rbx_acumulando = {**conferencia_rbx, "zerada": False, "pronta_para_transferencia": False}
        caminho3 = gerar_dashboard(
            resultado, fila, rebate_cmm, canal="FULL", data="01.01",
            data_ref_rb="01/01/2026", caminho_xlsx_completo="a.xlsx",
            caminho_xlsx_rebate="b.xlsx", pasta_saida=tmp, alvo=0.0,
            ml_atualizado=True, caminho_bat_scraper="rodar_scraper_full_0101.bat",
            caminho_bat_arquivar="arquivar_rbx_full_0101.bat",
            conferencia_rbx=conferencia_rbx_acumulando,
        )
        conteudo3 = open(caminho3, encoding="utf-8").read()
        assert "Diferença de R$ 731,84" in conteudo3 and 'class="card faixa bad"' in conteudo3
        assert f"acumulando (faltam {_moeda(7.0)})" in conteudo3 and 'class="pill warn"' in conteudo3
        assert "Valores do Mercado Livre atualizados" in conteudo3
        # aviso "Scraper pronto" removido (18/09/2026) -- caminho_bat_scraper
        # continua aceito (interativo.py ainda gera o .bat normalmente), so
        # nao tem mais texto solto apontando pra ele no dashboard.
        assert "rodar_scraper_full_0101.bat" not in conteudo3
        assert "Scraper pronto" not in conteudo3
        # aviso de arquivar (verde, destacado) continua
        assert ('Transferência pronta! Arquivar RBX: arquivar_rbx_full_0101.bat '
                '(uma pasta acima — dê 2 cliques nele)') in conteudo3
        # alerta vermelho de VERIFICAR removido (18/09/2026) -- so a propria
        # fila VERIFICAR mostra isso agora, sem banner duplicado.
        assert "pedido(s) em VERIFICAR" not in conteudo3

        # Cancelamento na Original (cenario real do FULL 09/09/2026, ver
        # conferir_rbx.py): Transitoria bruta (218490.47) excede o alvo
        # (217603.75) exatamente pelo total cancelado (-886.72) -- o painel
        # tem que listar os cancelamentos, mostrar o valor de transferencia
        # AJUSTADO (217603.75) e usar ele (nao o bruto) no "pronto pra
        # transferir". Pedido do Everton, 18/09/2026.
        conferencia_rbx_cancelado = {
            **conferencia_rbx,
            "liquido_transitoria_simulado": 218490.47, "qtd_transitoria_simulado": 1,
            "zerada": False, "pronta_para_transferencia": True,
            "liquido_cancelado_original": -886.72, "qtd_cancelado_original": 2,
            "cancelados_original": [
                ("2000018007305094", -512.27), ("2000017940791734", -374.45),
            ],
            "valor_transferencia": 217603.75,
        }
        caminho_cancelado = gerar_dashboard(
            resultado, fila, rebate_cmm, canal="FULL", data="01.01",
            data_ref_rb="01/01/2026", caminho_xlsx_completo="a.xlsx",
            caminho_xlsx_rebate="b.xlsx", pasta_saida=tmp, alvo=217603.75,
            conferencia_rbx=conferencia_rbx_cancelado,
        )
        conteudo_cancelado = open(caminho_cancelado, encoding="utf-8").read()
        assert "Cancelado hoje" in conteudo_cancelado
        assert _moeda(-886.72) in conteudo_cancelado
        assert '<span>2000018007305094</span><span class="v neg">−512,27</span>' in conteudo_cancelado
        assert '<span>2000017940791734</span><span class="v neg">−374,45</span>' in conteudo_cancelado
        assert "Valor de transferência (ajustado)" in conteudo_cancelado
        assert _moeda(217603.75) in conteudo_cancelado
        assert "pronta para transferência" in conteudo_cancelado

        # ES depois do FULL, mesma data -> MESMO arquivo, sobrescreve (so o
        # canal atual fica na tela -- pedido do Everton, sem guias)
        caminho4 = gerar_dashboard(
            resultado, fila, rebate_cmm, canal="ES", data="01.01",
            data_ref_rb="01/01/2026", caminho_xlsx_completo="c.xlsx",
            caminho_xlsx_rebate="d.xlsx", pasta_saida=tmp, alvo=None,
        )
        assert caminho4 == caminho  # continua sendo o mesmo 1 link
        conteudo4 = open(caminho4, encoding="utf-8").read()
        assert "Conciliação ES — 01.01" in conteudo4
        assert "Conciliação FULL" not in conteudo4  # FULL nao fica mais na tela

        # Dia fechado: fila vazia -> estado vazio amigavel + faixa verde.
        caminho_fechado = gerar_dashboard(
            resultado[resultado["STATUS"] != "VERIFICAR"], fila.iloc[0:0], rebate_cmm, canal="FULL", data="03.01",
            data_ref_rb="01/01/2026", caminho_xlsx_completo="a.xlsx",
            caminho_xlsx_rebate="b.xlsx", pasta_saida=tmp, alvo=190.0,
        )
        conteudo_fechado = open(caminho_fechado, encoding="utf-8").read()
        assert "Nenhum pedido em VERIFICAR 🎉" in conteudo_fechado
        assert "Dia fechado ✓" in conteudo_fechado and 'class="card faixa ok"' in conteudo_fechado
        assert 'id="tabela-verificar"' not in conteudo_fechado

        # Marcacao extra na fila VERIFICAR (Mediacao/Cashback/etc, ver
        # NATUREZAS_EXTRAS) -- fica de fora por padrao (dict vazio, sem
        # codigo de NATUREZA chutado), so testa aqui com codigos FICTICIOS
        # (9999/8888, nunca usados de verdade), restaurando o dict original
        # (vazio) no final pra nao vazar pro dashboard real.
        original_naturezas_extras = NATUREZAS_EXTRAS.copy()
        NATUREZAS_EXTRAS["Mediacao (teste)"] = (9999, "MOTIVO MEDIACAO TESTE")
        NATUREZAS_EXTRAS["Cashback (teste)"] = (8888, "MOTIVO CASHBACK TESTE")
        try:
            caminho_extra = gerar_dashboard(
                resultado, fila, rebate_cmm, canal="FULL", data="02.01",
                data_ref_rb="01/01/2026", caminho_xlsx_completo="a.xlsx",
                caminho_xlsx_rebate="b.xlsx", pasta_saida=tmp, alvo=None,
                conferencia_rbx=conferencia_rbx, responsavel="EVERTON",
            )
            conteudo_extra = open(caminho_extra, encoding="utf-8").read()
            assert "<th>Marcação</th>" in conteudo_extra
            assert 'class="sel-natureza-extra"' in conteudo_extra
            assert 'data-pedido="3"' in conteudo_extra
            assert 'data-valor="-921,84"' in conteudo_extra
            assert ('<option value="Mediacao (teste)" data-natureza="9999" '
                    'data-motivo="MOTIVO MEDIACAO TESTE">Mediacao (teste)</option>') in conteudo_extra
            assert ('<option value="Cashback (teste)" data-natureza="8888" '
                    'data-motivo="MOTIVO CASHBACK TESTE">Cashback (teste)</option>') in conteudo_extra
            # HISTORICO da marcacao extra e calculado no JS (so se sabe o
            # motivo depois da escolha) -- confere que PREFIXO_MOV foi
            # gerado certo pro FULL, igual ja confirmado pro Frete/Rebate.
            assert 'var PREFIXO_MOV = "MKP_ME_FULL";' in conteudo_extra
            linha3_extra = conteudo_extra.split('<tr data-pedido="3"')[1].split("</tr>")[0]
            assert linha3_extra.count("<td") == 14  # 13 de sempre (com AB total) + Marcacao
            assert 'sel-natureza-extra' in linha3_extra
            # JS: select entra na mesma lista que os checkboxes (mesmo
            # formato chk.dataset.*) e persiste em localStorage por pedido.
            assert 'document.querySelectorAll(".sel-natureza-extra")' in conteudo_extra
            assert 'var CHAVE_NATUREZA_EXTRA = CHAVE + "_natureza_extra";' in conteudo_extra
        finally:
            NATUREZAS_EXTRAS.clear()
            NATUREZAS_EXTRAS.update(original_naturezas_extras)

        # Sem nenhuma NATUREZAS_EXTRAS configurada (estado real/padrao), a
        # coluna Marcacao nao aparece -- nao faz sentido mostrar um
        # seletor sem nenhuma opcao dentro (o JS de wiring continua no HTML,
        # so nao acha nenhum elemento .sel-natureza-extra pra pegar).
        assert "<th>Marcação</th>" not in conteudo
        assert 'class="sel-natureza-extra"' not in conteudo

        print("dashboard.py: ok")

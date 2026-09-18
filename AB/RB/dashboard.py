"""Gera o dashboard HTML (autocontido, sem CDN/JS externo -- roda offline
numa pasta de rede) que substitui o resumo em texto do terminal: cards de
totais, rosca de STATUS, alvo do dia (calendario x RB+Rebate+Ajuste),
distancia do alvo (DIF) da fila VERIFICAR, a propria fila em tabela, e
botoes que linkam pros xlsx salvos do lado.

1 arquivo por dia (dashboard_<data>.html -- pedido do Everton, so quer 1
link): cada `gerar_dashboard()` SOBRESCREVE esse mesmo arquivo com o canal
que acabou de rodar (FULL, depois ES, etc. -- so o ultimo fica na tela; o
historico de cada canal continua nos xlsx salvos do lado). A pagina se
recarrega sozinha quando o Everton volta pra aba (visibilitychange), entao
ele so roda o processo de novo e alterna pra aba que ja esta aberta.

Paleta e specs de marca (donut, barra, cores de status) seguem o skill
dataviz do Claude Code (references/palette.md e marks-and-anatomy.md),
adaptadas pro tema dark.

ponytail: tema dark fixo, sem toggle -- so foi pedido dark, adicionar
toggle claro/escuro se o time pedir depois.
"""

import html
import json
import os

from canais import CANAIS
from rebate_cmm import COLUNAS as COLUNAS_CMM
from rebate_cmm import MOTIVO_REBATE, NATUREZA_REBATE

COR_OK = "#2fbf4f"
COR_ALERTA = "#e0524f"
COR_AVISO = "#e0b34f"
# So pra texto direto no fundo escuro da tabela (#3a3e46) -- COR_OK/COR_ALERTA
# foram pensadas pra bolinha/pastilha com fundo proprio, contraste baixo
# demais como cor de texto crua (principalmente o vermelho). Pedido do
# Everton, 17/09/2026.
COR_OK_TEXTO = "#7ee2a8"
COR_ALERTA_TEXTO = "#ff9b93"  # ja usado no aviso-critico, mesmo motivo

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
MOTIVO_COMISSAO = "TARIFA DE COBRANCA"
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
# (chk.dataset.*, ver _cel_natureza).
_CAMPOS_CMM_JS = {
    "RESPONSAVEL": "RESPONSAVEL", "DATA_LACTO": "DATA_LACTO", "ORIGEM": "ORIGEM",
    "PAGADOR": "PAGADOR", "LOJA": "LOJA",
    "MOTIVO DO MOVIMENTO": "chk.dataset.motivo", "NATUREZA": "chk.dataset.natureza",
    # apostrofo na frente, igual _pedido_html -- sem isso o pedido (ate 16
    # digitos) vira notacao cientifica ao colar no Excel.
    "PEDIDO MKP / CV": "\"'\" + chk.dataset.pedido", "VENDEDOR": "VENDEDOR",
    "BANCO": "chk.dataset.banco", "AGENCIA": "chk.dataset.agencia",
    "CONTA": "chk.dataset.conta", "HISTORICO": "chk.dataset.historico",
    "VALOR R$": "chk.dataset.valor", "OBSERVACAO": '""',
    "PEDIDO_ERP": '""', "PEDIDO_SITE": '""', "N. FISCAL": '""', "PROCESSO": '""',
}


def _moeda(valor):
    neg = valor < 0
    texto = f"{abs(valor):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"{'-' if neg else ''}R$ {texto}"


def _formatar_numero(valor):
    neg = valor < 0
    numero = f"{abs(valor):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"{'-' if neg else ''}{numero}"


def _valor_neutro_html(valor):
    """Pastilha cinza neutra (mesmo padrao de _valor_copiavel_html/
    _pedido_html -- preto no cinza, sempre legivel na tela e colado no
    Excel, nao herda o --ink claro do tema dark) -- usada nas colunas
    DIF/Comissao/Desc-Bonus/Frete/Repasse/Rebate/Diferenca extrato. Pedido
    do Everton, 17/09/2026: cor (vermelho/verde) so chama atencao de
    verdade nas colunas RBX/RB e no alerta de VERIFICAR; nas demais
    competia por atencao sem necessidade."""
    return _valor_copiavel_html(_formatar_numero(valor))


def _valor_dif_html(valor):
    """Como _valor_neutro_html (pastilha cinza neutra), mas com o SINAL de
    negativo em vermelho -- so o "-", o resto do numero continua neutro.
    Diferente de _valor_colorido_html (numero inteiro colorido, sem
    pastilha): aqui o Everton so quer um alerta rapido de debito/credito
    olhando a coluna DIF, sem competir com o resto da paleta nem perder a
    pastilha copiavel pro Excel. Pedido do Everton, 18/09/2026."""
    neg = valor < 0
    texto_abs = _formatar_numero(abs(valor))
    sinal = f'<span style="color:{COR_ALERTA}">-</span>' if neg else ""
    return _valor_copiavel_html(f"{sinal}{texto_abs}")


def _valor_colorido_html(valor):
    """Texto colorido (vermelho/negativo, verde/positivo), sem pastilha de
    fundo -- unico indicador nas colunas RBX/RB (achou/nao achou ja fica
    claro pela celula em branco; a cor so reforca debito x credito).
    Pedido do Everton, 17/09/2026: a pastilha de fundo ali era redundante
    (competia com o proprio "achou/nao achou" e com o checkbox da linha)."""
    cor = COR_ALERTA_TEXTO if valor < 0 else COR_OK_TEXTO
    return f'<span style="color:{cor}">{_formatar_numero(valor)}</span>'


def _valor_ml(v, colorido=False):
    """Celula em branco quando nao tem valor (None/NaN) -- sem "—" mudo
    (a borda leve da coluna ja marca a celula, nao precisa de um traco
    dentro dela). colorido=True so pras colunas RBX/RB (ver
    _valor_colorido_html); as demais ficam neutras (_valor_neutro_html)."""
    if v is None or (isinstance(v, float) and v != v):  # None ou NaN
        return ""
    if isinstance(v, (int, float)):
        return _valor_colorido_html(v) if colorido else _valor_neutro_html(v)
    return html.escape(str(v))


def _valor_copiavel_html(texto):
    """Badge com cor fixa -- nao herda o --ink do tema dark, que fica
    texto claro em fundo branco (invisivel) ao colar no Excel."""
    estilo = "color:#111;background:#e8e8e8;padding:2px 6px;border-radius:4px"
    return f'<span style="{estilo}">{texto}</span>'


def _pedido_html(pedido):
    """Como _valor_copiavel_html, mais um apostrofo na frente -- forca
    texto no Excel ao colar (mso-number-format sozinho nao e respeitado em
    todo paste; sem isso, o numero de ate 16 digitos vira notacao
    cientifica: 2,00002E+15). O Excel descarta o apostrofo sozinho ao
    colar, so fica o numero."""
    texto = html.escape(str(pedido))
    return _valor_copiavel_html(f"'{texto}")


def _painel_historico(raiz):
    """RAIZ + codigos _REPASSE/_COMISSAO/_FRETE do historico TOTVS pra
    colar direto ao resolver um VERIFICAR. Sufixos seguem sempre o padrao
    RAIZ + "_REPASSE" etc (confirmado nas 93 linhas da tabela HISTORICOS
    do CONTROLE_V2.xlsx -- ver extrair_historicos.py/historicos_totvs.json)
    -- deriva direto, nao depende dessa raiz estar na tabela (ex. ELETRO
    usa uma raiz propria desse pipeline, fora do CONTROLE_V2)."""
    linhas = [("RAIZ", raiz), ("Repasse", f"{raiz}_REPASSE"),
              ("Comissao", f"{raiz}_COMISSAO"), ("Frete", f"{raiz}_FRETE")]
    return "".join(
        f'<div class="alvo-linha"><span>{label}</span>{_valor_copiavel_html(html.escape(valor))}</div>'
        for label, valor in linhas
    )


def _cel_detalhe(valor_html, texto_detalhe):
    """<td> com tooltip nativo (title) listando as sublinhas que compoem
    esse total, com o mesmo rotulo que o Mercado usou -- pedido do Everton,
    16/09/2026 ("regua de calculo" ao passar o mouse). Sem JS: o navegador
    ja mostra o title no hover."""
    if not texto_detalhe:
        return f'<td class="num">{valor_html}</td>'
    return f'<td class="num tem-detalhe" title="{html.escape(texto_detalhe)}">{valor_html}</td>'


def _conteudo_natureza(valor, pedido, motivo, natureza, conta, historico):
    """Checkbox + valor de 1 marcacao de NATUREZA, sem o <td> em volta --
    normalmente 1 por celula (ver _cel_natureza), mas a coluna Frete pode
    precisar de 2 na MESMA celula quando o pedido tem custo de envio e
    tarifa de cobranca separados (ver _linhas_tabela). Acrescenta o
    checkbox so quando o valor existe e e != 0. Pedido do Everton,
    17/09/2026."""
    valor_html = _valor_ml(valor)
    if not valor_html or not valor:  # "" (None/NaN) e 0 nao geram checkbox
        return valor_html
    checkbox = (
        '<input type="checkbox" class="chk-natureza" '
        f'data-pedido="{pedido}" data-motivo="{html.escape(motivo)}" '
        f'data-natureza="{natureza}" data-conta="{conta}" '
        f'data-historico="{html.escape(historico)}" '
        f'data-banco="{BANCO_PADRAO}" data-agencia="{AGENCIA_PADRAO}" '
        f'data-valor="{_formatar_numero(valor)}"> '
    )
    return f"{checkbox}{valor_html}"


def _cel_natureza(valor, texto_detalhe, pedido, motivo, natureza, conta, historico):
    """Como _cel_detalhe, mas pras colunas que entram na marcacao de
    NATUREZA (Frete/Rebate, ver constantes no topo do arquivo) -- 1 unica
    marcacao por celula (ver _conteudo_natureza)."""
    conteudo = _conteudo_natureza(valor, pedido, motivo, natureza, conta, historico)
    return _cel_detalhe(conteudo, texto_detalhe)


def _cel_marcacao_extra(dif, pedido, conta, banco, agencia):
    """<select> pra marcar o DIF INTEIRO da linha VERIFICAR com uma
    natureza extra (Mediacao/Cashback/etc, ver NATUREZAS_EXTRAS) -- ao
    contrario de _conteudo_natureza (checkbox por sub-valor calculado,
    pode ter mais de 1 marcado por pedido), aqui e no maximo 1 escolha por
    pedido: o valor marcado e sempre o DIF inteiro da linha, entao so faz
    sentido ele ser UMA coisa (Mediacao OU Cashback OU nenhuma). Celula em
    branco quando DIF e None/0 (nada pra marcar). So chamada quando
    NATUREZAS_EXTRAS ja tem pelo menos 1 opcao (ver _linhas_tabela)."""
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
    """Valor liquido (debito vermelho, credito verde, texto colorido sem
    pastilha -- ver _valor_colorido_html) lancado nesse pedido numa conta
    (RB ou RBX) -- celula em branco quando nao achou nada la (ou quando
    conferencia_rbx is None, sem rb.xlsx/rbx.xlsx pra conferir esse dia).
    Pedido do Everton, 15-17/09/2026: substitui a coluna Acao sugerida
    (crua demais) por um check/negativo, depois pelo valor real lancado."""
    return f'<td class="num">{_valor_ml((valores or {}).get(pedido), colorido=True)}</td>'


def _semaforo(v_rb, v_rbx, diferenca_extrato):
    """Ponto colorido (mesmo padrao .ponto ja usado no painel Alvo/Original
    x Transitoria): amarelo quando o pedido nao foi encontrado nem no RB
    nem no RBX (nada pra comparar ainda); verde quando achou e a Diferenca
    extrato fecha (dentro de TOLERANCIA_ALVO); vermelho quando achou mas
    sobra diferenca. Pedido do Everton, 16/09/2026."""
    if v_rb is None and v_rbx is None:
        cor = COR_AVISO
    elif abs(diferenca_extrato) < TOLERANCIA_ALVO:
        cor = COR_OK
    else:
        cor = COR_ALERTA
    return f'<span class="ponto" style="background:{cor}"></span>'


def _linhas_tabela(fila_verificar, conferencia_rbx=None, historicos=None):
    valores_rb = conferencia_rbx["valores_rb"] if conferencia_rbx else None
    valores_rbx = conferencia_rbx["valores_rbx"] if conferencia_rbx else None
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
        # separadas, a celula Frete ganha 2 checkboxes (custo de envio +
        # tarifa); quando so tem 1 (ou nenhuma), continua com 1 so, igual
        # antes.
        cel_comissao = _cel_detalhe(_valor_ml(comissao), row.get('ML_COMISSAO_DETALHE'))
        cel_desconto = _cel_detalhe(_valor_ml(row.get('ML_DESCONTO_BONUS')), row.get('ML_DESCONTO_BONUS_DETALHE'))
        tarifa_cobranca = None
        if frete is not None and custo_envio is not None and round(frete - custo_envio, 2) != 0:
            tarifa_cobranca = round(frete - custo_envio, 2)
        if tarifa_cobranca is not None:
            # 2 marcacoes na mesma celula, uma embaixo da outra (nao lado a
            # lado) -- <div> e block-level por padrao, cada uma cai numa
            # linha propria dentro do <td> sem precisar de CSS novo. Pedido
            # do Everton, 17/09/2026.
            conteudo_envio = (
                "<div>" + _conteudo_natureza(custo_envio, pedido, MOTIVO_FRETE, NATUREZA_FRETE,
                                              CONTA_TRANSITORIA, historicos.get("frete", "")) + "</div>"
                "<div>" + _conteudo_natureza(tarifa_cobranca, pedido, MOTIVO_COMISSAO, NATUREZA_COMISSAO,
                                              CONTA_TRANSITORIA, historicos.get("comissao", "")) + "</div>"
            )
            cel_envio = _cel_detalhe(conteudo_envio, row.get('ML_ENVIO_DETALHE'))
        else:
            cel_envio = _cel_natureza(frete, row.get('ML_ENVIO_DETALHE'), pedido,
                                       MOTIVO_FRETE, NATUREZA_FRETE, CONTA_TRANSITORIA,
                                       historicos.get("frete", ""))
        cel_rebate = _cel_natureza(row.get('ML_REBATE'), row.get('ML_REBATE_DETALHE'), pedido,
                                    MOTIVO_REBATE, NATUREZA_REBATE, CONTA_TRANSITORIA,
                                    historicos.get("rebate", ""))

        repasse = _repasse_calculado(dif, comissao, custo_envio)
        detalhe_repasse = None
        if repasse is not None:
            detalhe_repasse = "\n".join([
                f"DIF (banco, absoluto): {abs(dif):.2f}".replace(".", ","),
                f"Comissao (absoluto): {abs(comissao):.2f}".replace(".", ","),
                f"Custo de envio (absoluto, sem tarifa de cobranca): {abs(custo_envio):.2f}".replace(".", ","),
                f"= Repasse: {repasse:.2f}".replace(".", ","),
            ])
        cel_repasse = _cel_detalhe(_valor_ml(repasse), detalhe_repasse)
        pedido_mkp = str(row["PEDIDO_MKP"]).strip()
        cel_rbx = _cel_valor_conta(pedido_mkp, valores_rbx)
        cel_rb = _cel_valor_conta(pedido_mkp, valores_rb)

        # Diferenca extrato = DIF - valor RB - valor RBX (soma os 2 quando o
        # pedido aparece lancado nas duas contas ao mesmo tempo) -- None
        # quando nao achou em nenhuma das duas, pra nao sugerir uma
        # diferenca que na verdade e so a DIF original (o semaforo fica
        # amarelo nesse caso, ver _semaforo). Nao muda o STATUS (continua
        # VERIFICAR), so informa o analista. Pedido do Everton, 16/09/2026.
        v_rb = (valores_rb or {}).get(pedido_mkp)
        v_rbx = (valores_rbx or {}).get(pedido_mkp)
        diferenca_extrato = (
            round(dif - (v_rb or 0) - (v_rbx or 0), 2)
            if v_rb is not None or v_rbx is not None else None
        )
        semaforo = _semaforo(v_rb, v_rbx, diferenca_extrato)
        cel_diferenca = f'<td class="num">{semaforo} {_valor_ml(diferenca_extrato)}</td>'
        # DESCRICAO (coluna E do AB, ver cruzamento.py) -- pro Everton ver
        # "mediation"/"refund"/etc na fila sem abrir o AB pra procurar o
        # pedido. Vazio quando o pedido nao tem nenhuma linha no AB do dia
        # (so no RB/RBX). Pedido do Everton, 18/09/2026.
        cel_descricao = f'<td>{html.escape(str(row.get("DESCRICAO") or ""))}</td>'
        cel_marcacao = (
            _cel_marcacao_extra(dif, pedido, CONTA_TRANSITORIA, BANCO_PADRAO, AGENCIA_PADRAO)
            if NATUREZAS_EXTRAS else ""
        )

        saida.append(f"""
            <tr data-pedido="{pedido}">
              <td class="chk-col"><input type="checkbox" class="chk-ajustado" data-pedido="{pedido}"></td>
              {cel_rbx}
              {cel_rb}
              {cel_diferenca}
              <td style="mso-number-format:'\\@'">{_pedido_html(row['PEDIDO_MKP'])}</td>
              {cel_descricao}
              <td class="num">{_valor_dif_html(dif)}</td>
              {cel_marcacao}
              {cel_comissao}
              {cel_desconto}
              {cel_envio}
              {cel_repasse}
              {cel_rebate}
            </tr>""")
    return "".join(saida)


def _painel_alvo(alvo, rb_total, rebate_total, ajuste_pendente):
    calculado = rb_total + rebate_total + ajuste_pendente

    linhas_html = "".join(
        f'<div class="alvo-linha"><span>{label}</span><span class="num">{_moeda(v)}</span></div>'
        for label, v in [
            ("RB atual", rb_total),
            ("+ Rebate (cupom)", rebate_total),
            ("+ Ajuste pendente (fila VERIFICAR)", ajuste_pendente),
            ("= Calculado", calculado),
        ]
    )

    if alvo is None:
        status_html = '<p class="ink-muted">Alvo nao cadastrado no calendario pra esse canal/dia.</p>'
    else:
        diferenca = alvo - calculado
        bate = abs(diferenca) < TOLERANCIA_ALVO
        cor = COR_OK if bate else COR_ALERTA
        rotulo = "bate com o calendario" if bate else "NAO bate com o calendario"
        status_html = (
            f'<div class="alvo-linha alvo-destaque"><span>Alvo (calendario)</span>'
            f'<span class="num">{_moeda(alvo)}</span></div>'
            f'<div class="alvo-status" style="color:{cor}">'
            f'<span class="ponto" style="background:{cor}"></span>{rotulo}'
            f' ({_moeda(diferenca)})</div>'
        )

    return linhas_html + status_html


def _painel_conferencia_rbx(conf, alvo):
    """Conta ORIGINAL (103925) x conta TRANSITORIA (X103925) -- ver
    conferir_rbx.py. So chamado quando conf nao e None (arquivos
    rb.xlsx/rbx.xlsx existem em AUXILIAR/).

    Correcao de regra, 17/09/2026: a Transitoria NAO zera durante o dia --
    ela acumula o movimento. Quando o liquido dela (chave do canal + RA)
    bate com o MESMO alvo usado na linha Original, e sinal de que ja pode
    fazer a transferencia manual pra Original (que ai sim zera a
    Transitoria). Por isso a linha da Transitoria nao usa o _linha()
    generico de bate/nao-bate (que compararia contra um "esperado" fixo).

    2o estado do ciclo, 17/09/2026: a Transitoria tem 2 momentos CERTOS --
    bate com o alvo (pronta pra transferir) OU voltou a zero (transferencia
    ja feita). So no meio do dia, acumulando sem bater com nenhum dos dois,
    e que e "normal, nao e erro" -- por isso essa 3a situacao usa cor
    neutra, nao vermelha."""
    def _linha(rotulo, valor, qtd, bate, msg_zero, verbo_ok, verbo_nao):
        if qtd == 0:
            return (
                f'<div class="alvo-linha"><span>{rotulo}</span>'
                f'<span class="ink-muted">{msg_zero}</span></div>'
            )
        cor = COR_OK if bate else COR_ALERTA
        rotulo_status = verbo_ok if bate else verbo_nao
        return (
            f'<div class="alvo-linha alvo-destaque"><span>{rotulo}</span>'
            f'<span class="num">{_moeda(valor)}</span></div>'
            f'<div class="alvo-status" style="color:{cor}">'
            f'<span class="ponto" style="background:{cor}"></span>{rotulo_status} '
            f'({qtd} lancamento(s))</div>'
        )

    rotulo_transitoria = "Transitoria (alvo do calendario)"
    qtd_transitoria_total = conf["qtd_transitoria_simulado"]
    if qtd_transitoria_total == 0:
        transitoria_html = (
            f'<div class="alvo-linha"><span>{rotulo_transitoria}</span>'
            f'<span class="ink-muted">sem lancamentos encontrados</span></div>'
        )
    else:
        valor_linha = (
            f'<div class="alvo-linha alvo-destaque"><span>{rotulo_transitoria}</span>'
            f'<span class="num">{_moeda(conf["liquido_transitoria_simulado"])}</span></div>'
        )
        if conf["zerada"]:
            status_linha = (
                f'<div class="alvo-status" style="color:{COR_OK}">'
                f'<span class="ponto" style="background:{COR_OK}"></span>'
                f'zerada (transferencia ja feita)</div>'
            )
        elif conf["pronta_para_transferencia"]:
            status_linha = (
                f'<div class="alvo-status" style="color:{COR_OK}">'
                f'<span class="ponto" style="background:{COR_OK}"></span>'
                f'pronto para fazer a transferencia</div>'
            )
        elif alvo is not None:
            faltam = _moeda(alvo - conf["liquido_transitoria_simulado"])
            status_linha = f'<div class="alvo-status ink-muted">acumulando (faltam {faltam} pro alvo)</div>'
        else:
            status_linha = '<div class="alvo-status ink-muted">sem alvo do calendario pra comparar</div>'
        transitoria_html = valor_linha + status_linha

    return (
        _linha("Original (calendario)", conf["liquido_original"], conf["qtd_original"],
               conf["original_bate"], "sem lancamento de fechamento ainda",
               "bate com o calendario", "NAO bate com o calendario")
        + transitoria_html
    )


def _tabela_lotes(lotes_status):
    """Guia 2: status dos lotes do pipeline AB bruto -> Para Processar (ver
    pipeline.py/conciliacao.db) -- vazio quando o banco ainda nao existe ou
    nao tem lote desse canal (feature opcional, mesmo padrao de
    conferencia_rbx=None). Pedido do Everton, 18/09/2026."""
    if not lotes_status:
        return '<p class="ink-muted">Nenhum lote registrado (pipeline.py sync).</p>'
    linhas = "".join(
        f"<tr><td>{html.escape(lote)}</td><td>{html.escape(status)}</td>"
        f"<td>{pedidos if pedidos is not None else '-'}</td>"
        f"<td>{acao_manual if acao_manual is not None else '-'}</td>"
        f"<td>{html.escape(rodado_em or '-')}</td></tr>"
        for lote, status, pedidos, acao_manual, rodado_em in lotes_status
    )
    return (
        "<table><thead><tr><th>Lote</th><th>Status</th><th>Pedidos</th>"
        f"<th>Acao manual</th><th>Rodado em</th></tr></thead><tbody>{linhas}</tbody></table>"
    )


# Colunas de valor de verdade da aba "Para Processar" (ver
# processar_full.COLUNAS_PARA_PROCESSAR) -- so essas usam _valor_ml
# (formatacao de moeda). As demais (SOURCE_ID/EXTERNAL_REFERENCE/etc) sao
# ID/texto: o pandas as carrega como float (NaN em alguma linha forca a
# coluna inteira pra float64) e _valor_ml formatava esses IDs como dinheiro
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
    if v is None or (isinstance(v, float) and v != v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return html.escape(str(int(v)))
    return html.escape(str(v))


def _tabela_para_processar(para_processar):
    """Guia 2: linhas cruas da aba "Para Processar" do AB do dia (ver
    processar_full.gerar_para_processar) -- colunas de valor (_AMOUNT)
    formatadas como moeda (_valor_ml), o resto como ID/texto puro
    (_valor_texto). Pedido do Everton, 18/09/2026."""
    if para_processar is None or not len(para_processar):
        return '<p class="ink-muted">Sem dados do AB (Para Processar) pra esse lote.</p>'
    colunas = list(para_processar.columns)
    cabecalho = "".join(f"<th>{html.escape(str(c))}</th>" for c in colunas)
    linhas = "".join(
        "<tr>" + "".join(
            f"<td>{_valor_ml(row[c]) if c in _COLUNAS_VALOR_PARA_PROCESSAR else _valor_texto(row[c])}</td>"
            for c in colunas
        ) + "</tr>"
        for _, row in para_processar.iterrows()
    )
    return f"<table><thead><tr>{cabecalho}</tr></thead><tbody>{linhas}</tbody></table>"


def _tabela_valores_conta(valores):
    """Guia 2: 1 linha por pedido com o valor liquido lancado numa conta
    (RB ou RBX) -- mesmos dados de _cel_valor_conta (valores_rb/valores_rbx
    de conferir_rbx.py), so que numa tabela dedicada em vez de coluna da
    fila VERIFICAR. Ordenado do mais negativo pro mais positivo (mesmo
    criterio de leitura em lote que _ordenar_fila_verificar usa). Pedido do
    Everton, 18/09/2026."""
    if not valores:
        return '<p class="ink-muted">Sem rb.xlsx/rbx.xlsx pra conferir ainda.</p>'
    linhas = "".join(
        f'<tr><td style="mso-number-format:\'\\@\'">{_pedido_html(pedido)}</td>'
        f'<td class="num">{_valor_colorido_html(valor)}</td></tr>'
        for pedido, valor in sorted(valores.items(), key=lambda kv: kv[1])
    )
    return f"<table><thead><tr><th>Pedido</th><th>Valor</th></tr></thead><tbody>{linhas}</tbody></table>"


def gerar_dashboard(resultado, fila_verificar, rebate_cmm, *, canal, data, data_ref_rb,
                     caminho_xlsx_completo, caminho_xlsx_rebate, pasta_saida, alvo=None,
                     ml_atualizado=False, caminho_bat_scraper=None, raiz=None,
                     caminho_xlsx_descricoes=None, conferencia_rbx=None, caminho_bat_atualizar=None,
                     caminho_bat_arquivar=None, ml_de_cache=False, responsavel=None,
                     rebate_processar=None, caminho_xlsx_rebate_processar=None,
                     para_processar=None, lotes_status=None):
    contagem = resultado["STATUS"].value_counts()
    total = len(resultado)

    # NATUREZA (Frete/Rebate marcados na fila VERIFICAR) -- HISTORICO/LOJA
    # derivados de CANAIS[canal] direto (nao do parametro `raiz`, que fica
    # None pros canais em CANAIS_SEM_RAIZ_CONFIRMADA -- essa marcacao nao
    # depende daquela confirmacao especifica). loja_calendario, nao
    # raiz.split("_")[-1] (dava "FU" pro FULL -- "FU" e o chave_rbx, usado
    # pra achar linha no extrato, nao o valor certo do campo LOJA do
    # CMM/MOV). Confirmado com o Everton, 18/09/2026.
    cfg_canal = CANAIS.get(canal, {})
    prefixo_mov = cfg_canal.get("prefixo_mov")
    loja = cfg_canal.get("loja_calendario", "")
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
    est_com_total = resultado.loc[resultado["STATUS"] == "EST COM", "DIF"].sum()
    mediacao_total = resultado.loc[resultado["STATUS"] == "MEDIACAO", "DIF"].sum()
    rb_total = resultado["TOTAL_RB"].sum()
    ajuste_pendente = fila_verificar["DIF"].sum()
    verificar_total = fila_verificar["DIF"].abs().sum()

    # Rebate (formula) x RBX: bate quando o que ja caiu direto do ML
    # (chave_cupom_rebate) + o que ja foi lancado a mao no CMM
    # (prefixo_mov + MOTIVO_REBATE) fecha com o total calculado -- ver
    # conferir_rbx.conferir. Sem rb.xlsx/rbx.xlsx (conferencia_rbx None),
    # sem bolinha -- nada pra conferir ainda. Pedido do Everton, 18/09/2026.
    rotulo_rebate = "Rebate (formula)"
    sub_rebate = f"{contagem.get('REBATE', 0)} pedido(s)"
    if conferencia_rbx:
        cor_rebate = COR_OK if conferencia_rbx["rebate_bate"] else COR_ALERTA
        rotulo_rebate += f' <span class="ponto" style="background:{cor_rebate}"></span>'
        sub_rebate += (
            f" | RBX: {_moeda(conferencia_rbx['liquido_rebate_rbx'])}"
            f" (cupom {conferencia_rbx['qtd_cupom_rebate']} + CMM {conferencia_rbx['qtd_cmm_rebate']})"
        )

    cards = [
        ("Total de pedidos", f"{total}", None),
        (rotulo_rebate, _moeda(rebate_total), sub_rebate),
        ("Est. comissao", _moeda(est_com_total), f"{contagem.get('EST COM', 0)} pedido(s)"),
        ("Mediacao", _moeda(mediacao_total), f"{contagem.get('MEDIACAO', 0)} pedido(s)"),
    ]
    cards_html = "".join(
        f'<div class="card"><div class="card-label">{l}</div><div class="card-valor">{v}</div>'
        + (f'<div class="card-sub">{s}</div>' if s else "") + "</div>"
        for l, v, s in cards
    )

    verificar_qtd = len(fila_verificar)
    alerta_html = ""
    if verificar_qtd:
        alerta_html = (
            f'<div class="alerta"><span class="icone">!</span> '
            f'<strong>{verificar_qtd} pedido(s) em VERIFICAR</strong> -- '
            f'R$ em jogo: {_moeda(verificar_total)}. Precisam de conferencia manual.</div>'
        )

    alvo_html = _painel_alvo(alvo, rb_total, rebate_total, ajuste_pendente)
    historico_html = _painel_historico(raiz) if raiz else ""
    conferencia_html = _painel_conferencia_rbx(conferencia_rbx, alvo) if conferencia_rbx else ""

    colspan_verificar = 13 if NATUREZAS_EXTRAS else 12
    tabela_html = (
        _linhas_tabela(fila_verificar, conferencia_rbx, historicos) if verificar_qtd
        else f'<tr><td colspan="{colspan_verificar}" class="ink-muted">Nenhum pedido em VERIFICAR.</td></tr>'
    )

    botoes = []
    botoes.append(f'<a class="botao" href="{html.escape(caminho_xlsx_completo)}">Baixar planilha completa</a>')
    if len(rebate_cmm):
        botoes.append(f'<a class="botao" href="{html.escape(caminho_xlsx_rebate)}">Baixar XLSX Rebate CMM ({len(rebate_cmm)})</a>')
    if caminho_xlsx_rebate_processar and rebate_processar is not None and len(rebate_processar):
        botoes.append(
            f'<a class="botao" href="{html.escape(caminho_xlsx_rebate_processar)}">'
            f'Baixar PROCESSAR_REBATE ({len(rebate_processar)})</a>'
        )
    if caminho_xlsx_descricoes:
        botoes.append(f'<a class="botao" href="{html.escape(caminho_xlsx_descricoes)}">Baixar descricoes ML/MP</a>')
    if caminho_bat_atualizar:
        botoes.append(
            f'<span class="botao" style="cursor:default">'
            f'Atualizar (RB/RBX/AB): {html.escape(caminho_bat_atualizar)} (uma pasta acima -- da 2 cliques nele)</span>'
        )
    botoes_html = "".join(botoes)

    valores_rbx = (conferencia_rbx or {}).get("valores_rbx") or {}
    valores_rb = (conferencia_rbx or {}).get("valores_rb") or {}
    guia2_html = f"""
      <div class="grid-main">
        <div class="col">
          <div class="painel">
            <h2>Lotes (pipeline AB)</h2>
            {_tabela_lotes(lotes_status)}
          </div>
          <div class="painel tabela-verificar">
            <h2>RBX ({len(valores_rbx)})</h2>
            <div class="tabela-scroll">{_tabela_valores_conta(valores_rbx)}</div>
          </div>
          <div class="painel tabela-verificar">
            <h2>RB ({len(valores_rb)})</h2>
            <div class="tabela-scroll">{_tabela_valores_conta(valores_rb)}</div>
          </div>
        </div>
        <div class="col">
          <div class="painel tabela-verificar">
            <h2>AB Para Processar ({len(para_processar) if para_processar is not None else 0})</h2>
            <div class="tabela-scroll">{_tabela_para_processar(para_processar)}</div>
          </div>
        </div>
      </div>"""

    titulo = f"Conciliacao {canal} - {data}"
    chave_local_storage = f"ajustados_{canal.lower()}_{data.replace('.', '')}"
    # Avisos discretos do topo (texto simples, sem barra de cor cheia) --
    # o alerta vermelho de VERIFICAR (alerta_html) e o unico que precisa
    # chamar atencao de verdade; esses aqui so informam de passagem. Pedido
    # do Everton, 17/09/2026: antes "Scraper pronto" era um botao azul
    # cheio e "Valores do ML atualizados" uma barra verde cheia, competindo
    # com o alerta pela atencao.
    avisos_topo = []
    if caminho_bat_arquivar:
        # Destacado (verde) mesmo no meio dos avisos discretos -- e o unico
        # que exige uma acao concreta do Everton (arquivar a copia do RBX
        # que motivou a transferencia). So aparece quando a Transitoria ja
        # bateu com o alvo (ver executar()/conferir_rbx.pronta_para_transferencia).
        # Pedido do Everton, 17/09/2026.
        avisos_topo.append(
            f'<div class="aviso-topo" style="color:{COR_OK_TEXTO}">'
            f'<span class="ponto" style="background:{COR_OK}"></span>'
            f'Transferencia pronta! Arquivar RBX: {html.escape(caminho_bat_arquivar)} '
            '(uma pasta acima -- da 2 cliques nele)</div>'
        )
    if ml_atualizado:
        avisos_topo.append(
            f'<div class="aviso-topo"><span class="ponto" style="background:{COR_OK}"></span>'
            f'Valores do Mercado Livre atualizados</div>'
        )
    elif ml_de_cache:
        avisos_topo.append(
            '<div class="aviso-topo"><span class="ponto" style="background:var(--ink-muted)"></span>'
            'Valores do Mercado Livre de uma consulta anterior (rodou sem scraper) -- '
            'pode ter pedido novo ainda sem consultar</div>'
        )
    if caminho_bat_scraper:
        # Sem link "download": o Chrome ignora esse atributo quando a
        # propria pagina e file:// (seguranca do navegador) -- em vez de
        # baixar, so navega pro .bat e mostra o texto dele na tela. So avisa
        # onde esta. Os .bat ficam na RAIZ (pasta_saida), uma pasta acima do
        # dashboard (Conciliacoes/) -- nao "nessa mesma pasta". Pedido do
        # Everton, 15/09/2026: juntar todos os .bat na raiz, so xlsx/html
        # dentro das subpastas.
        avisos_topo.append(
            f'<div class="aviso-topo">Scraper pronto: {html.escape(caminho_bat_scraper)} '
            '(uma pasta acima -- da 2 cliques nele)</div>'
        )
    avisos_topo_html = "".join(avisos_topo)

    corpo = f"""<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(titulo)}</title>
<style>
  :root {{
    --page: #14161a; --surface: #1c1f24; --surface-2: #22262c;
    --ink: #f2f1ee; --ink-secondary: #b8b6b0; --ink-muted: #83817c;
    --gridline: #33363c; --border: rgba(255,255,255,0.08);
  }}
  * {{ box-sizing: border-box; }}
  html {{ font-size: 17px; }}
  html, body {{ margin: 0; }}
  body {{
    padding: 20px; background: var(--page); color: var(--ink);
    font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
  }}
  /* Sem altura travada em 100vh/overflow:hidden -- isso forcava tudo a
     encolher pra caber numa tela so, ficando "espremido" em monitores/
     janelas mais baixas (o Everton tinha que dar zoom in/out toda hora
     pra compensar). Agora a pagina rola normal quando o conteudo nao
     cabe; so a tabela da fila VERIFICAR/RBX/RB mantem rolagem propria
     (ver .tabela-scroll), limitada a uma altura razoavel, pra nao virar
     uma lista infinita dominando a tela. Pedido do Everton, 18/09/2026. */
  .wrap {{ max-width: 100%; margin: 0 auto; display: flex; flex-direction: column; gap: 16px; }}
  .header {{ display: flex; align-items: baseline; justify-content: space-between; gap: 16px; flex-wrap: wrap; }}
  h1 {{ font-size: 1.5rem; margin: 0; }}
  .ref {{ color: var(--ink-secondary); margin: 0; font-size: 0.95rem; }}
  .ink-muted {{ color: var(--ink-muted); }}
  .abas {{ display: flex; gap: 6px; }}
  .aba-btn {{
    background: var(--surface-2); border: 1px solid var(--border); color: var(--ink-secondary);
    border-radius: 8px; padding: 7px 14px; font-size: 0.92rem; font-weight: 600; cursor: pointer;
  }}
  .aba-btn.ativo {{ color: var(--ink); border-color: var(--ink-secondary); }}
  .guia {{ display: flex; flex-direction: column; gap: 16px; }}
  .guia[hidden] {{ display: none; }}
  .aviso-topo {{
    color: var(--ink-secondary); font-size: 0.85rem;
    display: flex; align-items: center; gap: 6px; padding: 1px 2px;
  }}
  .aviso-topo .ponto {{ width: 7px; height: 7px; border-radius: 50%; display: inline-block; flex-shrink: 0; }}
  .alerta {{
    background: #2e1a1a; border: 1px solid #e0524f55; color: #ff9b93;
    border-radius: 8px; padding: 12px 18px; font-size: 1rem; font-weight: 600;
  }}
  .alerta .icone {{
    display: inline-flex; width: 20px; height: 20px; border-radius: 50%;
    background: #e0524f; color: #14161a; align-items: center; justify-content: center;
    font-weight: bold; margin-right: 6px; font-size: 0.85rem;
  }}
  .grid-main {{ display: grid; grid-template-columns: 340px 1fr; gap: 16px; align-items: start; }}
  @media (max-width: 1000px) {{ .grid-main {{ grid-template-columns: 1fr; }} }}
  .col {{ min-width: 0; display: flex; flex-direction: column; gap: 16px; }}
  .painel {{
    background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 16px;
    min-width: 0; display: flex; flex-direction: column;
  }}
  .painel h2 {{ font-size: 0.92rem; margin: 0 0 10px; color: var(--ink-secondary); text-transform: uppercase; letter-spacing: 0.03em; font-weight: 700; }}
  .cards {{ display: flex; flex-direction: row; flex-wrap: wrap; gap: 12px; }}
  .card {{
    background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 14px 18px;
    width: auto; min-width: 160px; min-height: 82px; flex: 0 0 auto; display: flex; flex-direction: column; justify-content: center;
  }}
  .card-label {{ font-size: 0.85rem; color: var(--ink-secondary); white-space: nowrap; }}
  .card-valor {{ font-size: 1.55rem; font-weight: 700; margin-top: 4px; white-space: nowrap; }}
  .card-sub {{ font-size: 0.78rem; color: var(--ink-muted); margin-top: 2px; white-space: nowrap; }}
  .ponto {{ width: 9px; height: 9px; border-radius: 50%; display: inline-block; flex-shrink: 0; }}
  .alvo-linha {{ display: flex; justify-content: space-between; gap: 10px; font-size: 0.9rem; padding: 5px 0; border-bottom: 1px dashed var(--gridline); }}
  .alvo-linha span:first-child {{ flex: 1 1 auto; min-width: 0; overflow-wrap: break-word; }}
  .alvo-linha .num {{ flex: 0 0 auto; white-space: nowrap; font-variant-numeric: tabular-nums; font-weight: 600; }}
  .alvo-destaque {{ border-bottom: none; padding-top: 8px; font-weight: 700; font-size: 1rem; }}
  .alvo-status {{ display: flex; align-items: center; gap: 6px; font-size: 0.88rem; margin-top: 8px; }}
  table {{ width: max-content; border-collapse: collapse; font-size: 0.92rem; background: #3a3e46; }}
  tr {{ height: 48px; }}
  th, td {{
    text-align: right; padding: 8px 12px; white-space: nowrap; font-weight: 500;
    border-bottom: 1px solid var(--gridline); border-right: 1px solid var(--gridline);
  }}
  th {{ position: sticky; top: 0; background: var(--surface); color: var(--ink-secondary); font-weight: 700; text-align: center; }}
  td.num {{ font-variant-numeric: tabular-nums; }}
  td.tem-detalhe {{ cursor: help; }}
  td.chk-col, th.chk-col {{ width: 30px; text-align: center; padding: 8px 4px; }}
  .chk-ajustado {{ width: 17px; height: 17px; accent-color: {COR_OK}; cursor: pointer; }}
  tr.linha-feita {{ opacity: 0.45; }}
  .tabela-scroll {{ overflow: auto; max-height: 65vh; }}
  .botoes {{ display: flex; gap: 10px; margin-top: 4px; flex-wrap: wrap; }}
  .botao {{
    display: inline-block; padding: 9px 16px; border-radius: 8px; text-decoration: none;
    background: var(--surface-2); border: 1px solid var(--border); color: var(--ink); font-size: 0.92rem; font-weight: 600;
  }}
  .chk-natureza {{ accent-color: {COR_OK}; cursor: pointer; margin-right: 3px; vertical-align: middle; }}
  .chk-natureza:checked + span {{ opacity: 0.45; }}
  .sel-natureza-extra {{
    background: var(--surface-2); color: var(--ink); border: 1px solid var(--border);
    border-radius: 4px; font-size: 0.82rem; padding: 3px 5px; max-width: 130px;
  }}
  #painel-natureza {{ display: none; flex-direction: row; gap: 16px; max-height: 220px; }}
  .natureza-col {{ flex: 1 1 0; min-width: 0; display: flex; flex-direction: column; gap: 6px; }}
  .natureza-col textarea {{
    flex: 1 1 auto; min-height: 80px; resize: vertical; font-family: "Consolas", monospace;
    font-size: 0.8rem; background: var(--surface-2); color: var(--ink);
    border: 1px solid var(--border); border-radius: 6px; padding: 7px 9px;
  }}
</style>
</head>
<body>
<div class="wrap">
  <div class="header">
    <div>
      <h1>{html.escape(titulo)}</h1>
      <p class="ref">RB ref: {html.escape(data_ref_rb or '?')}</p>
    </div>
    <div class="botoes">{botoes_html}</div>
  </div>

  <div class="abas">
    <button class="aba-btn ativo" data-aba="guia-1">Conciliacao</button>
    <button class="aba-btn" data-aba="guia-2">AB / RBX / RB</button>
  </div>

  <div class="guia" id="guia-1">
  {alerta_html}
  {avisos_topo_html}

  <div class="cards">{cards_html}</div>

  <div class="grid-main">
    <div class="col">
      <div class="painel">
        <h2>Alvo do dia</h2>
        {alvo_html}
      </div>
      {f'<div class="painel"><h2>Original x Transitoria</h2>{conferencia_html}</div>' if conferencia_html else ""}
      {f'<div class="painel"><h2>Historico TOTVS ({html.escape(canal)})</h2>{historico_html}</div>' if historico_html else ""}
    </div>

    <div class="col">
      <div class="painel tabela-verificar">
        <h2>Fila VERIFICAR</h2>
        <div class="tabela-scroll">
          <table>
            <thead><tr>
              <th class="chk-col"></th><th>RBX</th><th>RB</th><th>Diferença extrato</th>
              <th>Pedido</th><th>Descrição</th><th>DIF</th>
              {'<th>Marcação</th>' if NATUREZAS_EXTRAS else ''}
              <th>Comissao</th><th>Desc/Bonus</th><th>Frete</th><th>Repasse</th><th>Rebate</th>
            </tr></thead>
            <tbody>{tabela_html}</tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

  <div class="painel" id="painel-natureza">
    <div class="natureza-col">
      <h2>Colar no CMM (Movimento_Bancario)</h2>
      <textarea id="ta-natureza-cmm" readonly></textarea>
      <span class="botao" style="cursor:pointer" onclick="copiarNatureza('ta-natureza-cmm', this)">Copiar CMM</span>
    </div>
    <div class="natureza-col">
      <h2>Colar no MOV</h2>
      <textarea id="ta-natureza-mov" readonly></textarea>
      <span class="botao" style="cursor:pointer" onclick="copiarNatureza('ta-natureza-mov', this)">Copiar MOV</span>
    </div>
  </div>
  </div>

  <div class="guia" id="guia-2" hidden>{guia2_html}</div>
</div>
<script>
(function () {{
  var CHAVE = "{chave_local_storage}";
  var marcados;
  try {{ marcados = new Set(JSON.parse(localStorage.getItem(CHAVE) || "[]")); }}
  catch (e) {{ marcados = new Set(); }}

  function salvar() {{
    try {{ localStorage.setItem(CHAVE, JSON.stringify(Array.from(marcados))); }} catch (e) {{}}
  }}

  document.querySelectorAll(".chk-ajustado").forEach(function (chk) {{
    var linha = chk.closest("tr");
    if (marcados.has(chk.dataset.pedido)) {{
      chk.checked = true;
      linha.classList.add("linha-feita");
    }}
    chk.addEventListener("change", function () {{
      if (chk.checked) {{ marcados.add(chk.dataset.pedido); linha.classList.add("linha-feita"); }}
      else {{ marcados.delete(chk.dataset.pedido); linha.classList.remove("linha-feita"); }}
      salvar();
    }});
  }});

  // Abas (Conciliacao / AB-RBX-RB) -- guia ativa persiste em localStorage
  // (chave separada de CHAVE). O reload do visibilitychange acontece toda
  // vez que o Everton volta pra aba (nao so entre dias, ver mais abaixo) --
  // sem persistir, ele perdia a guia 2 toda vez que saia pra colar algo no
  // TOTVS e voltava. Pedido do Everton, 18/09/2026.
  var CHAVE_ABA = CHAVE + "_aba";

  function ativarAba(nomeAba) {{
    document.querySelectorAll(".aba-btn").forEach(function (b) {{
      b.classList.toggle("ativo", b.dataset.aba === nomeAba);
    }});
    document.querySelectorAll(".guia").forEach(function (g) {{
      g.hidden = g.id !== nomeAba;
    }});
  }}

  document.querySelectorAll(".aba-btn").forEach(function (botao) {{
    botao.addEventListener("click", function () {{
      ativarAba(botao.dataset.aba);
      try {{ localStorage.setItem(CHAVE_ABA, botao.dataset.aba); }} catch (e) {{}}
    }});
  }});

  try {{
    var abaSalva = localStorage.getItem(CHAVE_ABA);
    if (abaSalva && document.getElementById(abaSalva)) {{ ativarAba(abaSalva); }}
  }} catch (e) {{}}

  // Volta pra essa aba do navegador (depois de rodar o processo de novo,
  // outro canal ou nao) -> recarrega sozinho, sem precisar abrir outro link.
  document.addEventListener("visibilitychange", function () {{
    if (document.visibilityState === "visible") {{ location.reload(); }}
  }});

  // Marcacao de NATUREZA (Frete/Rebate) -- persiste em localStorage tambem
  // (chave separada de CHAVE), porque o visibilitychange reload (acima)
  // acontece a cada troca de aba, nao so entre sessoes -- sem isso o
  // Everton perdia a marcacao so de alternar pro TOTVS colar e voltar.
  // Chave e pedido+motivo (nao so pedido -- o Frete pode ter 2 checkboxes
  // no mesmo pedido, custo de envio + tarifa, ver _cel_natureza).
  // Pedido do Everton, 17/09/2026.
  var RESPONSAVEL = {json.dumps(responsavel or "")};
  var DATA_LACTO = {json.dumps(data_ref_rb or "")};
  var ORIGEM = "MKP";
  var PAGADOR = "MELI";
  var LOJA = {json.dumps(loja)};
  // "751" confirmado pro FULL (18/09/2026) -- so o FULL existe
  // configurado hoje, ver rebate_cmm.py/rebate_fora_da_regra.py.
  var VENDEDOR = "751";
  // Usado pra montar o HISTORICO das marcacoes extras (Mediacao/Cashback/
  // etc, ver NATUREZAS_EXTRAS): prefixo_mov + "_" + motivo da opcao
  // escolhida, mesmo padrao do Frete/Comissao/Rebate (calculado em Python
  // la, calculado aqui em JS porque o motivo so se sabe depois da escolha).
  var PREFIXO_MOV = {json.dumps(prefixo_mov or "")};

  var CHAVE_NATUREZA = CHAVE + "_natureza";
  var marcadosNatureza;
  try {{ marcadosNatureza = new Set(JSON.parse(localStorage.getItem(CHAVE_NATUREZA) || "[]")); }}
  catch (e) {{ marcadosNatureza = new Set(); }}

  function chaveNatureza(chk) {{ return chk.dataset.pedido + "|" + chk.dataset.motivo; }}

  function salvarNatureza() {{
    try {{ localStorage.setItem(CHAVE_NATUREZA, JSON.stringify(Array.from(marcadosNatureza))); }} catch (e) {{}}
  }}

  function atualizarNatureza() {{
    var marcados = Array.prototype.slice.call(document.querySelectorAll(".chk-natureza:checked"));
    // Marcacoes extras (Mediacao/Cashback/etc, ver NATUREZAS_EXTRAS): o
    // <select> nao e um checkbox, entao monta um objeto com o mesmo
    // formato (.dataset.*) esperado pelas expressoes de {campos_cmm_js}
    // (chk.dataset.motivo etc.) -- assim entra na mesma lista sem duplicar
    // a logica de montar linha do CMM/MOV.
    document.querySelectorAll(".sel-natureza-extra").forEach(function (sel) {{
      if (!sel.value) return;
      var opt = sel.options[sel.selectedIndex];
      marcados.push({{dataset: {{
        pedido: sel.dataset.pedido, motivo: opt.dataset.motivo, natureza: opt.dataset.natureza,
        conta: sel.dataset.conta, banco: sel.dataset.banco, agencia: sel.dataset.agencia,
        historico: PREFIXO_MOV ? (PREFIXO_MOV + "_" + opt.dataset.motivo) : "",
        valor: sel.dataset.valor
      }}}});
    }});
    var painel = document.getElementById("painel-natureza");
    if (!marcados.length) {{ painel.style.display = "none"; return; }}
    painel.style.display = "flex";

    var linhasCmm = marcados.map(function (chk) {{
      return [{campos_cmm_js}].join("\\t");
    }});
    document.getElementById("ta-natureza-cmm").value = linhasCmm.join("\\n");

    var linhasMov = marcados.map(function (chk) {{
      return [
        DATA_LACTO, chk.dataset.natureza, "'" + chk.dataset.pedido, "",
        chk.dataset.banco, chk.dataset.agencia, chk.dataset.conta,
        chk.dataset.historico, chk.dataset.valor
      ].join("\\t");
    }});
    document.getElementById("ta-natureza-mov").value = linhasMov.join("\\n");
  }}

  document.querySelectorAll(".chk-natureza").forEach(function (chk) {{
    if (marcadosNatureza.has(chaveNatureza(chk))) {{ chk.checked = true; }}
    chk.addEventListener("change", function () {{
      if (chk.checked) {{ marcadosNatureza.add(chaveNatureza(chk)); }}
      else {{ marcadosNatureza.delete(chaveNatureza(chk)); }}
      salvarNatureza();
      atualizarNatureza();
    }});
  }});

  // Persistencia das marcacoes extras (select), mesmo motivo do resto: o
  // reload do visibilitychange (toda vez que o Everton troca de aba)
  // perderia a escolha sem isso. Chave por pedido (nao por pedido+motivo
  // como o Set dos checkboxes) -- so 1 escolha possivel por pedido aqui.
  var CHAVE_NATUREZA_EXTRA = CHAVE + "_natureza_extra";
  var marcadosNaturezaExtra;
  try {{ marcadosNaturezaExtra = JSON.parse(localStorage.getItem(CHAVE_NATUREZA_EXTRA) || "{{}}"); }}
  catch (e) {{ marcadosNaturezaExtra = {{}}; }}

  function salvarNaturezaExtra() {{
    try {{ localStorage.setItem(CHAVE_NATUREZA_EXTRA, JSON.stringify(marcadosNaturezaExtra)); }} catch (e) {{}}
  }}

  document.querySelectorAll(".sel-natureza-extra").forEach(function (sel) {{
    var salvo = marcadosNaturezaExtra[sel.dataset.pedido];
    if (salvo) {{
      for (var i = 0; i < sel.options.length; i++) {{
        if (sel.options[i].value === salvo) {{ sel.value = salvo; break; }}
      }}
    }}
    sel.addEventListener("change", function () {{
      if (sel.value) {{ marcadosNaturezaExtra[sel.dataset.pedido] = sel.value; }}
      else {{ delete marcadosNaturezaExtra[sel.dataset.pedido]; }}
      salvarNaturezaExtra();
      atualizarNatureza();
    }});
  }});

  atualizarNatureza();
}})();

function copiarNatureza(id, botao) {{
  var ta = document.getElementById(id);
  if (navigator.clipboard) {{ navigator.clipboard.writeText(ta.value); }}
  if (botao) {{
    var textoOriginal = botao.textContent;
    botao.textContent = "Copiado!";
    setTimeout(function () {{ botao.textContent = textoOriginal; }}, 1500);
  }}
}}
</script>
</body>
</html>"""

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

    # Split Frete/Tarifa de cobranca (pedido isolado, sem passar pelo resto
    # do dashboard, pra nao mexer nas somas agregadas usadas nos testes de
    # Alvo mais abaixo): ML_ENVIO=-30 x ML_CUSTO_ENVIO=-22 -> tarifa=-8,
    # tem que gerar 2 checkboxes na MESMA celula Frete (2209/X103925 pro
    # custo de envio, 2804/X103925 pra tarifa -- SEMPRE Transitoria, mesmo
    # o CMM real mostrando "103925" naquelas linhas, ver CONTA_TRANSITORIA),
    # sem quebrar a contagem de <td> da linha (12, igual as demais colunas
    # da tabela).
    linhas_split = _linhas_tabela(
        pd.DataFrame({
            "PEDIDO_MKP": ["4"], "DIF": [-50.0],
            "ML_ENVIO": [-30.0], "ML_CUSTO_ENVIO": [-22.0], "ML_REBATE": [0.0],
        }),
        historicos={"frete": "MKP_ME_FULL_DESPESAS DE FRETE ENVIO",
                    "comissao": "MKP_ME_FULL_TARIFA DE COBRANCA"},
    )
    assert linhas_split.count("<td") == 12
    assert ('data-motivo="DESPESAS DE FRETE ENVIO" data-natureza="2209" '
            'data-conta="X103925"') in linhas_split
    assert 'data-valor="-22,00"' in linhas_split
    assert ('data-motivo="TARIFA DE COBRANCA" data-natureza="2804" '
            'data-conta="X103925"') in linhas_split
    assert 'data-historico="MKP_ME_FULL_TARIFA DE COBRANCA"' in linhas_split
    assert 'data-valor="-8,00"' in linhas_split

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
        )
        assert os.path.exists(caminho)
        assert os.path.basename(caminho) == "dashboard_0101.html"  # 1 link, sem canal no nome
        conteudo = open(caminho, encoding="utf-8").read()
        assert "VERIFICAR" in conteudo and "-R$ 921,84" in conteudo
        assert "Status dos pedidos" not in conteudo and "Distancia do alvo" not in conteudo
        assert "#14161a" in conteudo  # tema dark
        assert "Alvo nao cadastrado" in conteudo  # alvo=None
        assert "visibilitychange" in conteudo  # recarrega sozinho ao voltar pra aba
        assert "Acao sugerida" not in conteudo and 'class="acao"' not in conteudo  # coluna removida
        assert ">RBX<" in conteudo and ">RB<" in conteudo and "Diferença extrato" in conteudo
        assert "min-width: 100%" not in conteudo  # bug de CSS (colunas espalhadas) corrigido
        assert 'REBATE/c.xlsx">Baixar PROCESSAR_REBATE (1)</a>' in conteudo
        # Card Rebate (formula) x RBX: bate (cmm 50.0 + cupom 0.0 = 50.0,
        # rebate_bate=True no fixture) -> bolinha verde + sub com o detalhe.
        assert 'Rebate (formula) <span class="ponto" style="background:#2fbf4f"></span>' in conteudo
        assert 'RBX: R$ 50,00 (cupom 0 + CMM 1)' in conteudo
        # pedido "3": achado no RB com o valor exato da DIF (debito) -> DIF fica
        # neutro (pastilha cinza), RB fica colorido (texto vermelho, sem
        # pastilha), Diferenca extrato fecha em 0,00 (neutro) com semaforo
        # verde. Nao achado no RBX -> celula em branco (sem "—").
        assert "R$" not in conteudo.split("<table")[1]  # sem "R$" dentro da tabela
        # DIF negativo: pastilha cinza neutra, mas com o sinal "-" em
        # vermelho (nao o numero inteiro) -- pedido do Everton, 18/09/2026.
        assert _valor_dif_html(-921.84) in conteudo
        assert f'<span style="color:{COR_ALERTA}">-</span>921,84' in conteudo
        assert _valor_colorido_html(-921.84) in conteudo  # RB: texto vermelho, sem pastilha
        assert _valor_neutro_html(0.0) in conteudo  # Diferenca extrato: neutro
        assert f'<span class="ponto" style="background:{COR_OK}"></span>' in conteudo  # semaforo verde
        assert '<span class="ink-muted">—</span>' not in conteudo  # sem traco mudo nas celulas
        # "Outros (RA)" separada nao existe mais -- ja embutida no simulado
        assert "Outros (RA)" not in conteudo
        # linha "Transitoria" mostra o SIMULADO (0.0 + -7.0 = -7.0); alvo=None
        # nessa chamada -> nao compara contra alvo nenhum, texto neutro
        assert 'Transitoria (alvo do calendario)' in conteudo
        assert _moeda(-7.0) in conteudo
        assert '<div class="alvo-status ink-muted">sem alvo do calendario pra comparar</div>' in conteudo
        assert "NAO zerou" not in conteudo  # regra antiga (comparar contra 0) removida
        assert "pronto para fazer a transferencia" not in conteudo
        assert "zerada (transferencia ja feita)" not in conteudo

        # Marcacao de NATUREZA: pedido "3", Frete=-25.5 e Rebate=30.0 tem
        # que gerar checkbox com motivo/natureza/conta/historico corretos
        # (Comissao fica de fora por enquanto, ver comentario em
        # _linhas_tabela). HISTORICO usa prefixo_mov="MKP_ME_FULL" (FULL).
        assert 'data-motivo="DESPESAS DE FRETE ENVIO" data-natureza="2209" data-conta="X103925"' in conteudo
        assert 'data-historico="MKP_ME_FULL_DESPESAS DE FRETE ENVIO"' in conteudo
        assert 'data-motivo="RECEBIMENTO DE REBATE" data-natureza="1107" data-conta="X103925"' in conteudo
        assert 'data-historico="MKP_ME_FULL_RECEBIMENTO DE REBATE"' in conteudo
        assert 'data-valor="-25,50"' in conteudo and 'data-valor="30,00"' in conteudo
        assert 'chk-natureza' in conteudo and 'painel-natureza' in conteudo
        assert 'var RESPONSAVEL = "EVERTON";' in conteudo
        assert 'var LOJA = "FULL";' in conteudo  # loja_calendario, nao raiz.split("_")[-1] ("FU")
        assert 'var VENDEDOR = "751";' in conteudo
        assert "onclick=\"copiarNatureza('ta-natureza-cmm', this)\"" in conteudo
        assert 'botao.textContent = "Copiado!";' in conteudo

        # Guia 2 (AB / RBX / RB): abas presentes, guia-2 comeca escondida
        # (hidden), lotes/AB Para Processar/RB aparecem na tabela certa.
        assert 'data-aba="guia-1"' in conteudo and 'data-aba="guia-2"' in conteudo
        assert '<div class="guia" id="guia-2" hidden>' in conteudo
        guia2 = conteudo.split('id="guia-2"')[1]
        assert "01.01.2026" in guia2 and ">rodado<" in guia2  # lotes_status
        assert "EXTERNAL_REFERENCE" in guia2 and ">3<" in guia2  # para_processar cru
        assert _valor_colorido_html(-921.84) in guia2  # RB: pedido "3" (valores_rb do fixture)
        assert "Sem rb.xlsx/rbx.xlsx pra conferir ainda." in guia2  # RBX vazio no fixture
        # SOURCE_ID/EXTERNAL_REFERENCE (float64 por causa do None numa linha)
        # tem que sair como ID puro, NAO formatado como dinheiro.
        assert "<td>2000017591558558</td>" in guia2
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

        # Pedido colado no CMM/MOV com apostrofo na frente (igual
        # _pedido_html) -- sem isso o pedido (ate 16 digitos) vira notacao
        # cientifica ao colar no Excel.
        assert "\"'\" + chk.dataset.pedido" in conteudo
        assert '"\'" + chk.dataset.pedido, ""' in conteudo  # linhasMov

        # pedido "3" nao tem ML_CUSTO_ENVIO -> custo_envio e None, sem split
        # (tarifa_cobranca fica None) -- continua com 1 checkbox so no Frete.
        linha3 = conteudo.split('<tr data-pedido="3">')[1].split("</tr>")[0]
        assert linha3.count("<td") == 12
        assert linha3.count('class="chk-natureza"') == 2  # so Frete + Rebate
        assert "<td>mediation, refund</td>" in linha3  # coluna Descricao (AB)

        # rb_total(140) + rebate_total(50) + ajuste_pendente(-921.84) = -731.84 -> deve bater
        # liquido_transitoria_simulado batendo com o alvo -> "pronto para
        # fazer a transferencia" (verde)
        conferencia_rbx_pronta = {**conferencia_rbx, "liquido_transitoria_simulado": -731.84,
                                   "qtd_transitoria_simulado": 1, "zerada": False,
                                   "pronta_para_transferencia": True}
        caminho2 = gerar_dashboard(
            resultado, fila, rebate_cmm, canal="FULL", data="01.01",
            data_ref_rb="01/01/2026", caminho_xlsx_completo="a.xlsx",
            caminho_xlsx_rebate="b.xlsx", pasta_saida=tmp, alvo=-731.84,
            conferencia_rbx=conferencia_rbx_pronta,
        )
        assert caminho2 == caminho  # mesmo arquivo, sobrescrito (mesmo dia)
        conteudo2 = open(caminho2, encoding="utf-8").read()
        assert "NAO bate com o calendario" not in conteudo2
        assert "pronto para fazer a transferencia" in conteudo2

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
        assert "zerada (transferencia ja feita)" in conteudo_zerada
        assert "acumulando" not in conteudo_zerada
        assert "pronto para fazer a transferencia" not in conteudo_zerada

        # alvo=0.0, liquido_transitoria_simulado=-7.0 -> nem zerou nem bateu,
        # so acumulando ainda -- neutro (nao e erro no meio do dia)
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
        assert "NAO bate com o calendario" in conteudo3
        assert f"acumulando (faltam {_moeda(7.0)} pro alvo)" in conteudo3
        assert "ainda nao bate com o alvo" not in conteudo3  # texto antigo (vermelho) removido
        assert "Valores do Mercado Livre atualizados" in conteudo3
        assert "rodar_scraper_full_0101.bat" in conteudo3
        # aviso de arquivar (verde, destacado) -- primeiro entre os avisos
        # discretos, junto do resto (Scraper pronto / ML atualizado)
        assert (f'Transferencia pronta! Arquivar RBX: arquivar_rbx_full_0101.bat '
                '(uma pasta acima -- da 2 cliques nele)') in conteudo3
        assert f'style="color:{COR_OK_TEXTO}"' in conteudo3
        assert conteudo3.index("Transferencia pronta!") < conteudo3.index("rodar_scraper_full_0101.bat")
        # alerta VERIFICAR primeiro, avisos discretos (ML atualizado + scraper
        # pronto) depois, sem barra de cor cheia (botao-primario removido)
        assert conteudo3.index('class="alerta"') < conteudo3.index('class="aviso-topo"')
        assert "botao-primario" not in conteudo3

        # ES depois do FULL, mesma data -> MESMO arquivo, sobrescreve (so o
        # canal atual fica na tela -- pedido do Everton, sem guias)
        caminho4 = gerar_dashboard(
            resultado, fila, rebate_cmm, canal="ES", data="01.01",
            data_ref_rb="01/01/2026", caminho_xlsx_completo="c.xlsx",
            caminho_xlsx_rebate="d.xlsx", pasta_saida=tmp, alvo=None,
        )
        assert caminho4 == caminho  # continua sendo o mesmo 1 link
        conteudo4 = open(caminho4, encoding="utf-8").read()
        assert "Conciliacao ES - 01.01" in conteudo4
        assert "Conciliacao FULL" not in conteudo4  # FULL nao fica mais na tela

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
            linha3_extra = conteudo_extra.split('<tr data-pedido="3">')[1].split("</tr>")[0]
            assert linha3_extra.count("<td") == 13  # 12 de sempre + Marcacao
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

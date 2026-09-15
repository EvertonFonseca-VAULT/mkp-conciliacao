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
import os

STATUS_COR = {
    "SEM DIFF": "#2fbf4f",
    "REBATE": "#4a90e2",
    "EST COM": "#2dd4b0",
    "MEDIACAO": "#8a7cff",
    "VERIFICAR": "#e0524f",
}
STATUS_ORDEM = list(STATUS_COR)

COR_DIF_POS = "#4a90e2"
COR_DIF_NEG = "#e0524f"
COR_OK = "#2fbf4f"
COR_ALERTA = "#e0524f"

MAX_LINHAS_BARRA = 20
TOLERANCIA_ALVO = 0.01


def _moeda(valor):
    neg = valor < 0
    texto = f"{abs(valor):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"{'-' if neg else ''}R$ {texto}"


def _donut_css(fatias):
    stops, acumulado = [], 0.0
    for cor, pct in fatias:
        if pct <= 0:
            continue
        inicio, acumulado = acumulado, acumulado + pct
        stops.append(f"{cor} {inicio:.2f}% {acumulado:.2f}%")
    if not stops:
        return "var(--gridline)"
    return "conic-gradient(" + ", ".join(stops) + ")"


def _linhas_barra(fila_verificar):
    linhas = fila_verificar.head(MAX_LINHAS_BARRA)
    max_abs = linhas["DIF"].abs().max() if len(linhas) else 0
    saida = []
    for _, row in linhas.iterrows():
        dif = row["DIF"]
        pct = (abs(dif) / max_abs * 100) if max_abs else 0
        pedido = html.escape(str(row["PEDIDO_MKP"]))
        if dif < 0:
            saida.append(f"""
            <div class="bar-row">
              <div class="bar-label">{pedido}</div>
              <div class="bar-neg-track"><span class="bar-value">{_moeda(dif)}</span><div class="bar neg" style="width:{pct:.1f}%"></div></div>
              <div class="bar-center"></div>
              <div class="bar-pos-track"></div>
            </div>""")
        else:
            saida.append(f"""
            <div class="bar-row">
              <div class="bar-label">{pedido}</div>
              <div class="bar-neg-track"></div>
              <div class="bar-center"></div>
              <div class="bar-pos-track"><div class="bar pos" style="width:{pct:.1f}%"></div><span class="bar-value">{_moeda(dif)}</span></div>
            </div>""")
    return "".join(saida)


def _valor_ml(v):
    if v is None or (isinstance(v, float) and v != v):  # None ou NaN
        return '<span class="ink-muted">—</span>'
    if isinstance(v, (int, float)):
        return _moeda(v)
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


def _linhas_tabela(fila_verificar):
    saida = []
    for _, row in fila_verificar.iterrows():
        pedido = html.escape(str(row["PEDIDO_MKP"]))
        saida.append(f"""
            <tr data-pedido="{pedido}">
              <td class="chk-col"><input type="checkbox" class="chk-ajustado" data-pedido="{pedido}"></td>
              <td style="mso-number-format:'\\@'">{_pedido_html(row['PEDIDO_MKP'])}</td>
              <td class="num">{_valor_copiavel_html(_moeda(row['DIF']))}</td>
              <td class="num">{_valor_ml(row.get('ML_COMISSAO'))}</td>
              <td class="num">{_valor_ml(row.get('ML_ENVIO'))}</td>
              <td class="num">{_valor_ml(row.get('ML_TOTAL'))}</td>
              <td class="num">{_valor_ml(row.get('ML_REBATE'))}</td>
              <td class="acao">{html.escape(str(row.get('ACAO_SUGERIDA', '')))}</td>
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


def gerar_dashboard(resultado, fila_verificar, rebate_cmm, *, canal, data, data_ref_rb,
                     caminho_xlsx_completo, caminho_xlsx_rebate, pasta_saida, alvo=None,
                     ml_atualizado=False, caminho_bat_scraper=None):
    contagem = resultado["STATUS"].value_counts()
    total = len(resultado)

    fatias = [(STATUS_COR[s], (contagem.get(s, 0) / total * 100) if total else 0) for s in STATUS_ORDEM]
    donut = _donut_css(fatias)

    legenda = "".join(
        f'<div class="legenda-item"><span class="ponto" style="background:{STATUS_COR[s]}"></span>'
        f'{s} <span class="ink-muted">({contagem.get(s, 0)})</span></div>'
        for s in STATUS_ORDEM
    )

    rebate_total = resultado.loc[resultado["STATUS"] == "REBATE", "REBATE"].sum()
    rb_total = resultado["TOTAL_RB"].sum()
    ajuste_pendente = fila_verificar["DIF"].sum()
    verificar_total = fila_verificar["DIF"].abs().sum()

    cards = [
        ("Total de pedidos", f"{total}", None),
        ("Rebate (formula)", _moeda(rebate_total), f"{contagem.get('REBATE', 0)} pedido(s)"),
        ("Est. comissao", f"{contagem.get('EST COM', 0)}", "resolvido pela formula"),
        ("Mediacao", f"{contagem.get('MEDIACAO', 0)}", "resolvido pela formula"),
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

    barras_html = _linhas_barra(fila_verificar) if verificar_qtd else '<p class="ink-muted">Nenhum pedido em VERIFICAR.</p>'
    tabela_html = (
        _linhas_tabela(fila_verificar) if verificar_qtd
        else '<tr><td colspan="8" class="ink-muted">Nenhum pedido em VERIFICAR.</td></tr>'
    )

    botoes = []
    if caminho_bat_scraper:
        botoes.append(
            f'<a class="botao botao-primario" href="{html.escape(caminho_bat_scraper)}" download>'
            f'Rodar scraper (baixa .bat)</a>'
        )
    botoes.append(f'<a class="botao" href="{html.escape(caminho_xlsx_completo)}">Baixar planilha completa</a>')
    if len(rebate_cmm):
        botoes.append(f'<a class="botao" href="{html.escape(caminho_xlsx_rebate)}">Baixar XLSX Rebate CMM ({len(rebate_cmm)})</a>')
    botoes_html = "".join(botoes)

    titulo = f"Conciliacao {canal} - {data}"
    chave_local_storage = f"ajustados_{canal.lower()}_{data.replace('.', '')}"
    aviso_ml_html = (
        f'<div class="aviso-ml"><span class="ponto" style="background:{COR_OK}"></span>'
        f'Valores do Mercado Livre atualizados</div>'
        if ml_atualizado else ""
    )

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
  html, body {{ height: 100%; margin: 0; overflow: hidden; }}
  body {{
    padding: 20px; background: var(--page); color: var(--ink);
    font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
  }}
  .wrap {{ height: calc(100vh - 40px); max-width: 100%; margin: 0 auto; display: flex; flex-direction: column; gap: 14px; }}
  .header {{ flex: 0 0 auto; display: flex; align-items: baseline; justify-content: space-between; gap: 16px; flex-wrap: wrap; }}
  h1 {{ font-size: 1.3rem; margin: 0; }}
  .ref {{ color: var(--ink-secondary); margin: 0; font-size: 0.85rem; }}
  .ink-muted {{ color: var(--ink-muted); }}
  .aviso-ml {{
    flex: 0 0 auto; background: #16281c; border: 1px solid {COR_OK}55; color: #8fe0a0;
    border-radius: 8px; padding: 8px 16px; font-size: 0.82rem; display: flex; align-items: center; gap: 8px;
  }}
  .aviso-ml .ponto {{ width: 8px; height: 8px; border-radius: 50%; display: inline-block; }}
  .alerta {{
    flex: 0 0 auto; background: #2e1a1a; border: 1px solid #e0524f55; color: #ff9b93;
    border-radius: 8px; padding: 10px 16px; font-size: 0.9rem;
  }}
  .alerta .icone {{
    display: inline-flex; width: 18px; height: 18px; border-radius: 50%;
    background: #e0524f; color: #14161a; align-items: center; justify-content: center;
    font-weight: bold; margin-right: 6px; font-size: 0.8rem;
  }}
  .grid-main {{ flex: 1 1 auto; min-height: 0; display: grid; grid-template-columns: 260px 320px 1fr; gap: 14px; }}
  @media (max-width: 1000px) {{ html, body {{ overflow: auto; }} .wrap {{ height: auto; }} .grid-main {{ grid-template-columns: 1fr; }} }}
  .col {{ min-height: 0; min-width: 0; display: flex; flex-direction: column; gap: 14px; }}
  .painel {{
    background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 14px;
    min-height: 0; min-width: 0; display: flex; flex-direction: column;
  }}
  .painel h2 {{ font-size: 0.85rem; margin: 0 0 10px; color: var(--ink-secondary); text-transform: uppercase; letter-spacing: 0.03em; }}
  .cards {{ display: flex; flex-direction: column; gap: 10px; flex: 0 0 auto; }}
  .card {{ background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 12px 14px; }}
  .card-label {{ font-size: 0.75rem; color: var(--ink-secondary); }}
  .card-valor {{ font-size: 1.3rem; font-weight: 600; margin-top: 2px; }}
  .card-sub {{ font-size: 0.7rem; color: var(--ink-muted); margin-top: 2px; }}
  .donut-box {{ display: flex; flex-direction: column; align-items: center; gap: 14px; }}
  .donut {{ width: 110px; height: 110px; border-radius: 50%; background: {donut}; flex-shrink: 0; }}
  .legenda {{ display: flex; flex-direction: column; gap: 6px; font-size: 0.8rem; width: 100%; }}
  .legenda-item {{ display: flex; align-items: center; gap: 6px; }}
  .ponto {{ width: 9px; height: 9px; border-radius: 50%; display: inline-block; flex-shrink: 0; }}
  .alvo-linha {{ display: flex; justify-content: space-between; gap: 10px; font-size: 0.82rem; padding: 4px 0; border-bottom: 1px dashed var(--gridline); }}
  .alvo-linha span:first-child {{ flex: 1 1 auto; min-width: 0; overflow-wrap: break-word; }}
  .alvo-linha .num {{ flex: 0 0 auto; white-space: nowrap; font-variant-numeric: tabular-nums; }}
  .alvo-destaque {{ border-bottom: none; padding-top: 8px; font-weight: 600; font-size: 0.9rem; }}
  .alvo-status {{ display: flex; align-items: center; gap: 6px; font-size: 0.8rem; margin-top: 8px; }}
  .bar-row {{ display: grid; grid-template-columns: 80px 1fr 2px 1fr; align-items: center; gap: 4px; margin-bottom: 6px; font-size: 0.75rem; }}
  .bar-label {{ color: var(--ink-secondary); font-variant-numeric: tabular-nums; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
  .bar-center {{ background: var(--gridline); height: 100%; }}
  .bar-neg-track {{ display: flex; justify-content: flex-end; align-items: center; gap: 6px; }}
  .bar-pos-track {{ display: flex; align-items: center; gap: 6px; }}
  .bar {{ height: 16px; border-radius: 4px; }}
  .bar.neg {{ background: {COR_DIF_NEG}; }}
  .bar.pos {{ background: {COR_DIF_POS}; }}
  .bar-value {{ font-variant-numeric: tabular-nums; white-space: nowrap; }}
  .barras-scroll {{ overflow-y: auto; min-height: 0; }}
  table {{ width: max-content; min-width: 100%; border-collapse: collapse; font-size: 0.8rem; }}
  th, td {{ text-align: left; padding: 7px 10px; border-bottom: 1px solid var(--gridline); white-space: nowrap; }}
  th {{ position: sticky; top: 0; background: var(--surface); color: var(--ink-secondary); font-weight: 500; }}
  td.num {{ font-variant-numeric: tabular-nums; }}
  td.acao {{ white-space: normal; min-width: 420px; color: var(--ink-secondary); }}
  td.chk-col, th.chk-col {{ width: 28px; text-align: center; padding: 7px 4px; }}
  .chk-ajustado {{ width: 15px; height: 15px; accent-color: {COR_OK}; cursor: pointer; }}
  tr.linha-feita td {{ text-decoration: line-through; color: var(--ink-muted); }}
  .tabela-verificar {{ flex: 1 1 auto; min-height: 0; }}
  .tabela-scroll {{ overflow: auto; height: 100%; }}
  .botoes {{ flex: 0 0 auto; display: flex; gap: 10px; margin-top: 4px; flex-wrap: wrap; }}
  .botao {{
    display: inline-block; padding: 8px 14px; border-radius: 8px; text-decoration: none;
    background: var(--surface-2); border: 1px solid var(--border); color: var(--ink); font-size: 0.85rem;
  }}
  .botao-primario {{ background: #4a90e2; color: #0b1420; border-color: #4a90e2; font-weight: 600; }}
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

  {aviso_ml_html}
  {alerta_html}

  <div class="grid-main">
    <div class="col">
      <div class="cards">{cards_html}</div>
      <div class="painel">
        <h2>Status dos pedidos</h2>
        <div class="donut-box">
          <div class="donut"></div>
          <div class="legenda">{legenda}</div>
        </div>
      </div>
    </div>

    <div class="col">
      <div class="painel">
        <h2>Alvo do dia</h2>
        {alvo_html}
      </div>
      <div class="painel" style="flex: 1 1 auto;">
        <h2>Distancia do alvo -- fila VERIFICAR (DIF)</h2>
        <div class="barras-scroll">{barras_html}</div>
      </div>
    </div>

    <div class="col">
      <div class="painel tabela-verificar">
        <h2>Fila VERIFICAR</h2>
        <div class="tabela-scroll">
          <table>
            <thead><tr>
              <th class="chk-col"></th><th>Pedido</th><th>DIF</th>
              <th>Comissao</th><th>Frete</th><th>Repasse</th><th>Rebate</th>
              <th>Acao sugerida</th>
            </tr></thead>
            <tbody>{tabela_html}</tbody>
          </table>
        </div>
      </div>
    </div>
  </div>
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

  // Volta pra essa aba do navegador (depois de rodar o processo de novo,
  // outro canal ou nao) -> recarrega sozinho, sem precisar abrir outro link.
  document.addEventListener("visibilitychange", function () {{
    if (document.visibilityState === "visible") {{ location.reload(); }}
  }});
}})();
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

    resultado = pd.DataFrame({
        "PEDIDO_MKP": ["1", "2", "3"],
        "STATUS": ["SEM DIFF", "REBATE", "VERIFICAR"],
        "REBATE": [0.0, 50.0, 0.0],
        "TOTAL_RB": [100.0, 40.0, 0.0],
    })
    fila = pd.DataFrame({
        "PEDIDO_MKP": ["3"], "DIF": [-921.84], "ACAO_SUGERIDA": ["Conferir Saldo no TOTVS."],
    })
    rebate_cmm = pd.DataFrame({"VALOR R$": [50.0]})

    with tempfile.TemporaryDirectory() as tmp:
        caminho = gerar_dashboard(
            resultado, fila, rebate_cmm, canal="FULL", data="01.01",
            data_ref_rb="01/01/2026", caminho_xlsx_completo="a.xlsx",
            caminho_xlsx_rebate="b.xlsx", pasta_saida=tmp, alvo=None,
        )
        assert os.path.exists(caminho)
        assert os.path.basename(caminho) == "dashboard_0101.html"  # 1 link, sem canal no nome
        conteudo = open(caminho, encoding="utf-8").read()
        assert "VERIFICAR" in conteudo and "-R$ 921,84" in conteudo and "conic-gradient" in conteudo
        assert "#14161a" in conteudo  # tema dark
        assert "Alvo nao cadastrado" in conteudo  # alvo=None
        assert "visibilitychange" in conteudo  # recarrega sozinho ao voltar pra aba

        # rb_total(140) + rebate_total(50) + ajuste_pendente(-921.84) = -731.84 -> deve bater
        caminho2 = gerar_dashboard(
            resultado, fila, rebate_cmm, canal="FULL", data="01.01",
            data_ref_rb="01/01/2026", caminho_xlsx_completo="a.xlsx",
            caminho_xlsx_rebate="b.xlsx", pasta_saida=tmp, alvo=-731.84,
        )
        assert caminho2 == caminho  # mesmo arquivo, sobrescrito (mesmo dia)
        conteudo2 = open(caminho2, encoding="utf-8").read()
        assert "NAO bate com o calendario" not in conteudo2

        caminho3 = gerar_dashboard(
            resultado, fila, rebate_cmm, canal="FULL", data="01.01",
            data_ref_rb="01/01/2026", caminho_xlsx_completo="a.xlsx",
            caminho_xlsx_rebate="b.xlsx", pasta_saida=tmp, alvo=0.0,
            ml_atualizado=True, caminho_bat_scraper="rodar_scraper_full_0101.bat",
        )
        conteudo3 = open(caminho3, encoding="utf-8").read()
        assert "NAO bate com o calendario" in conteudo3
        assert "Valores do Mercado Livre atualizados" in conteudo3
        assert "rodar_scraper_full_0101.bat" in conteudo3

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
        print("dashboard.py: ok")

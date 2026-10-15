"""Validação/Pagamento/Frete/Para Processar do Mercado COLETAS ES — reusa
`montar_ab_pp`/`carregar_pti` do FULL (100% genéricos, não dependem do
layout bruto do AB, só do formato de 15 colunas já resolvido)."""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "FULL"))
from modelo_ab_full import montar_ab_pp, carregar_pti  # noqa: F401

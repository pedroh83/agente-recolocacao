# -*- coding: utf-8 -*-
"""
Pontuacao de cada vaga contra as palavras-chave dos curriculos A e B.

A nota e a soma dos pesos dos termos encontrados, mais um bonus quando a
vaga cita uma localidade de interesse. Vence o curriculo com mais pontos.
"""
from __future__ import annotations

import re
from typing import Dict, List, Tuple

from configuracao import PalavrasChave
from texto import normalizar

_CACHE_TERMOS: Dict[str, "re.Pattern[str]"] = {}


def casa_termo(texto_norm: str, termo: str) -> bool:
    """O termo aparece como palavra inteira?

    Sem isso, 'ti' casaria dentro de 'logistica' e 'pop' dentro de 'popular'.
    Os limites (?<![a-z0-9]) e (?![a-z0-9]) exigem que nao haja letra ou
    numero colado antes nem depois.
    """
    padrao = _CACHE_TERMOS.get(termo)
    if padrao is None:
        padrao = re.compile(r"(?<![a-z0-9])" + re.escape(termo) + r"(?![a-z0-9])")
        _CACHE_TERMOS[termo] = padrao
    return bool(padrao.search(texto_norm))


def pontuar_contra(texto_norm: str, pesos: Dict[int, List[str]]) -> Tuple[int, List[str]]:
    nota, achados = 0, []
    for peso in sorted(pesos, reverse=True):  # termos fortes primeiro na lista de achados
        for termo in pesos[peso]:
            if casa_termo(texto_norm, termo):
                nota += peso
                achados.append(termo)
    return nota, achados


def avaliar(vaga: dict, palavras: PalavrasChave) -> dict:
    """Preenche na vaga: nota, curriculo (A, B, 'A ou B' ou '?'), termos e descartada."""
    # O descarte olha SO o titulo: o contexto pode citar outra coisa
    titulo_norm = normalizar(vaga.get("titulo", ""))
    for termo in palavras.descarte:
        if casa_termo(titulo_norm, termo):
            vaga.update(descartada=f"contem '{termo}'", nota=0, curriculo="-", termos="")
            return vaga

    texto = normalizar(" ".join(vaga.get(k, "") for k in ("titulo", "empresa", "local", "contexto")))
    nota_a, achados_a = pontuar_contra(texto, palavras.curriculos["A"])
    nota_b, achados_b = pontuar_contra(texto, palavras.curriculos["B"])
    bonus = palavras.bonus_local if any(casa_termo(texto, t) for t in palavras.local) else 0

    if nota_b > nota_a:
        curriculo, nota, achados = "B", nota_b + bonus, achados_b
    elif nota_a > nota_b:
        curriculo, nota, achados = "A", nota_a + bonus, achados_a
    elif nota_a == 0:
        curriculo, nota, achados = "?", bonus, []
    else:
        curriculo, nota, achados = "A ou B", nota_a + bonus, achados_a + achados_b

    vaga.update(descartada="", nota=nota, curriculo=curriculo, termos=", ".join(achados[:6]))
    return vaga

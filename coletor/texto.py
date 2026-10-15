# -*- coding: utf-8 -*-
"""Funcoes de texto usadas por todos os modulos do agente."""
from __future__ import annotations

import re
import unicodedata
from html import unescape


def normalizar(texto: str) -> str:
    """Minusculas e sem acento, para comparar palavras com seguranca."""
    if not texto:
        return ""
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto.lower()


def limpar_html(fragmento: str) -> str:
    """Tira as tags e junta os espacos de um pedaco de HTML.

    O recorte de contexto pode cortar uma tag pela metade ('<table cellp'),
    entao tambem removemos qualquer '<' sem o '>' correspondente.
    """
    txt = fragmento or ""
    txt = re.sub(r"<[^>]*>", " ", txt)   # tags completas
    txt = re.sub(r"<[^>]*$", " ", txt)   # tag cortada no fim
    txt = re.sub(r"^[^<>]*>", " ", txt)  # tag cortada no inicio
    txt = unescape(txt)
    return re.sub(r"\s+", " ", txt).strip()

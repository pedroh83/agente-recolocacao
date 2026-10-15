import os
import sys

import pytest

PASTA_COLETOR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PASTA_EXEMPLOS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "exemplos")
sys.path.insert(0, PASTA_COLETOR)


@pytest.fixture
def palavras():
    from configuracao import carregar_palavras_chave
    return carregar_palavras_chave(os.path.join(PASTA_COLETOR, "palavras_chave.txt"))


@pytest.fixture
def html_exemplo():
    with open(os.path.join(PASTA_EXEMPLOS, "alerta_linkedin_exemplo.html"), encoding="utf-8") as f:
        return f.read()

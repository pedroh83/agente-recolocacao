# -*- coding: utf-8 -*-
"""
Leitura da configuracao do agente (config.txt) e das listas de
palavras-chave (palavras_chave.txt).

Nada aqui usa variaveis globais: cada funcao devolve um objeto com os
valores lidos, e quem precisa recebe esse objeto como parametro.
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from texto import normalizar

PASTA_AGENTE = os.path.dirname(os.path.abspath(__file__))


class ErroConfiguracao(Exception):
    """Problema no config.txt ou no palavras_chave.txt, com mensagem amigavel."""


def ler_linhas(caminho: str) -> List[str]:
    """Le um arquivo de texto salvo em UTF-8 ou em ANSI (o 'ANSI' do Bloco de Notas)."""
    for codificacao in ("utf-8-sig", "cp1252"):
        try:
            with open(caminho, encoding=codificacao) as f:
                return f.read().splitlines()
        except UnicodeDecodeError:
            continue
    with open(caminho, encoding="utf-8", errors="replace") as f:
        return f.read().splitlines()


# ---------------------------------------------------------------------------
# config.txt
# ---------------------------------------------------------------------------

@dataclass
class Config:
    email: str = ""
    senha: str = ""
    pasta: str = "INBOX"
    nota_minima: int = 1
    dias: int = 2
    servidor: str = "imap.gmail.com"
    # Onde ficam a planilha do dia, o historico e o controle de candidaturas
    dir_saida: str = field(default_factory=lambda: os.path.join(os.path.expanduser("~"), "AgenteVagas"))
    # Pasta do adaptador de curriculo (para ler o historico.csv e a chave do Gemini)
    pasta_adaptador: str = field(default_factory=lambda: os.path.join(PASTA_AGENTE, "Adaptador"))
    # Reavaliar com IA os titulos de nota baixa: "auto" (liga se houver chave), "sim" ou "nao"
    ia_titulos: str = "auto"
    ia_maximo: int = 40
    # Planilhas do dia mais velhas que isso vao para PODE_APAGAR (0 = nunca)
    dias_para_arquivar: int = 7

    @property
    def senha_sem_espacos(self) -> str:
        # O Google mostra a senha de app com espacos; para o login eles nao servem
        return self.senha.replace(" ", "").strip()


def _inteiro(valor: str, chave: str, linha: int) -> int:
    try:
        return int(valor)
    except ValueError:
        raise ErroConfiguracao(
            f"config.txt, linha {linha}: '{chave}' precisa ser um numero inteiro (veio '{valor}').")


def carregar_config(caminho: Optional[str] = None) -> Config:
    """Le o config.txt (formato 'chave = valor'). Linhas com # sao comentarios."""
    caminho = caminho or os.path.join(PASTA_AGENTE, "config.txt")
    if not os.path.exists(caminho):
        raise ErroConfiguracao(
            "Nao achei o arquivo config.txt na pasta do agente.\n"
            f"Ele deve ficar em {caminho}\n"
            "Copie o config.example.txt com esse nome e preencha.")

    cfg = Config()
    for n, linha in enumerate(ler_linhas(caminho), start=1):
        linha = linha.strip()
        if not linha or linha.startswith("#") or "=" not in linha:
            continue
        chave, valor = (p.strip() for p in linha.split("=", 1))
        chave = chave.lower()
        if chave == "email":
            cfg.email = valor
        elif chave == "senha":
            cfg.senha = valor
        elif chave == "pasta":
            cfg.pasta = valor or "INBOX"
        elif chave == "nota_minima":
            cfg.nota_minima = _inteiro(valor, chave, n)
        elif chave == "dias":
            cfg.dias = _inteiro(valor, chave, n)
        elif chave == "pasta_saida" and valor:
            cfg.dir_saida = os.path.expanduser(valor)
        elif chave == "pasta_adaptador" and valor:
            cfg.pasta_adaptador = os.path.expanduser(valor)
        elif chave == "ia_titulos":
            cfg.ia_titulos = normalizar(valor) or "auto"
        elif chave == "ia_maximo":
            cfg.ia_maximo = _inteiro(valor, chave, n)
        elif chave == "dias_para_arquivar":
            cfg.dias_para_arquivar = _inteiro(valor, chave, n)
        # chaves desconhecidas sao ignoradas, para nao quebrar configs antigos
    return cfg


# ---------------------------------------------------------------------------
# palavras_chave.txt
# ---------------------------------------------------------------------------

@dataclass
class PalavrasChave:
    """Termos de cada curriculo, separados por peso, mais descarte e localidade."""
    curriculos: Dict[str, Dict[int, List[str]]]
    descarte: List[str]
    local: List[str]
    bonus_local: int = 2


SECOES_CURRICULO = {"a": "A", "b": "B"}


def _termos(texto: str) -> List[str]:
    """Separa por virgula, normaliza (minusculo e sem acento) e descarta vazios."""
    termos = []
    for t in texto.split(","):
        t = " ".join(normalizar(t).split())
        if t and t not in termos:
            termos.append(t)
    return termos


def carregar_palavras_chave(caminho: Optional[str] = None) -> PalavrasChave:
    """
    Formato do arquivo (veja o palavras_chave.txt que acompanha o agente):

        [A]
        peso 3: iso 9001, auditoria interna, ...
        peso 1: qualidade, pdca, ...

        [B]
        peso 3: analista de requisitos, ...
        peso 1: scrum, jira, ...

        [DESCARTE]
        estagio, jovem aprendiz, ...

        [LOCAL]
        bonus: 2
        salvador, bahia, remoto, ...

    Acentos e maiusculas podem ser usados a vontade: tudo e normalizado.
    """
    caminho = caminho or os.path.join(PASTA_AGENTE, "palavras_chave.txt")
    if not os.path.exists(caminho):
        raise ErroConfiguracao(
            f"Nao achei o arquivo de palavras-chave: {caminho}\n"
            "Ele acompanha o agente. Copie o palavras_chave.txt para a pasta do agente.")

    curriculos: Dict[str, Dict[int, List[str]]] = {}
    descarte: List[str] = []
    local: List[str] = []
    bonus = 2
    secao = None

    for n, bruta in enumerate(ler_linhas(caminho), start=1):
        linha = bruta.split("#", 1)[0].strip()
        if not linha:
            continue

        if linha.startswith("[") and linha.endswith("]"):
            nome = normalizar(linha[1:-1]).strip()
            letra = nome.split()[0] if nome else ""
            if letra in SECOES_CURRICULO:
                secao = SECOES_CURRICULO[letra]
                curriculos.setdefault(secao, {})
            elif nome in ("descarte", "local"):
                secao = nome
            else:
                raise ErroConfiguracao(
                    f"palavras_chave.txt, linha {n}: secao '{linha}' desconhecida. "
                    "Use [A], [B], [DESCARTE] ou [LOCAL].")
            continue

        if secao is None:
            raise ErroConfiguracao(
                f"palavras_chave.txt, linha {n}: texto fora de uma secao. "
                "Comece com [A], [B], [DESCARTE] ou [LOCAL].")

        if secao in ("A", "B"):
            if ":" not in linha:
                raise ErroConfiguracao(
                    f"palavras_chave.txt, linha {n}: nas secoes [A] e [B] cada linha "
                    "precisa comecar com 'peso 3:' ou 'peso 1:'.")
            rotulo, resto = linha.split(":", 1)
            digitos = "".join(c for c in rotulo if c.isdigit())
            if not digitos:
                raise ErroConfiguracao(
                    f"palavras_chave.txt, linha {n}: nao entendi o peso em '{rotulo.strip()}'. "
                    "Escreva, por exemplo, 'peso 3:'.")
            lista = curriculos[secao].setdefault(int(digitos), [])
            lista.extend(t for t in _termos(resto) if t not in lista)
        elif secao == "local" and normalizar(linha).startswith("bonus"):
            digitos = "".join(c for c in linha if c.isdigit())
            bonus = int(digitos) if digitos else bonus
        elif secao == "local":
            local.extend(t for t in _termos(linha) if t not in local)
        else:
            descarte.extend(t for t in _termos(linha) if t not in descarte)

    for letra in ("A", "B"):
        if not curriculos.get(letra):
            raise ErroConfiguracao(
                f"palavras_chave.txt: a secao [{letra}] esta vazia ou faltando.")
    return PalavrasChave(curriculos=curriculos, descarte=descarte, local=local, bonus_local=bonus)

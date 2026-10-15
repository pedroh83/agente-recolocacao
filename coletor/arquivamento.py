# -*- coding: utf-8 -*-
"""
Arrumacao automatica da pasta de saida.

A cada execucao, depois que o controle de candidaturas foi atualizado, os
arquivos que nao sao mais necessarios vao para a subpasta PODE_APAGAR:

  - planilhas do dia (vagas_AAAA-MM-DD*.xlsx) com mais de N dias. O "Enviado?"
    delas ja foi levado para o controle, entao podem sair da pasta principal;
  - copias feitas em correcoes (arquivos com "_antes_da_correcao" no nome).

Nunca mexe no controle, no backup do controle nem no historico.
Arquivo aberto no Excel fica onde esta e e movido numa proxima execucao.
"""
from __future__ import annotations

import datetime as dt
import os
import re
from typing import List, Optional

NOME_PASTA = "PODE_APAGAR"

PADRAO_PLANILHA_DO_DIA = re.compile(r"^vagas_(\d{4}-\d{2}-\d{2})(?:_\d+)?\.xlsx$", re.I)
MARCA_CORRECAO = "_antes_da_correcao"

LEIA_ME = """ARQUIVOS QUE PODEM SER APAGADOS
===============================

O Agente de Vagas move para ca, sozinho, o que nao e mais necessario:

  - planilhas do dia com mais de {dias} dias (o "Enviado?" delas ja esta
    no controle_candidaturas);
  - copias de seguranca feitas em correcoes ("..._antes_da_correcao").

O agente nao usa nada desta pasta. Pode apagar os arquivos, ou a pasta
inteira, quando quiser: ela e criada de novo quando for preciso.
"""


def _destino_livre(pasta: str, nome: str) -> str:
    """Caminho dentro da pasta sem sobrescrever um arquivo que ja esta la."""
    base, ext = os.path.splitext(nome)
    caminho, n = os.path.join(pasta, nome), 2
    while os.path.exists(caminho):
        caminho = os.path.join(pasta, f"{base} ({n}){ext}")
        n += 1
    return caminho


def arquivos_para_arquivar(dir_saida: str, hoje: dt.date, dias: int,
                           incluir_planilhas: bool = True) -> List[str]:
    if not os.path.isdir(dir_saida):
        return []
    limite = hoje - dt.timedelta(days=dias)
    escolhidos = []
    for nome in sorted(os.listdir(dir_saida)):
        caminho = os.path.join(dir_saida, nome)
        if not os.path.isfile(caminho) or nome.startswith("~$"):  # ~$ = arquivo temporario do Excel aberto
            continue
        if MARCA_CORRECAO in nome.lower():
            escolhidos.append(nome)
            continue
        m = PADRAO_PLANILHA_DO_DIA.match(nome)
        if incluir_planilhas and dias > 0 and m:
            try:
                data = dt.date.fromisoformat(m.group(1))
            except ValueError:
                continue
            if data < limite:
                escolhidos.append(nome)
    return escolhidos


def arrumar_pasta(dir_saida: str, dias: int = 7, hoje: Optional[dt.date] = None,
                  controle_atualizado: bool = True) -> List[str]:
    """
    Move os arquivos dispensaveis para PODE_APAGAR. Devolve os nomes movidos.

    Se o controle nao foi atualizado nesta execucao, as planilhas do dia ficam
    onde estao: o "Enviado?" delas pode ainda nao ter sido levado ao controle.
    """
    hoje = hoje or dt.date.today()
    nomes = arquivos_para_arquivar(dir_saida, hoje, dias, incluir_planilhas=controle_atualizado)
    if not nomes:
        return []

    pasta = os.path.join(dir_saida, NOME_PASTA)
    os.makedirs(pasta, exist_ok=True)
    leia_me = os.path.join(pasta, "LEIA-ME.txt")
    if not os.path.exists(leia_me):
        with open(leia_me, "w", encoding="utf-8-sig", newline="\r\n") as f:
            f.write(LEIA_ME.format(dias=dias))

    movidos = []
    for nome in nomes:
        try:
            os.replace(os.path.join(dir_saida, nome), _destino_livre(pasta, nome))
            movidos.append(nome)
        except OSError:
            continue  # aberto no Excel ou bloqueado: tenta de novo na proxima execucao
    return movidos

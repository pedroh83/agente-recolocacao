# -*- coding: utf-8 -*-
"""
Segunda opiniao da IA (Google Gemini) sobre titulos de vaga com nota baixa.

A pontuacao por palavra-chave so reconhece termos cadastrados. Um cargo como
"Especialista em Excelencia Operacional" pode ficar com nota baixa mesmo sendo
do perfil. Aqui, os titulos que ficaram abaixo da nota minima sao enviados
numa unica chamada ao Gemini, que responde se cada um combina com o
curriculo A, com o B ou com nenhum.

- Usa a mesma chave gratuita do Adaptador de Curriculo (config_adaptador.txt).
- E opcional: se nao houver chave, ou se a IA falhar, o agente segue normalmente.
- Vagas aprovadas pela IA entram no relatorio com a nota minima e a marca "IA".
"""
from __future__ import annotations

import json
import os
import re
from typing import Callable, Dict, List, Optional

from configuracao import Config

MODELOS_RESERVA = ["gemini-flash-latest", "gemini-2.5-flash"]

PROMPT = """Classifique titulos de vagas de emprego para um candidato com dois perfis:

A = Qualidade e Processos: gestao da qualidade, ISO 9001/14001/45001, auditoria, melhoria continua,
    mapeamento e analise de processos, excelencia operacional, controles internos, compliance.
B = Requisitos e Tecnologia: analise de requisitos, analise de negocios, business analyst,
    product owner, QA e testes de software, analista de sistemas ou funcional, implantacao de sistemas.

O candidato tem nivel analista (pleno/senior). NAO servem: estagio, vendas, atendimento, operacao de
maquinas, logistica operacional, engenharia de campo, desenvolvimento de software (programador),
cargos de diretoria.

Para cada titulo, responda "A", "B" ou "N" (nao combina). Na duvida, responda "N".
Responda APENAS com JSON, sem texto extra, neste formato:
[{{"i": 0, "perfil": "N"}}, {{"i": 1, "perfil": "A"}}]

TITULOS:
{titulos}
"""


def ler_config_adaptador(pasta_adaptador: str) -> Dict[str, str]:
    caminho = os.path.join(pasta_adaptador, "config_adaptador.txt")
    cfg: Dict[str, str] = {}
    if os.path.exists(caminho):
        with open(caminho, encoding="utf-8-sig") as f:
            for linha in f:
                linha = linha.strip()
                if linha and not linha.startswith("#") and "=" in linha:
                    chave, valor = linha.split("=", 1)
                    cfg[chave.strip().upper()] = valor.strip()
    if os.environ.get("GEMINI_API_KEY"):
        cfg["GEMINI_API_KEY"] = os.environ["GEMINI_API_KEY"]
    return cfg


def chave_disponivel(cfg_ia: Dict[str, str]) -> bool:
    chave = cfg_ia.get("GEMINI_API_KEY", "")
    return len(chave) >= 20 and "COLE" not in chave.upper()


def ia_ligada(cfg: Config, cfg_ia: Dict[str, str]) -> bool:
    if cfg.ia_titulos in ("nao", "n", "0", "desligado", "off"):
        return False
    return chave_disponivel(cfg_ia)


def interpretar_resposta(texto: str, quantidade: int) -> Dict[int, str]:
    """Le o JSON da IA e devolve {indice: 'A' ou 'B'} (o que for 'N' fica de fora)."""
    texto = re.sub(r"^```(?:json)?\s*|\s*```$", "", (texto or "").strip())
    inicio, fim = texto.find("["), texto.rfind("]")
    if inicio == -1 or fim == -1:
        raise ValueError("a IA nao devolveu uma lista JSON")
    resultado = {}
    for item in json.loads(texto[inicio:fim + 1]):
        try:
            i = int(item.get("i"))
        except (TypeError, ValueError, AttributeError):
            continue
        perfil = str(item.get("perfil", "")).strip().upper()[:1]
        if 0 <= i < quantidade and perfil in ("A", "B"):
            resultado[i] = perfil
    return resultado


def _chamar_gemini(prompt: str, cfg_ia: Dict[str, str], pasta_adaptador: str) -> str:
    from google import genai
    from google.genai import types

    cliente = genai.Client(api_key=cfg_ia["GEMINI_API_KEY"],
                           http_options=types.HttpOptions(timeout=90 * 1000))
    arq_ok = os.path.join(pasta_adaptador, "ultimo_modelo_ok.txt")
    ultimo_ok = ""
    if os.path.exists(arq_ok):
        with open(arq_ok, encoding="utf-8") as f:
            ultimo_ok = f.read().strip()
    ordem = [ultimo_ok, cfg_ia.get("MODELO", "")] + MODELOS_RESERVA
    modelos = [m for i, m in enumerate(ordem) if m and m not in ordem[:i]]

    ultimo_erro: Optional[Exception] = None
    for modelo in modelos:
        try:
            resp = cliente.models.generate_content(
                model=modelo, contents=prompt,
                config=types.GenerateContentConfig(
                    temperature=0, response_mime_type="application/json",
                    automatic_function_calling=types.AutomaticFunctionCallingConfig(disable=True)))
            return resp.text or ""
        except Exception as e:  # noqa: BLE001
            ultimo_erro = e
            continue
    raise RuntimeError(f"nenhum modelo respondeu ({str(ultimo_erro)[:120]})")


def reavaliar_titulos(vagas: List[dict], cfg: Config, nota_minima: int,
                      chamar: Optional[Callable[[str], str]] = None,
                      cfg_ia: Optional[Dict[str, str]] = None,
                      avisar: Callable[[str], None] = print) -> int:
    """
    Pede a opiniao da IA sobre as vagas abaixo da nota minima (nao descartadas).
    Altera as vagas aprovadas e devolve quantas foram resgatadas.
    """
    candidatas = [v for v in vagas if not v.get("descartada") and v["nota"] < nota_minima]
    if not candidatas:
        return 0
    candidatas = candidatas[:max(cfg.ia_maximo, 0)]

    if chamar is None:
        cfg_ia = cfg_ia if cfg_ia is not None else ler_config_adaptador(cfg.pasta_adaptador)
        if not ia_ligada(cfg, cfg_ia):
            return 0
        chamar = lambda p: _chamar_gemini(p, cfg_ia, cfg.pasta_adaptador)  # noqa: E731

    lista = "\n".join(f"{i}. {v['titulo']}" for i, v in enumerate(candidatas))
    try:
        aprovadas = interpretar_resposta(chamar(PROMPT.format(titulos=lista)), len(candidatas))
    except Exception as e:  # noqa: BLE001
        avisar(f"      [aviso] A IA nao respondeu ({str(e)[:100]}). Seguindo so com as palavras-chave.")
        return 0

    for i, perfil in aprovadas.items():
        v = candidatas[i]
        v["nota"] = nota_minima
        v["curriculo"] = perfil
        v["termos"] = "IA: título compatível com o perfil" + (f" ({v['termos']})" if v.get("termos") else "")
        v["origem"] = "IA"
    return len(aprovadas)

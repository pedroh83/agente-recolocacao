# -*- coding: utf-8 -*-
"""
Extracao das vagas a partir do HTML dos e-mails de alerta do LinkedIn.

Cada alerta traz varias vagas. Para cada link de vaga, o titulo e o texto
do link, e empresa/local vem logo depois, no formato:
    Empresa (ponto medio) Cidade, Estado, Pais
"""
from __future__ import annotations

import hashlib
import re
from html import unescape
from typing import Dict, List, Tuple

from texto import limpar_html, normalizar

# Link de vaga do LinkedIn, com ou sem o "/comm/" que aparece nos e-mails
PADRAO_LINKEDIN = re.compile(r"linkedin\.com/(?:comm/)?jobs/view/(\d+)", re.I)

# Captura <a href="...">texto</a>
PADRAO_ANCORA = re.compile(r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", re.I | re.S)

# Textos de botao que tambem apontam para a vaga, mas nao sao o titulo dela
TEXTOS_DE_BOTAO = {
    "ver vaga", "candidatar-se", "ver todas as vagas", "ver mais vagas", "candidate-se",
    "ver detalhes", "acessar", "clique aqui", "ver mais", "see job", "apply", "view job",
}

# Quanto texto olhar depois do titulo para achar empresa e local
TAMANHO_CONTEXTO = 900


def id_da_vaga(url: str, titulo: str = "", empresa: str = "") -> str:
    """ID estavel da vaga, para nao repeti-la nos proximos dias."""
    m = PADRAO_LINKEDIN.search(url or "")
    if m:
        return "li-" + m.group(1)
    base = normalizar(f"{titulo}|{empresa}")
    return "hash-" + hashlib.md5(base.encode()).hexdigest()[:12]


def link_limpo(url: str) -> str:
    """
    Endereco curto da vaga, sem os parametros de rastreamento do e-mail.

    Os links dos alertas trazem tokens pessoais (trackingId, otpToken...).
    O endereco curto abre a mesma vaga e pode ser guardado ou compartilhado
    sem expor esses dados.
    """
    m = PADRAO_LINKEDIN.search(url or "")
    return f"https://www.linkedin.com/jobs/view/{m.group(1)}/" if m else (url or "")


def extrair_vagas(html: str, assunto: str = "") -> List[Dict[str, str]]:
    """Percorre os links do e-mail e devolve os que apontam para uma vaga.

    O contexto de cada vaga e recortado ate o inicio da vaga seguinte, para
    que o texto de uma nao contamine a pontuacao da outra.
    """
    if not html:
        return []

    candidatas = []
    for m in PADRAO_ANCORA.finditer(html):
        url = unescape(m.group(1)).strip()
        if PADRAO_LINKEDIN.search(url):
            candidatas.append((m.start(), m.end(), url, m.group(2)))

    vagas = []
    for i, (_inicio, fim, url, miolo) in enumerate(candidatas):
        titulo, dentro_do_link = separar_titulo(miolo)
        # Links de imagem ou de botao vem vazios ou genericos: descarta
        if len(titulo) < 4 or normalizar(titulo) in TEXTOS_DE_BOTAO:
            continue

        limite = candidatas[i + 1][0] if i + 1 < len(candidatas) else len(html)
        depois = limpar_html(html[fim:min(limite, fim + TAMANHO_CONTEXTO)])
        contexto = f"{dentro_do_link} {depois}".strip()
        empresa, local = deduzir_empresa_e_local(contexto, titulo)

        vagas.append({
            "titulo": titulo,
            "empresa": empresa,
            "local": local,
            "plataforma": "LinkedIn",
            "link": link_limpo(url),
            "contexto": contexto[:300],
            "assunto_email": assunto,
        })
    return vagas


def separar_titulo(miolo: str) -> Tuple[str, str]:
    """Separa o titulo do resto quando o link envolve o cartao inteiro da vaga.

    Em parte dos alertas, o link abrange titulo, empresa e local, cada um num
    bloco HTML proprio. Sem separar, o titulo vira
    "Analista de Processos Empresa X · Salvador, BA" e a empresa fica vazia.
    Devolve (titulo, texto restante de dentro do link).
    """
    texto = limpar_html(miolo)
    if "\u00b7" not in texto:
        return texto, ""
    pedacos = [p for p in (limpar_html(x) for x in re.split(r"<[^>]+>", miolo or "")) if p]
    if len(pedacos) >= 2 and "\u00b7" not in pedacos[0] and len(pedacos[0]) >= 4:
        return pedacos[0], " ".join(pedacos[1:])
    return texto, ""


def deduzir_empresa_e_local(contexto: str, titulo: str) -> Tuple[str, str]:
    """Separa empresa e local a partir do texto que vem depois do titulo.

    Ex.: "Empresa Exemplo · Salvador, Bahia, Brasil (Hibrido)"
         -> ("Empresa Exemplo", "Salvador, Bahia (Hibrido)")
    """
    empresa, local = "", ""
    resto = contexto
    if titulo and titulo in resto:
        resto = resto.split(titulo, 1)[1]
    resto = resto.strip(" -|,")

    # Caminho principal: separador de ponto medio usado pelo LinkedIn
    for sep in ("·", "•", "|"):
        if sep in resto:
            partes = [p.strip(" -,") for p in resto.split(sep)]
            empresa = partes[0]
            if len(partes) > 1:
                local = partes[1]
            break
    else:
        # Sem separador: procura "Cidade, UF" ou uma palavra de modalidade
        m = re.search(r"([A-ZÀ-Ú][\wÀ-ú\s]{2,30},\s?[A-ZÀ-Ú][\wÀ-ú]+)", resto)
        if m:
            local = m.group(1).strip()
            empresa = resto[:m.start()].strip(" -|,")
        else:
            for termo in ("Remoto", "Home Office", "Híbrido", "Presencial"):
                idx = resto.lower().find(termo.lower())
                if idx != -1:
                    local, empresa = termo, resto[:idx].strip(" -|,")
                    break
            else:
                empresa = resto[:60]

    # Tira do local o salario e os selos que o LinkedIn poe depois dele
    local = re.split(r"\s+(?:R\$|US\$|\u20ac)", local)[0]
    if ")" in local:
        local = local[:local.index(")") + 1]

    # Corta o local em "Cidade, Estado": o pais nao acrescenta nada
    if local:
        pedacos = [p.strip() for p in local.split(",")]
        if len(pedacos) >= 3:
            local = f"{pedacos[0]}, {pedacos[1]}"
        # modalidade entre parenteses, ex.: "(Hibrido)"
        m = re.search(r"\(([^)]{3,12})\)", resto)
        if m and m.group(1).lower() not in local.lower():
            local = f"{local} ({m.group(1)})"

    empresa = re.sub(r"\s+", " ", empresa).strip(" -|,")[:50]
    local = re.sub(r"\s+", " ", local).strip(" -|,")[:40]
    return empresa, local

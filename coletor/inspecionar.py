#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
INSPETOR DE E-MAILS
===================
Mostra o que tem DENTRO de cada e-mail: assunto, quantos links e quantas
vagas o agente reconhece. Serve para descobrir por que ele extraiu 0 vagas.

    python inspecionar.py

Opcional - salva o HTML de um e-mail num arquivo para analise:

    python inspecionar.py --salvar 3

Opcional - salva o e-mail ja anonimizado, pronto para virar exemplo dos testes:

    python inspecionar.py --salvar 3 --anonimizar
"""
from __future__ import annotations

import argparse
import email
import re
import sys
from email.utils import parsedate_to_datetime

from configuracao import ErroConfiguracao, carregar_config
from email_gmail import ErroEmail, assunto_de, conectar, html_de, selecionar_pasta
from extracao import PADRAO_ANCORA, PADRAO_LINKEDIN, extrair_vagas


def cabecalho(txt: str) -> None:
    print("\n" + "=" * 66)
    print("  " + txt)
    print("=" * 66)


def anonimizar(html: str) -> str:
    """Tira do HTML os dados pessoais (enderecos de e-mail e tokens dos links)."""
    html = re.sub(r"[\w.+-]+@[\w-]+\.[\w.]+", "voce@exemplo.com", html)
    # Mantem so o numero da vaga nos links do LinkedIn; os parametros trazem tokens pessoais
    html = re.sub(r"(https?://[^\"'\s>]*linkedin\.com/(?:comm/)?jobs/view/\d+/?)[^\"'\s>]*", r"\1", html)
    html = re.sub(r"(https?://[^\"'\s>]*linkedin\.com/[^\"'\s>?]*)\?[^\"'\s>]*", r"\1", html)
    return html


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--salvar", type=int, default=None, help="numero do e-mail cujo HTML sera salvo")
    ap.add_argument("--anonimizar", action="store_true", help="tira dados pessoais do HTML salvo")
    args = ap.parse_args()

    try:
        cfg = carregar_config()
        conn = conectar(cfg)
    except (ErroConfiguracao, ErroEmail) as e:
        print(f"ERRO: {e}")
        return 1

    selecionar_pasta(conn, cfg.pasta, somente_leitura=True)
    _, dados = conn.search(None, "ALL")
    ids = dados[0].split()
    cabecalho(f"{len(ids)} E-MAIL(S) EM '{cfg.pasta}'")

    total_vagas = 0
    for n, uid in enumerate(ids, start=1):
        _, bruto = conn.fetch(uid, "(BODY.PEEK[])")
        if not bruto or not bruto[0] or not isinstance(bruto[0], tuple):
            continue
        msg = email.message_from_bytes(bruto[0][1])
        try:
            data = parsedate_to_datetime(msg.get("Date")).strftime("%d/%m %H:%M")
        except Exception:  # noqa: BLE001
            data = "?"

        html = html_de(msg)
        urls = [m.group(1) for m in PADRAO_ANCORA.finditer(html)]
        urls_li = [u for u in urls if "linkedin.com" in u.lower()]
        urls_vaga = [u for u in urls if PADRAO_LINKEDIN.search(u)]
        vagas = extrair_vagas(html)
        total_vagas += len(vagas)

        print(f"\n[{n}] {data} | {msg.get('From', '?')[:40]}")
        print(f"    Assunto: {assunto_de(msg)[:62]}")
        print(f"    HTML: {len(html)} caracteres | Links: {len(urls)} | do LinkedIn: {len(urls_li)} "
              f"| de vaga: {len(urls_vaga)} | vagas reconhecidas: {len(vagas)}")
        for v in vagas[:3]:
            print(f"        - {v['titulo'][:50]} | {v['empresa'][:25]} | {v['local'][:20]}")
        if not urls_vaga and urls_li:
            print("    Enderecos do LinkedIn neste e-mail (amostra):")
            vistos = []
            for u in urls_li:
                base = u.split("?")[0]
                if base not in vistos:
                    vistos.append(base)
                    print(f"        {base[:95]}")
                if len(vistos) >= 8:
                    break

        if args.salvar == n:
            nome = f"email_{n}{'_anonimo' if args.anonimizar else ''}.html"
            with open(nome, "w", encoding="utf-8") as f:
                f.write(anonimizar(html) if args.anonimizar else html)
            print(f"    >>> HTML salvo em {nome}")

    cabecalho("CONCLUSAO")
    if total_vagas > 0:
        print(f"  O agente reconhece {total_vagas} vaga(s) nesses e-mails.")
    else:
        print("  Nenhuma vaga reconhecida. Olhe a coluna 'Assunto' acima:")
        print("   - Se sao so e-mails de seguranca/confirmacao, esta tudo certo:")
        print("     os alertas de vaga ainda nao chegaram.")
        print("   - Se JA existe algum alerta de vaga na lista, o formato mudou.")
        print("     Rode 'python inspecionar.py --salvar N --anonimizar' (N = numero do")
        print("     e-mail de alerta) e use o arquivo para ajustar o codigo.")
    print()
    conn.logout()
    return 0


if __name__ == "__main__":
    sys.exit(main())

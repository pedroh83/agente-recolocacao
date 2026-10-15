#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
INSPETOR DE E-MAILS
===================
Mostra o que tem DENTRO de cada e-mail: assunto, quantos links, e uma
amostra dos enderecos do LinkedIn encontrados.

Serve para descobrir por que o agente extraiu 0 vagas.

    python inspecionar.py

Opcional - salva o HTML de um e-mail num arquivo para analise:

    python inspecionar.py --salvar 3
"""

import argparse
import email
import imaplib
import re
import sys
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime

try:
    import agente_vagas as ag
except ImportError:
    sys.exit("Nao achei o agente_vagas.py. Coloque este arquivo na mesma pasta.")

# O e-mail e a senha ficam no config.txt; e preciso carrega-lo antes de usar
if not ag.carregar_config():
    sys.exit("Nao achei o config.txt na pasta do agente.")
IMAP_SERVIDOR = ag.IMAP_SERVIDOR
IMAP_USUARIO = ag.IMAP_USUARIO
IMAP_SENHA_APP = ag.IMAP_SENHA_APP
PASTA_EMAIL = ag.PASTA_EMAIL
PADRAO_ANCORA = ag.PADRAO_ANCORA
PADRAO_LINKEDIN = ag.PADRAO_LINKEDIN


def cabecalho(txt):
    print("\n" + "=" * 66)
    print("  " + txt)
    print("=" * 66)


def pegar_html(msg):
    html = ""
    if msg.is_multipart():
        for parte in msg.walk():
            if parte.get_content_type() == "text/html":
                carga = parte.get_payload(decode=True)
                if carga:
                    html += carga.decode(
                        parte.get_content_charset() or "utf-8", errors="ignore")
    elif msg.get_content_type() == "text/html":
        carga = msg.get_payload(decode=True)
        if carga:
            html = carga.decode(
                msg.get_content_charset() or "utf-8", errors="ignore")
    return html


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--salvar", type=int, default=None,
                    help="numero do e-mail cujo HTML deve ser salvo em arquivo")
    args = ap.parse_args()

    senha = IMAP_SENHA_APP.replace(" ", "").strip()
    conn = imaplib.IMAP4_SSL(IMAP_SERVIDOR)
    conn.login(IMAP_USUARIO, senha)
    conn.select(f'"{PASTA_EMAIL}"' if PASTA_EMAIL != "INBOX" else "INBOX",
                readonly=True)

    _, dados = conn.search(None, "ALL")
    ids = dados[0].split()

    cabecalho(f"{len(ids)} E-MAIL(S) EM '{PASTA_EMAIL}'")

    total_vagas_geral = 0

    for n, uid in enumerate(ids, start=1):
        _, bruto = conn.fetch(uid, "(RFC822)")
        if not bruto or not bruto[0]:
            continue
        msg = email.message_from_bytes(bruto[0][1])

        try:
            assunto = str(make_header(decode_header(msg.get("Subject", ""))))
        except Exception:
            assunto = msg.get("Subject", "?")
        remetente = msg.get("From", "?")
        try:
            data = parsedate_to_datetime(msg.get("Date")).strftime("%d/%m %H:%M")
        except Exception:
            data = "?"

        html = pegar_html(msg)
        ancoras = PADRAO_ANCORA.findall(html)
        urls = [a[0] for a in ancoras]
        urls_li = [u for u in urls if "linkedin.com" in u.lower()]
        urls_vaga = [u for u in urls if PADRAO_LINKEDIN.search(u)]
        total_vagas_geral += len(urls_vaga)

        print(f"\n[{n}] {data} | {remetente[:40]}")
        print(f"    Assunto: {assunto[:62]}")
        print(f"    HTML: {len(html)} caracteres | Links: {len(urls)} "
              f"| do LinkedIn: {len(urls_li)} | de VAGA: {len(urls_vaga)}")

        if urls_vaga:
            print("    >>> ACHOU LINK DE VAGA <<<")
            for u in urls_vaga[:2]:
                print(f"        {u[:95]}")
        elif urls_li:
            print("    Enderecos do LinkedIn neste e-mail (amostra):")
            vistos = set()
            for u in urls_li:
                # corta os parametros para ficar legivel
                base = u.split("?")[0]
                if base in vistos:
                    continue
                vistos.add(base)
                print(f"        {base[:95]}")
                if len(vistos) >= 8:
                    break

        if args.salvar == n:
            nome = f"email_{n}.html"
            with open(nome, "w", encoding="utf-8") as f:
                f.write(html)
            print(f"    >>> HTML salvo em {nome}")

    cabecalho("CONCLUSAO")
    if total_vagas_geral > 0:
        print(f"  Foram encontrados {total_vagas_geral} link(s) de vaga no total.")
        print("  Se o agente ainda extraiu 0, o problema esta na leitura do")
        print("  titulo, nao do link. Me mande a saida acima.")
    else:
        print("  Nenhum link de vaga em nenhum e-mail.")
        print("  Olhe a coluna 'Assunto' acima:")
        print("   - Se sao so e-mails de seguranca/confirmacao, esta tudo certo.")
        print("     Os alertas de vaga ainda nao chegaram. Espere ate amanha.")
        print("   - Se JA existe algum alerta de vaga na lista, o formato mudou.")
        print("     Rode 'python inspecionar.py --salvar N' (N = numero do")
        print("     e-mail de alerta) e me mande o arquivo gerado.")
    print()
    conn.logout()


if __name__ == "__main__":
    main()

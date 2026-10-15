#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DIAGNOSTICO DO AGENTE DE VAGAS
==============================
Olha a conta do Gmail inteira e diz o que tem la dentro, sem filtro de data.
Serve para descobrir por que o agente nao encontrou e-mails.

Coloque este arquivo na mesma pasta do agente_vagas.py e rode:

    python diagnostico.py

Ele nao altera nada. So le e mostra.
"""

import datetime as dt
import email
import imaplib
import sys
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime

# Reaproveita a configuracao que voce ja preencheu no agente
try:
    import agente_vagas as ag
    ag.carregar_config()
    IMAP_SERVIDOR = ag.IMAP_SERVIDOR
    IMAP_USUARIO = ag.IMAP_USUARIO
    IMAP_SENHA_APP = ag.IMAP_SENHA_APP
    PASTA_EMAIL = ag.PASTA_EMAIL
except ImportError:
    sys.exit("Nao achei o agente_vagas.py. Coloque este arquivo na mesma pasta.")


def titulo(txt):
    print("\n" + "=" * 62)
    print("  " + txt)
    print("=" * 62)


def assunto_de(msg):
    try:
        return str(make_header(decode_header(msg.get("Subject", "(sem assunto)"))))
    except Exception:
        return msg.get("Subject", "(sem assunto)")


def data_de(msg):
    try:
        return parsedate_to_datetime(msg.get("Date")).strftime("%d/%m/%Y %H:%M")
    except Exception:
        return "?"


def mostrar_ultimos(conn, quantos=8):
    """Mostra os ultimos e-mails da pasta selecionada."""
    _, dados = conn.search(None, "ALL")
    ids = dados[0].split()
    if not ids:
        print("  (nenhum e-mail aqui)")
        return 0

    print(f"  Total: {len(ids)} e-mail(s). Mostrando os {min(quantos, len(ids))} mais recentes:\n")
    for uid in reversed(ids[-quantos:]):
        _, bruto = conn.fetch(uid, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
        if not bruto or not bruto[0]:
            continue
        msg = email.message_from_bytes(bruto[0][1])
        remetente = msg.get("From", "?")[:45]
        print(f"   {data_de(msg):<17} | {remetente}")
        print(f"   {'':<17} | {assunto_de(msg)[:60]}\n")
    return len(ids)


def main():
    senha = IMAP_SENHA_APP.replace(" ", "").strip()

    titulo("1. CONEXAO")
    print(f"  Conta: {IMAP_USUARIO}")
    print(f"  Senha de app: {len(senha)} caracteres "
          f"({'ok' if len(senha) == 16 else 'ATENCAO: deveria ter 16'})")

    try:
        conn = imaplib.IMAP4_SSL(IMAP_SERVIDOR)
        conn.login(IMAP_USUARIO, senha)
        print("  Login: OK")
    except Exception as e:
        sys.exit(f"  Login: FALHOU -> {e}")

    titulo("2. MARCADORES EXISTENTES NA CONTA")
    _, lista = conn.list()
    marcadores = []
    for item in lista:
        nome = item.decode(errors="ignore").split(' "/" ')[-1].strip('"')
        marcadores.append(nome)
        marca = "  <-- e este que o agente procura" if nome == PASTA_EMAIL else ""
        print(f"  - {nome}{marca}")

    if PASTA_EMAIL not in marcadores:
        print(f"\n  PROBLEMA: nao existe marcador chamado '{PASTA_EMAIL}'.")
        print("  Compare com a lista acima (maiuscula/minuscula e acento contam).")

    titulo(f"3. O QUE TEM NO MARCADOR '{PASTA_EMAIL}'")
    status, _ = conn.select(f'"{PASTA_EMAIL}"', readonly=True)
    total_vagas = 0
    if status == "OK":
        total_vagas = mostrar_ultimos(conn)
    else:
        print("  Nao consegui abrir esse marcador.")

    titulo("4. O QUE TEM NA CAIXA DE ENTRADA")
    conn.select("INBOX", readonly=True)
    total_inbox = mostrar_ultimos(conn, 5)

    titulo("5. PROCURANDO QUALQUER E-MAIL DO LINKEDIN NA CONTA")
    conn.select("INBOX", readonly=True)
    _, d1 = conn.search(None, '(FROM "linkedin")')
    achados_inbox = len(d1[0].split())
    print(f"  Na caixa de entrada: {achados_inbox} e-mail(s) do LinkedIn")

    achados_vagas = 0
    if status == "OK":
        conn.select(f'"{PASTA_EMAIL}"', readonly=True)
        _, d2 = conn.search(None, '(FROM "linkedin")')
        achados_vagas = len(d2[0].split())
        print(f"  No marcador '{PASTA_EMAIL}': {achados_vagas} e-mail(s) do LinkedIn")

    titulo("CONCLUSAO")
    if achados_vagas > 0:
        print("  Tem e-mail do LinkedIn no lugar certo.")
        print("  Se o agente achou 0, e porque sao mais antigos que a janela de dias.")
        print("  Rode assim para olhar mais para tras:")
        print("      python agente_vagas.py --debug --dias 30")
    elif achados_inbox > 0:
        print("  Chegou e-mail do LinkedIn, mas ficou na CAIXA DE ENTRADA.")
        print("  O filtro do Gmail nao esta aplicando o marcador.")
        print("  Conserte em: Configuracoes > Ver todas > Filtros e enderecos bloqueados")
        print(f"  Enquanto isso, selecione esses e-mails e aplique '{PASTA_EMAIL}' na mao.")
    elif total_inbox == 0 and total_vagas == 0:
        print("  A conta esta completamente vazia.")
        print("  Ou os alertas ainda nao foram enviados (o primeiro lote vem")
        print("  na manha seguinte a criacao), ou o LinkedIn esta mandando")
        print("  para outro endereco de e-mail.")
        print("  Confira em: LinkedIn > Vagas > Alertas de vaga")
    else:
        print("  Tem e-mail na conta, mas nenhum do LinkedIn.")
        print("  Provavelmente os alertas estao indo para outro endereco.")
        print("  Confira qual e-mail esta cadastrado no LinkedIn.")

    print()
    conn.logout()


if __name__ == "__main__":
    main()

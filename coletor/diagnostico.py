#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DIAGNOSTICO DO AGENTE DE VAGAS
==============================
Olha a conta do Gmail e diz o que tem la dentro, sem filtro de data.
Serve para descobrir por que o agente nao encontrou e-mails.

    python diagnostico.py

Nao altera nada: so le e mostra.
"""
from __future__ import annotations

import email
import sys
from email.utils import parsedate_to_datetime

from configuracao import ErroConfiguracao, carregar_config
from email_gmail import ErroEmail, assunto_de, conectar, listar_marcadores


def titulo(txt: str) -> None:
    print("\n" + "=" * 62)
    print("  " + txt)
    print("=" * 62)


def data_de(msg) -> str:
    try:
        return parsedate_to_datetime(msg.get("Date")).strftime("%d/%m/%Y %H:%M")
    except Exception:  # noqa: BLE001
        return "?"


def mostrar_ultimos(conn, quantos: int = 8) -> int:
    _, dados = conn.search(None, "ALL")
    ids = dados[0].split()
    if not ids:
        print("  (nenhum e-mail aqui)")
        return 0
    print(f"  Total: {len(ids)} e-mail(s). Mostrando os {min(quantos, len(ids))} mais recentes:\n")
    for uid in reversed(ids[-quantos:]):
        _, bruto = conn.fetch(uid, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
        if not bruto or not bruto[0] or not isinstance(bruto[0], tuple):
            continue
        msg = email.message_from_bytes(bruto[0][1])
        print(f"   {data_de(msg):<17} | {msg.get('From', '?')[:45]}")
        print(f"   {'':<17} | {assunto_de(msg)[:60]}\n")
    return len(ids)


def contar_linkedin(conn) -> int:
    _, dados = conn.search(None, '(FROM "linkedin")')
    return len(dados[0].split())


def main() -> int:
    try:
        cfg = carregar_config()
    except ErroConfiguracao as e:
        print(f"ERRO: {e}")
        return 1

    titulo("1. CONEXAO")
    print(f"  Conta: {cfg.email}")
    tamanho = len(cfg.senha_sem_espacos)
    print(f"  Senha de app: {tamanho} caracteres ({'ok' if tamanho == 16 else 'ATENCAO: deveria ter 16'})")
    try:
        conn = conectar(cfg)
        print("  Login: OK")
    except ErroEmail as e:
        print(f"  Login: FALHOU\n{e}")
        return 1

    pasta = cfg.pasta or "INBOX"
    usa_marcador = pasta.upper() != "INBOX"

    titulo("2. MARCADORES EXISTENTES NA CONTA")
    marcadores = listar_marcadores(conn)
    for nome in marcadores:
        marca = "  <-- e este que o agente procura" if nome == pasta else ""
        print(f"  - {nome}{marca}")
    if usa_marcador and pasta not in marcadores:
        print(f"\n  PROBLEMA: nao existe marcador chamado '{pasta}'.")
        print("  Compare com a lista acima (maiuscula/minuscula e acento contam).")

    no_marcador = 0
    if usa_marcador:
        titulo(f"3. O QUE TEM NO MARCADOR '{pasta}'")
        if conn.select(f'"{pasta}"', readonly=True)[0] == "OK":
            mostrar_ultimos(conn)
            no_marcador = contar_linkedin(conn)
        else:
            print("  Nao consegui abrir esse marcador.")

    titulo("4. O QUE TEM NA CAIXA DE ENTRADA")
    conn.select("INBOX", readonly=True)
    total_inbox = mostrar_ultimos(conn, 5)

    titulo("5. E-MAILS DO LINKEDIN NA CONTA")
    na_inbox = contar_linkedin(conn)
    print(f"  Na caixa de entrada: {na_inbox} e-mail(s) do LinkedIn")
    if usa_marcador:
        print(f"  No marcador '{pasta}': {no_marcador} e-mail(s) do LinkedIn")

    titulo("CONCLUSAO")
    no_lugar_certo = no_marcador if usa_marcador else na_inbox
    if no_lugar_certo > 0:
        print("  Tem e-mail do LinkedIn no lugar certo.")
        print("  Se o agente achou 0, e porque sao mais antigos que a janela de dias.")
        print("  Rode assim para olhar mais para tras:")
        print("      python agente_vagas.py --debug --dias 30")
    elif usa_marcador and na_inbox > 0:
        print("  Chegou e-mail do LinkedIn, mas ficou na CAIXA DE ENTRADA.")
        print("  Troque no config.txt para 'pasta = INBOX' ou corrija o filtro do Gmail.")
    elif total_inbox == 0:
        print("  A conta esta vazia. Ou os alertas ainda nao foram enviados (o primeiro lote")
        print("  chega na manha seguinte a criacao), ou o LinkedIn esta mandando para outro")
        print("  endereco. Confira em: LinkedIn > Vagas > Alertas de vaga")
    else:
        print("  Tem e-mail na conta, mas nenhum do LinkedIn.")
        print("  Provavelmente os alertas estao indo para outro endereco.")
        print("  Confira qual e-mail esta cadastrado no LinkedIn.")
    print()
    conn.logout()
    return 0


if __name__ == "__main__":
    sys.exit(main())

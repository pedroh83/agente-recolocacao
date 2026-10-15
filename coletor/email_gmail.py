# -*- coding: utf-8 -*-
"""
Conexao com o Gmail por IMAP (senha de app) e leitura dos e-mails.

Usado pelo agente e pelas ferramentas de suporte (diagnostico e inspecionar).
"""
from __future__ import annotations

import datetime as dt
import email
import imaplib
import socket
from email.header import decode_header, make_header
from email.message import Message
from typing import Callable, Dict, List, Optional

from configuracao import Config

TEMPO_LIMITE = 45  # segundos esperando o Gmail antes de desistir


class ErroEmail(Exception):
    """Falha de conexao ou login, com mensagem amigavel."""


def conectar(cfg: Config) -> imaplib.IMAP4_SSL:
    if not cfg.email or not cfg.senha:
        raise ErroEmail("O config.txt esta sem 'email' ou sem 'senha'. Abra o arquivo e confira.")
    try:
        try:
            conn = imaplib.IMAP4_SSL(cfg.servidor, timeout=TEMPO_LIMITE)
        except TypeError:  # Python antigo nao aceita timeout aqui
            socket.setdefaulttimeout(TEMPO_LIMITE)
            conn = imaplib.IMAP4_SSL(cfg.servidor)
    except Exception as e:  # noqa: BLE001
        raise ErroEmail(f"Nao consegui conectar ao Gmail: {e}\n"
                        "Pode ser internet instavel. Tente rodar de novo.")
    try:
        conn.login(cfg.email, cfg.senha_sem_espacos)
    except imaplib.IMAP4.error:
        raise ErroEmail(
            "Login recusado pelo Gmail. Causas mais comuns:\n"
            "  1. A senha de app foi digitada errada (sao 16 letras)\n"
            "  2. A verificacao em duas etapas nao esta ativada na conta\n"
            "  3. Foi usada a senha normal da conta em vez da senha de app\n"
            "Gere uma nova em: myaccount.google.com/apppasswords")
    return conn


def selecionar_pasta(conn: imaplib.IMAP4_SSL, pasta: str, somente_leitura: bool = True) -> bool:
    """Abre a pasta/marcador. Se nao existir, abre a caixa de entrada e devolve False."""
    nome = "INBOX" if not pasta or pasta.upper() == "INBOX" else f'"{pasta}"'
    status, _ = conn.select(nome, readonly=somente_leitura)
    if status == "OK":
        return True
    conn.select("INBOX", readonly=somente_leitura)
    return False


def listar_marcadores(conn: imaplib.IMAP4_SSL) -> List[str]:
    _, lista = conn.list()
    return [item.decode(errors="ignore").split(' "/" ')[-1].strip('"') for item in lista or []]


def assunto_de(msg: Message) -> str:
    try:
        return str(make_header(decode_header(msg.get("Subject", ""))))
    except Exception:  # noqa: BLE001
        return msg.get("Subject", "")


def html_de(msg: Message) -> str:
    """Junta todas as partes HTML do e-mail."""
    partes = msg.walk() if msg.is_multipart() else [msg]
    html = ""
    for parte in partes:
        if parte.get_content_type() == "text/html":
            carga = parte.get_payload(decode=True)
            if carga:
                html += carga.decode(parte.get_content_charset() or "utf-8", errors="ignore")
    return html


def ler_emails(cfg: Config, pasta: str, dias: int,
               progresso: Optional[Callable[[str], None]] = print) -> List[Dict[str, str]]:
    """Le os e-mails dos ultimos `dias` dias. Devolve [{'assunto', 'html'}]."""
    conn = conectar(cfg)
    try:
        if not selecionar_pasta(conn, pasta, somente_leitura=True) and progresso:
            progresso(f"[aviso] Pasta/marcador '{pasta}' nao encontrado. Usando a caixa de entrada.")

        desde = (dt.date.today() - dt.timedelta(days=dias)).strftime("%d-%b-%Y")
        _, dados = conn.search(None, f'(SINCE "{desde}")')
        ids = dados[0].split()

        emails, falhas = [], 0
        for i, uid in enumerate(ids, start=1):
            if progresso and len(ids) > 10 and i % 10 == 0:
                progresso(f"      ... {i} de {len(ids)}")
            try:
                _, bruto = conn.fetch(uid, "(BODY.PEEK[])")  # PEEK: nao marca como lido
            except Exception:  # noqa: BLE001
                falhas += 1
                continue
            if not bruto or not bruto[0] or not isinstance(bruto[0], tuple):
                continue
            msg = email.message_from_bytes(bruto[0][1])
            emails.append({"assunto": assunto_de(msg), "html": html_de(msg)})

        if falhas and progresso:
            progresso(f"      ({falhas} e-mail(s) nao puderam ser lidos e foram ignorados)")
        return emails
    finally:
        try:
            conn.close()
            conn.logout()
        except Exception:  # noqa: BLE001
            pass

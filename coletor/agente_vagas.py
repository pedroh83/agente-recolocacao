#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGENTE DE VAGAS - Pedro Henrique Simas
=======================================
Le os e-mails de alerta de vaga do LinkedIn e do Catho, remove as vagas
que ja apareceram em dias anteriores, pontua cada uma contra as palavras-chave
dos dois curriculos mestres (A = Qualidade/Processos, B = Requisitos/Tecnologia)
e gera uma planilha Excel com o link de candidatura de cada vaga.

MODOS DE LEITURA
----------------
  --modo imap     -> le de uma conta Gmail (senha de app) [PADRAO]
  --modo outlook  -> le do app Outlook classico instalado no Windows

USO
---
  python agente_vagas.py --debug        (primeira vez: mostra o que extraiu)
  python agente_vagas.py                (execucao normal)
"""

import argparse
import datetime as dt
import hashlib
import json
import os
import re
import sys
import unicodedata
from html import unescape

VERSAO = "1.4"

# ---------------------------------------------------------------------------
# CONFIGURACAO - ajuste aqui
# ---------------------------------------------------------------------------

# Pasta (no e-mail) onde caem os alertas - sobrescrito pelo config.txt
PASTA_EMAIL = "INBOX"
PASTA_EMAIL_PADRAO = "__padrao__"
DIAS_PADRAO = -1

# Quantos dias de e-mail olhar para tras a cada execucao
DIAS_PARA_TRAS = 2

# Onde salvar os arquivos
DIR_SAIDA = os.path.join(os.path.expanduser("~"), "AgenteVagas")
ARQ_HISTORICO = os.path.join(DIR_SAIDA, "vagas_vistas.json")

# Configuracao do Gmail (modo imap)
# Estes valores sao lidos do arquivo config.txt, na mesma pasta.
# Assim, quando o agente for atualizado, sua configuracao nao se perde.
IMAP_SERVIDOR = "imap.gmail.com"
IMAP_USUARIO = ""
IMAP_SENHA_APP = ""


def carregar_config():
    """Le config.txt (formato chave = valor) e aplica as configuracoes."""
    global IMAP_USUARIO, IMAP_SENHA_APP, PASTA_EMAIL, NOTA_MINIMA, DIAS_PARA_TRAS

    caminho = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.txt")
    if not os.path.exists(caminho):
        return False

    with open(caminho, encoding="utf-8") as f:
        for linha in f:
            linha = linha.strip()
            if not linha or linha.startswith("#") or "=" not in linha:
                continue
            chave, valor = linha.split("=", 1)
            chave, valor = chave.strip().lower(), valor.strip()
            if chave == "email":
                IMAP_USUARIO = valor
            elif chave == "senha":
                IMAP_SENHA_APP = valor
            elif chave == "pasta":
                PASTA_EMAIL = valor
            elif chave == "nota_minima":
                try:
                    NOTA_MINIMA = int(valor)
                except ValueError:
                    pass
            elif chave == "dias":
                try:
                    DIAS_PARA_TRAS = int(valor)
                except ValueError:
                    pass
    return True

# Nota minima para a vaga entrar no relatorio (0 = traz tudo)
NOTA_MINIMA = 1

# ---------------------------------------------------------------------------
# PALAVRAS-CHAVE DOS CURRICULOS
# peso 3 = termo forte (quase define a vaga) | peso 1 = termo de apoio
# ---------------------------------------------------------------------------

CURRICULO_A = {  # Qualidade / Processos / ISO
    3: ["iso 9001", "iso 14001", "iso 45001", "sistema de gestao da qualidade",
        "sistema de gestao integrado", "sgq", "sgi",
        "auditoria interna", "auditor interno", "auditora interna",
        "analista da qualidade", "analista de qualidade",
        "coordenador da qualidade", "coordenadora da qualidade",
        "gestao da qualidade", "quality management",
        "mapeamento de processos", "melhoria continua", "controles internos",
        # variacoes de "analista de processos" que aparecem nas vagas reais
        "analista de processos", "analista processos", "gestao de processos",
        "automacao de processos", "consultor de processos", "consultora de processos",
        "processos e sistemas", "excelencia operacional"],
    1: ["qualidade", "auditoria", "processos", "processo", "pdca",
        "nao conformidade", "acao corretiva", "indicadores", "kpi", "bpmn",
        "procedimento operacional", "pop", "5s", "lean", "meio ambiente",
        "saude e seguranca", "sst", "certificacao", "conformidade", "compliance",
        "integridade", "riscos", "governanca", "padronizacao", "itil"],
}

CURRICULO_B = {  # Requisitos / Tecnologia / Negocios
    3: ["levantamento de requisitos", "analista de requisitos", "analista requisitos",
        # negocios aparece em varias formas nas vagas reais
        "analista de negocios", "analista de negocio", "analista negocios",
        "business analyst", "consultor de negocios", "consultora de negocios",
        "analise de negocios", "analista de implantacao",
        "product owner", "product ownership", "analista funcional",
        "analista de sistemas", "analista sistemas",
        # testes e qualidade de software
        "analista de testes", "analista de qa", "analista qa", "qa",
        "quality assurance", "quality engineer", "quality analyst",
        "engenheiro de qa", "engenheira de qa", "teste manual", "testes manuais",
        "homologacao", "analista de solucoes"],
    1: ["scrum", "agil", "kanban", "jira", "sprint", "backlog", "user story",
        "historia de usuario", "cerimonias", "uat", "caso de teste", "sql",
        "requisitos", "documentacao", "stakeholder", "product", "dados",
        "inteligencia artificial", "automacao", "sistemas", "ti",
        "negocio", "negocios", "implantacao", "analytics",
        "digitalizacao", "transformacao digital", "no code", "low code"],
}

# Sinais de que a vaga NAO serve (senioridade/area fora do alvo)
TERMOS_DESCARTE = ["estagio", "estagiario", "jovem aprendiz", "aprendiz",
                   "vendedor", "vendas externas", "motorista", "operador de caixa"]

# Bonus por localizacao
TERMOS_LOCAL = ["salvador", "bahia", "ba", "remoto", "home office", "hibrido"]

# ---------------------------------------------------------------------------
# UTILITARIOS
# ---------------------------------------------------------------------------


def normalizar(texto):
    """Minusculas, sem acento - para comparar palavras-chave com seguranca."""
    if not texto:
        return ""
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto.lower()


def limpar_html(fragmento):
    """Tira tags e normaliza espacos de um pedaco de HTML.

    O recorte de contexto pode cortar uma tag pela metade ('<table cellp'),
    entao removemos tambem qualquer '<' sem o '>' correspondente.
    """
    txt = fragmento or ""
    txt = re.sub(r"<[^>]*>", " ", txt)   # tags completas
    txt = re.sub(r"<[^>]*$", " ", txt)   # tag cortada no fim
    txt = re.sub(r"^[^<>]*>", " ", txt)  # tag cortada no inicio
    txt = unescape(txt)
    return re.sub(r"\s+", " ", txt).strip()


# ---------------------------------------------------------------------------
# EXTRACAO DAS VAGAS A PARTIR DO HTML DO E-MAIL
# ---------------------------------------------------------------------------

# Padroes de link de vaga de cada plataforma
PADRAO_LINKEDIN = re.compile(r"linkedin\.com/(?:comm/)?jobs/view/(\d+)", re.I)
PADRAO_CATHO = re.compile(r"catho\.com\.br/vagas?/[^\"'\s>]*?(\d{6,})", re.I)
PADRAO_CATHO_GENERICO = re.compile(r"catho\.com\.br/vagas?/[^\"'\s>]+", re.I)

# Captura <a href="...">texto</a>
PADRAO_ANCORA = re.compile(
    r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", re.I | re.S)


def identificar_plataforma(url):
    u = url.lower()
    if "linkedin.com" in u:
        return "LinkedIn"
    if "catho.com.br" in u or "catho." in u:
        return "Catho"
    return None


def extrair_id(url, plataforma, titulo, empresa):
    """ID estavel para nao repetir a vaga amanha."""
    m = PADRAO_LINKEDIN.search(url)
    if m:
        return "li-" + m.group(1)
    m = PADRAO_CATHO.search(url)
    if m:
        return "ct-" + m.group(1)
    base = normalizar(f"{plataforma}|{titulo}|{empresa}")
    return "hash-" + hashlib.md5(base.encode()).hexdigest()[:12]


def eh_link_de_vaga(url):
    if PADRAO_LINKEDIN.search(url):
        return True
    if PADRAO_CATHO_GENERICO.search(url):
        return True
    return False


def extrair_vagas(html, assunto=""):
    """Percorre as ancoras do e-mail e devolve as que apontam para uma vaga.

    O contexto de cada vaga e recortado ate o inicio da vaga seguinte, para
    que o texto de uma nao contamine a pontuacao da outra.
    """
    vagas = []
    if not html:
        return vagas

    # Guarda posicao inicial e final de cada ancora que aponta para uma vaga
    candidatas = []
    for m in PADRAO_ANCORA.finditer(html):
        url = unescape(m.group(1)).strip()
        if eh_link_de_vaga(url):
            candidatas.append((m.start(), m.end(), url, m.group(2)))

    for i, (ini, fim, url, miolo) in enumerate(candidatas):
        titulo = limpar_html(miolo)
        # Ancoras de imagem ou de botao vem vazias/genericas - descarta
        if not titulo or len(titulo) < 4:
            continue
        if normalizar(titulo) in ("ver vaga", "candidatar-se", "ver todas as vagas",
                                  "ver mais vagas", "candidate-se", "ver detalhes",
                                  "acessar", "clique aqui", "ver mais"):
            continue

        plataforma = identificar_plataforma(url) or "Outra"

        # Contexto = do fim desta ancora ate o inicio da proxima vaga (max 900)
        limite = candidatas[i + 1][0] if i + 1 < len(candidatas) else len(html)
        contexto = limpar_html(html[fim: min(limite, fim + 900)])
        empresa, local = deduzir_empresa_e_local(contexto, titulo)

        vagas.append({
            "titulo": titulo,
            "empresa": empresa,
            "local": local,
            "plataforma": plataforma,
            "link": url,
            "contexto": contexto[:300],
            "assunto_email": assunto,
        })
    return vagas


def deduzir_empresa_e_local(contexto, titulo):
    """Separa empresa e local a partir do texto que segue o titulo.

    O LinkedIn usa o formato:  Empresa (ponto medio) Cidade, Estado, Pais
    Ex: "Oficial Atakarejo - Salvador, Bahia, Brasil"
    """
    empresa, local = "", ""
    resto = contexto
    if titulo and titulo in resto:
        resto = resto.split(titulo, 1)[1]
    resto = resto.strip(" -|,")

    # Caminho principal: separador de ponto medio usado pelo LinkedIn
    for sep in ("\u00b7", "\u2022", "|"):
        if sep in resto:
            partes = [p.strip(" -,") for p in resto.split(sep)]
            empresa = partes[0]
            if len(partes) > 1:
                local = partes[1]
            break
    else:
        # Sem separador: procura "Cidade, UF" ou palavra de modalidade
        m = re.search(r"([A-Z\u00c0-\u00da][\w\u00c0-\u00fa\s]{2,30},\s?[A-Z\u00c0-\u00da][\w\u00c0-\u00fa]+)", resto)
        if m:
            local = m.group(1).strip()
            empresa = resto[: m.start()].strip(" -|,")
        else:
            for termo in ("Remoto", "Home Office", "H\u00edbrido", "Presencial"):
                idx = resto.lower().find(termo.lower())
                if idx != -1:
                    local, empresa = termo, resto[:idx].strip(" -|,")
                    break
            else:
                empresa = resto[:60]

    # Corta o local em "Cidade, Estado" - o pais nao acrescenta nada
    if local:
        pedacos = [p.strip() for p in local.split(",")]
        if len(pedacos) >= 3:
            local = f"{pedacos[0]}, {pedacos[1]}"
        # modalidade entre parenteses, ex: "(H\u00edbrido)"
        m = re.search(r"\(([^)]{3,12})\)", resto)
        if m and m.group(1).lower() not in local.lower():
            local = f"{local} ({m.group(1)})"

    empresa = re.sub(r"\s+", " ", empresa).strip(" -|,")[:50]
    local = re.sub(r"\s+", " ", local).strip(" -|,")[:40]
    return empresa, local


# ---------------------------------------------------------------------------
# PONTUACAO
# ---------------------------------------------------------------------------


# Cache dos padroes compilados - evita recompilar a cada vaga
_CACHE_TERMOS = {}


def casa_termo(texto_norm, termo):
    """Verifica se o termo aparece como palavra inteira.

    Sem isso, 'ti' casaria dentro de 'logistica' e 'ativos', e 'pop'
    dentro de 'popular'. O (?<![a-z0-9]) e o (?![a-z0-9]) exigem que
    nao haja letra ou numero colado antes nem depois.
    """
    padrao = _CACHE_TERMOS.get(termo)
    if padrao is None:
        padrao = re.compile(r"(?<![a-z0-9])" + re.escape(termo) + r"(?![a-z0-9])")
        _CACHE_TERMOS[termo] = padrao
    return bool(padrao.search(texto_norm))


def pontuar_contra(texto_norm, mapa):
    nota, achados = 0, []
    for peso, termos in mapa.items():
        for termo in termos:
            if casa_termo(texto_norm, termo):
                nota += peso
                achados.append(termo)
    return nota, achados


def avaliar(vaga):
    texto = normalizar(f"{vaga['titulo']} {vaga['empresa']} "
                       f"{vaga['local']} {vaga['contexto']}")

    # O descarte olha SO o titulo da vaga - o contexto pode citar outra coisa
    titulo_norm = normalizar(vaga["titulo"])
    for termo in TERMOS_DESCARTE:
        if casa_termo(titulo_norm, termo):
            vaga["descartada"] = f"contem '{termo}'"
            vaga["nota"] = 0
            vaga["curriculo"] = "-"
            vaga["termos"] = ""
            return vaga

    nota_a, achados_a = pontuar_contra(texto, CURRICULO_A)
    nota_b, achados_b = pontuar_contra(texto, CURRICULO_B)

    bonus = 2 if any(casa_termo(texto, t) for t in TERMOS_LOCAL) else 0

    if nota_b > nota_a:
        vaga["curriculo"], vaga["nota"], vaga["termos"] = "B", nota_b + bonus, ", ".join(achados_b[:6])
    elif nota_a > nota_b:
        vaga["curriculo"], vaga["nota"], vaga["termos"] = "A", nota_a + bonus, ", ".join(achados_a[:6])
    elif nota_a == 0:
        vaga["curriculo"], vaga["nota"], vaga["termos"] = "?", bonus, ""
    else:
        vaga["curriculo"] = "A ou B"
        vaga["nota"] = nota_a + bonus
        vaga["termos"] = ", ".join((achados_a + achados_b)[:6])

    vaga["descartada"] = ""
    return vaga


# ---------------------------------------------------------------------------
# LEITORES DE E-MAIL
# ---------------------------------------------------------------------------


def ler_outlook_desktop(pasta, dias):
    """Le do app Outlook classico via COM. Nao pede senha."""
    try:
        import win32com.client  # pywin32
    except ImportError:
        sys.exit("Falta a biblioteca pywin32. Rode:  pip install pywin32")

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    caixa = outlook.GetDefaultFolder(6)  # 6 = Caixa de Entrada

    alvo = caixa
    if pasta:
        encontrada = None
        for sub in caixa.Folders:
            if normalizar(sub.Name) == normalizar(pasta):
                encontrada = sub
                break
        if encontrada is None:
            for sub in outlook.Folders.Item(1).Folders:
                if normalizar(sub.Name) == normalizar(pasta):
                    encontrada = sub
                    break
        if encontrada is None:
            print(f"[aviso] Pasta '{pasta}' nao encontrada. Usando a Caixa de Entrada.")
        else:
            alvo = encontrada

    itens = alvo.Items
    itens.Sort("[ReceivedTime]", True)

    corte = dt.datetime.now() - dt.timedelta(days=dias)
    emails = []
    for item in itens:
        try:
            recebido = item.ReceivedTime
            recebido = dt.datetime(recebido.year, recebido.month, recebido.day,
                                   recebido.hour, recebido.minute)
            if recebido < corte:
                break
            emails.append({"assunto": item.Subject or "",
                           "html": item.HTMLBody or "",
                           "data": recebido})
        except Exception:
            continue
    return emails


def ler_imap(pasta, dias):
    """Le de uma conta Gmail usando senha de app."""
    import email
    import imaplib
    from email.header import decode_header, make_header

    if not IMAP_USUARIO or not IMAP_SENHA_APP:
        sys.exit("O config.txt existe mas esta faltando 'email' ou 'senha'.\n"
                 "Abra C:\\AgenteVagas\\config.txt no Bloco de Notas e confira.")

    # O Google mostra a senha com espacos; aqui eles nao servem
    senha = IMAP_SENHA_APP.replace(" ", "").strip()

    # timeout evita que o programa fique travado para sempre esperando o Gmail
    try:
        try:
            conn = imaplib.IMAP4_SSL(IMAP_SERVIDOR, timeout=45)
        except TypeError:
            # Python antigo nao aceita timeout aqui
            import socket
            socket.setdefaulttimeout(45)
            conn = imaplib.IMAP4_SSL(IMAP_SERVIDOR)
    except Exception as e:
        sys.exit(f"Nao consegui conectar ao Gmail: {e}\n"
                 "Pode ser internet instavel. Tente rodar de novo.")

    try:
        conn.login(IMAP_USUARIO, senha)
    except imaplib.IMAP4.error:
        sys.exit(
            "Login recusado pelo Gmail. Causas mais comuns:\n"
            "  1. A senha de app foi digitada errada (sao 16 letras, sem espaco)\n"
            "  2. A verificacao em duas etapas nao esta ativada na conta\n"
            "  3. Voce usou a senha normal da conta em vez da senha de app\n"
            "Gere uma nova em: myaccount.google.com/apppasswords")

    # No Gmail, 'pasta' e na verdade um marcador (label)
    status, _ = conn.select(f'"{pasta}"' if pasta else "INBOX")
    if status != "OK":
        print(f"[aviso] Marcador '{pasta}' nao encontrado nesta conta.")
        print("        Marcadores disponiveis:")
        _, lista = conn.list()
        for item in lista:
            nome = item.decode(errors="ignore").split(' "/" ')[-1].strip('"')
            if not nome.startswith("[Gmail]"):
                print(f"          - {nome}")
        print("        Usando a Caixa de Entrada por enquanto.\n")
        conn.select("INBOX")

    desde = (dt.date.today() - dt.timedelta(days=dias)).strftime("%d-%b-%Y")
    _, dados = conn.search(None, f'(SINCE "{desde}")')
    ids = dados[0].split()

    emails = []
    total = len(ids)
    falhas = 0
    for i, uid in enumerate(ids, start=1):
        if total > 10 and i % 10 == 0:
            print(f"      ... {i} de {total}")
        try:
            _, bruto = conn.fetch(uid, "(RFC822)")
        except Exception:
            falhas += 1
            continue
        if not bruto or not bruto[0]:
            continue
        msg = email.message_from_bytes(bruto[0][1])
        assunto = str(make_header(decode_header(msg.get("Subject", ""))))

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

        emails.append({"assunto": assunto, "html": html, "data": dt.datetime.now()})

    if falhas:
        print(f"      ({falhas} e-mail(s) nao puderam ser lidos - ignorados)")
    try:
        conn.close()
        conn.logout()
    except Exception:
        pass
    return emails


# ---------------------------------------------------------------------------
# HISTORICO
# ---------------------------------------------------------------------------


def carregar_historico():
    if os.path.exists(ARQ_HISTORICO):
        try:
            with open(ARQ_HISTORICO, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def salvar_historico(hist):
    os.makedirs(DIR_SAIDA, exist_ok=True)
    with open(ARQ_HISTORICO, "w", encoding="utf-8") as f:
        json.dump(hist, f, ensure_ascii=False, indent=1)


# ---------------------------------------------------------------------------
# RELATORIO EXCEL
# ---------------------------------------------------------------------------


def gerar_excel(vagas, caminho):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Vagas do dia"

    cabecalho = ["Nota", "Currículo", "Vaga", "Empresa", "Local",
                 "Plataforma", "Link para candidatura", "Termos que bateram", "Enviado?"]
    ws.append(cabecalho)

    azul = PatternFill("solid", fgColor="1F3864")
    for col in range(1, len(cabecalho) + 1):
        c = ws.cell(row=1, column=col)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = azul
        c.alignment = Alignment(horizontal="center", vertical="center")

    verde = PatternFill("solid", fgColor="E2EFDA")
    amarelo = PatternFill("solid", fgColor="FFF2CC")

    for v in vagas:
        linha = [v["nota"], v["curriculo"], v["titulo"], v["empresa"], v["local"],
                 v["plataforma"], "Candidatar-se", v["termos"], ""]
        ws.append(linha)
        r = ws.max_row

        celula_link = ws.cell(row=r, column=7)
        celula_link.hyperlink = v["link"]
        celula_link.font = Font(color="0563C1", underline="single")

        if v["nota"] >= 6:
            for col in range(1, len(cabecalho) + 1):
                ws.cell(row=r, column=col).fill = verde
        elif v["nota"] >= 3:
            for col in range(1, len(cabecalho) + 1):
                ws.cell(row=r, column=col).fill = amarelo

    larguras = [7, 11, 46, 30, 24, 12, 22, 40, 11]
    for i, w in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cabecalho))}{ws.max_row}"

    wb.save(caminho)


# ---------------------------------------------------------------------------
# PRINCIPAL
# ---------------------------------------------------------------------------


def main():
    # Escreve a saida imediatamente, sem acumular em memoria.
    # Sem isso, se o programa travar, o registro.txt fica vazio.
    try:
        sys.stdout.reconfigure(line_buffering=True)
        sys.stderr.reconfigure(line_buffering=True)
    except AttributeError:
        pass

    ap = argparse.ArgumentParser(description="Agente de vagas - LinkedIn e Catho")
    ap.add_argument("--modo", choices=["imap", "outlook"], default="imap")
    ap.add_argument("--pasta", default=PASTA_EMAIL_PADRAO)
    ap.add_argument("--dias", type=int, default=DIAS_PADRAO)
    ap.add_argument("--debug", action="store_true")
    ap.add_argument("--refazer", action="store_true",
                    help="ignora o historico e mostra todas as vagas de novo")
    args = ap.parse_args()

    if not carregar_config():
        sys.exit("Nao achei o arquivo config.txt na pasta do agente.\n"
                 "Ele deve ficar em C:\\AgenteVagas\\config.txt")
    if args.pasta == PASTA_EMAIL_PADRAO:
        args.pasta = PASTA_EMAIL
    if args.dias == DIAS_PADRAO:
        args.dias = DIAS_PARA_TRAS

    print(f"=== Agente de Vagas v{VERSAO} ===")
    print(f"[1/5] Lendo e-mails (modo {args.modo}, pasta '{args.pasta}')...")
    if args.modo == "outlook":
        emails = ler_outlook_desktop(args.pasta, args.dias)
    else:
        emails = ler_imap(args.pasta, args.dias)
    print(f"      {len(emails)} e-mail(s) encontrado(s).")

    print("[2/5] Extraindo vagas...")
    brutas = []
    for em in emails:
        brutas.extend(extrair_vagas(em["html"], em["assunto"]))
    print(f"      {len(brutas)} link(s) de vaga.")

    print("[3/5] Removendo repetidas e vagas ja vistas...")
    hist = carregar_historico()
    if args.refazer:
        print(f"      (--refazer: ignorando as {len(hist)} vaga(s) do historico)")
        hist = {}
    hoje = dt.date.today().isoformat()
    vistas_agora, novas = set(), []

    for v in brutas:
        vid = extrair_id(v["link"], v["plataforma"], v["titulo"], v["empresa"])
        if vid in vistas_agora:
            continue
        vistas_agora.add(vid)
        if vid in hist:
            continue
        v["id"] = vid
        novas.append(v)
    print(f"      {len(novas)} vaga(s) nova(s).")

    print("[4/5] Pontuando contra os curriculos A e B...")
    avaliadas = [avaliar(v) for v in novas]
    finais = [v for v in avaliadas
              if not v["descartada"] and v["nota"] >= NOTA_MINIMA]
    finais.sort(key=lambda x: x["nota"], reverse=True)
    print(f"      {len(finais)} vaga(s) no relatorio.")

    if args.debug:
        print("\n--- DEBUG ---")
        for v in avaliadas:
            marca = "x" if v["descartada"] else " "
            print(f"[{marca}] {v['nota']:>3} {v['curriculo']:<6} | "
                  f"{v['titulo'][:45]:<45} | {v['empresa'][:22]:<22} | {v['plataforma']}")
        print("--- fim ---\n")

    print("[5/5] Gerando planilha...")
    os.makedirs(DIR_SAIDA, exist_ok=True)
    caminho = os.path.join(DIR_SAIDA, f"vagas_{hoje}.xlsx")

    # Nunca sobrescrever um relatorio existente com um vazio.
    # Sem isso, uma segunda execucao no mesmo dia apagaria o relatorio bom.
    if not finais:
        print("      Nenhuma vaga nova hoje - nada a salvar.")
        if os.path.exists(caminho):
            print(f"      A planilha de hoje foi mantida: {os.path.basename(caminho)}")
        print(f"\nHistorico: {len(hist)} vaga(s) acumulada(s).")
        return

    try:
        gerar_excel(finais, caminho)
    except PermissionError:
        # A planilha esta aberta no Excel - salva com outro nome
        alternativo = os.path.join(
            DIR_SAIDA, f"vagas_{hoje}_{dt.datetime.now():%H%M}.xlsx")
        print(f"      A planilha {os.path.basename(caminho)} esta aberta no Excel.")
        print(f"      Salvando com outro nome: {os.path.basename(alternativo)}")
        gerar_excel(finais, alternativo)
        caminho = alternativo

    for v in finais:
        hist[v["id"]] = {"data": hoje, "titulo": v["titulo"],
                         "empresa": v["empresa"], "link": v["link"]}
    salvar_historico(hist)

    print(f"\nPronto: {caminho}")
    print(f"Historico: {len(hist)} vaga(s) acumulada(s).")

    if os.name == "nt" and finais:
        try:
            os.startfile(caminho)
        except Exception:
            pass


if __name__ == "__main__":
    main()
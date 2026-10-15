#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGENTE DE VAGAS
===============
Le os e-mails de alerta de vaga do LinkedIn, descarta as vagas que ja
apareceram, pontua cada uma contra as palavras-chave dos dois curriculos
(A = Qualidade/Processos, B = Requisitos/Tecnologia) e gera:

  - vagas_AAAA-MM-DD.xlsx       -> as vagas novas do dia, com o curriculo indicado
  - controle_candidaturas.xlsx  -> todas as vagas, envios, adaptacoes e indicadores

USO
  python agente_vagas.py                  execucao normal
  python agente_vagas.py --debug          mostra a nota de cada vaga
  python agente_vagas.py --refazer        ignora o historico e mostra tudo de novo
  python agente_vagas.py --dias 10        olha mais dias para tras (ex.: depois de uma viagem)
  python agente_vagas.py --sem-ia         nao consulta a IA nesta execucao
  python agente_vagas.py --so-controle    so atualiza o controle de candidaturas
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys
from typing import Dict, List, Optional

from arquivamento import NOME_PASTA, arrumar_pasta
from configuracao import Config, ErroConfiguracao, carregar_config, carregar_palavras_chave
from controle import NOME_ARQUIVO as NOME_CONTROLE, atualizar_controle
from email_gmail import ErroEmail, ler_emails
from extracao import extrair_vagas, id_da_vaga
from ia_titulos import reavaliar_titulos
from planilha_dia import gerar_planilha_do_dia
from pontuacao import avaliar
from texto import normalizar

VERSAO = "1.5.3"


# ---------------------------------------------------------------------------
# Historico de vagas ja exibidas
# ---------------------------------------------------------------------------

def caminho_historico(cfg: Config) -> str:
    return os.path.join(cfg.dir_saida, "vagas_vistas.json")


def carregar_historico(cfg: Config) -> Dict[str, dict]:
    caminho = caminho_historico(cfg)
    if not os.path.exists(caminho):
        return {}
    try:
        with open(caminho, encoding="utf-8") as f:
            return json.load(f)
    except Exception:  # noqa: BLE001
        print("      [aviso] vagas_vistas.json ilegivel. Comecando um historico novo.")
        return {}


def salvar_historico(cfg: Config, historico: Dict[str, dict]) -> None:
    os.makedirs(cfg.dir_saida, exist_ok=True)
    caminho = caminho_historico(cfg)
    temporario = caminho + ".tmp"
    with open(temporario, "w", encoding="utf-8") as f:
        json.dump(historico, f, ensure_ascii=False, indent=1)
    os.replace(temporario, caminho)  # troca de uma vez: o historico nunca fica pela metade


# ---------------------------------------------------------------------------
# Etapas
# ---------------------------------------------------------------------------

DIAS_ANUNCIO_REPETIDO = 30


def chave_do_anuncio(titulo: str, empresa: str) -> str:
    """Mesmo cargo na mesma empresa. Vazio se faltar a empresa (nao da para comparar)."""
    titulo, empresa = " ".join(normalizar(titulo).split()), " ".join(normalizar(empresa).split())
    return f"{titulo}|{empresa}" if titulo and empresa else ""


def separar_novas(brutas: List[dict], historico: Dict[str, dict],
                  hoje: Optional[dt.date] = None) -> List[dict]:
    """
    Remove as vagas repetidas no mesmo lote e as que ja estao no historico.

    Alem do numero da vaga, compara cargo + empresa: o LinkedIn publica a mesma
    vaga varias vezes, com numeros diferentes (uma por cidade, por exemplo).
    Um anuncio igual a outro dos ultimos 30 dias conta como repetido.
    """
    hoje = hoje or dt.date.today()
    limite = (hoje - dt.timedelta(days=DIAS_ANUNCIO_REPETIDO)).isoformat()
    anuncios_recentes = {chave_do_anuncio(h.get("titulo", ""), h.get("empresa", ""))
                         for h in historico.values() if str(h.get("data", "")) >= limite}
    anuncios_recentes.discard("")

    vistas, novas = set(), []
    for v in brutas:
        ident = id_da_vaga(v["link"], v["titulo"], v["empresa"])
        anuncio = chave_do_anuncio(v["titulo"], v["empresa"])
        if ident in vistas or ident in historico or (anuncio and anuncio in anuncios_recentes):
            continue
        vistas.add(ident)
        if anuncio:
            anuncios_recentes.add(anuncio)
        v["id"] = ident
        novas.append(v)
    return novas


def salvar_planilha_do_dia(cfg: Config, finais: List[dict], hoje: str) -> str:
    caminho = os.path.join(cfg.dir_saida, f"vagas_{hoje}.xlsx")
    alternativo = os.path.join(cfg.dir_saida, f"vagas_{hoje}_{dt.datetime.now():%H%M%S}.xlsx")
    if os.path.exists(caminho):
        # Segunda execucao do dia: nunca sobrescreve a planilha da manha,
        # que pode ter "Enviado?" marcado e ainda nao levado ao controle
        print(f"      Ja existe uma planilha de hoje. As vagas novas vao para {os.path.basename(alternativo)}")
        caminho = alternativo
    try:
        gerar_planilha_do_dia(finais, caminho)
    except PermissionError:
        # A planilha do dia esta aberta no Excel: salva com outro nome
        print(f"      {os.path.basename(caminho)} esta aberta no Excel. "
              f"Salvando como {os.path.basename(alternativo)}")
        gerar_planilha_do_dia(finais, alternativo)
        caminho = alternativo
    return caminho


def registrar_no_historico(historico: Dict[str, dict], finais: List[dict], hoje: str) -> None:
    for v in finais:
        # Com --refazer, a vaga pode ja estar no historico: mantem a data em que apareceu pela primeira vez
        primeira_vez = historico.get(v["id"], {}).get("data") or hoje
        historico[v["id"]] = {
            "data": primeira_vez, "titulo": v["titulo"], "empresa": v["empresa"], "link": v["link"],
            "local": v["local"], "nota": v["nota"], "curriculo": v["curriculo"],
            "termos": v["termos"], "origem": v.get("origem", "Agente"),
        }


def atualizar_e_informar_controle(cfg: Config, historico: Dict[str, dict], palavras) -> bool:
    """Atualiza o controle e mostra o resumo. Devolve True se o controle foi gravado."""
    resumo = atualizar_controle(cfg.dir_saida, historico, cfg.pasta_adaptador, palavras)
    if resumo.get("erro"):
        print(f"      [aviso] Controle de candidaturas: {resumo['erro']}")
        return False
    print(f"      Controle: {resumo['linhas']} vaga(s), {resumo['enviadas']} candidatura(s) enviada(s).")
    if resumo["enviados_importados"]:
        print(f"      {resumo['enviados_importados']} envio(s) trazido(s) das planilhas do dia.")
    if resumo["adaptacoes_ligadas"] or resumo["adaptacoes_avulsas"]:
        print(f"      Adaptacoes de curriculo: {resumo['adaptacoes_ligadas']} ligada(s) a vagas do agente, "
              f"{resumo['adaptacoes_avulsas']} de vagas encontradas por fora.")
    return True


def arrumar_e_informar(cfg: Config, controle_ok: bool) -> None:
    movidos = arrumar_pasta(cfg.dir_saida, cfg.dias_para_arquivar, controle_atualizado=controle_ok)
    if movidos:
        print(f"      {len(movidos)} arquivo(s) que nao sao mais necessarios foram para {NOME_PASTA}:")
        for nome in movidos:
            print(f"        - {nome}")


def abrir_no_windows(caminho: str) -> None:
    if os.name == "nt":
        try:
            os.startfile(caminho)  # type: ignore[attr-defined]
        except Exception:  # noqa: BLE001
            pass


# ---------------------------------------------------------------------------
# Principal
# ---------------------------------------------------------------------------

def ler_argumentos(argv=None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Agente de vagas do LinkedIn")
    ap.add_argument("--pasta", default=None, help="pasta/marcador do Gmail (padrao: o do config.txt)")
    ap.add_argument("--dias", type=int, default=None, help="dias de e-mail para tras (padrao: config.txt)")
    ap.add_argument("--debug", action="store_true", help="mostra a nota de cada vaga")
    ap.add_argument("--refazer", action="store_true", help="ignora o historico e mostra tudo de novo")
    ap.add_argument("--sem-ia", action="store_true", help="nao consulta a IA nesta execucao")
    ap.add_argument("--so-controle", action="store_true", help="so atualiza o controle de candidaturas")
    return ap.parse_args(argv)


def executar(args: argparse.Namespace) -> int:
    cfg = carregar_config()
    palavras = carregar_palavras_chave()
    pasta = args.pasta or cfg.pasta
    dias = args.dias if args.dias is not None else cfg.dias
    hoje = dt.date.today().isoformat()

    print(f"=== Agente de Vagas v{VERSAO} ===")
    historico = carregar_historico(cfg)

    if args.so_controle:
        print("Atualizando so o controle de candidaturas...")
        arrumar_e_informar(cfg, atualizar_e_informar_controle(cfg, historico, palavras))
        return 0

    print(f"[1/6] Lendo e-mails (pasta '{pasta}', ultimos {dias} dia(s))...")
    emails = ler_emails(cfg, pasta, dias)
    print(f"      {len(emails)} e-mail(s) encontrado(s).")

    print("[2/6] Extraindo vagas...")
    brutas = [v for e in emails for v in extrair_vagas(e["html"], e["assunto"])]
    print(f"      {len(brutas)} link(s) de vaga.")

    print("[3/6] Removendo repetidas e vagas ja vistas...")
    if args.refazer:
        print(f"      (--refazer: ignorando as {len(historico)} vaga(s) do historico)")
    novas = separar_novas(brutas, {} if args.refazer else historico)
    print(f"      {len(novas)} vaga(s) nova(s).")

    print("[4/6] Pontuando contra os curriculos A e B...")
    avaliadas = [avaliar(v, palavras) for v in novas]
    if not args.sem_ia:
        resgatadas = reavaliar_titulos(avaliadas, cfg, cfg.nota_minima)
        if resgatadas:
            print(f"      A IA reconheceu {resgatadas} titulo(s) do seu perfil que as palavras-chave nao pegaram.")
    finais = sorted((v for v in avaliadas if not v["descartada"] and v["nota"] >= cfg.nota_minima),
                    key=lambda v: v["nota"], reverse=True)
    print(f"      {len(finais)} vaga(s) no relatorio.")

    if args.debug:
        print("\n--- DEBUG ---")
        for v in sorted(avaliadas, key=lambda v: v["nota"], reverse=True):
            marca = "x" if v["descartada"] else ("IA" if v.get("origem") == "IA" else " ")
            print(f"[{marca:>2}] {v['nota']:>3} {v['curriculo']:<6} | "
                  f"{v['titulo'][:45]:<45} | {v['empresa'][:22]:<22}")
        print("--- fim ---\n")

    print("[5/6] Gerando a planilha do dia...")
    os.makedirs(cfg.dir_saida, exist_ok=True)
    caminho = None
    if finais:
        # Uma execucao sem vagas novas nao grava nada: nao apaga o relatorio bom do dia
        caminho = salvar_planilha_do_dia(cfg, finais, hoje)
        registrar_no_historico(historico, finais, hoje)
        salvar_historico(cfg, historico)
        print(f"      Pronto: {caminho}")
    elif novas:
        print(f"      {len(novas)} vaga(s) nova(s), mas nenhuma atingiu a nota minima ({cfg.nota_minima}).")
        print("      A planilha do dia nao foi criada. Para ver as notas: python agente_vagas.py --debug")
    else:
        print("      Nenhuma vaga nova hoje - a planilha do dia nao foi alterada.")

    print("[6/6] Atualizando o controle de candidaturas...")
    arrumar_e_informar(cfg, atualizar_e_informar_controle(cfg, historico, palavras))

    print(f"\nHistorico: {len(historico)} vaga(s) acumulada(s).")
    # Sempre mostra algo ao terminar, para ficar claro que o agente rodou:
    # a planilha do dia, ou o controle de candidaturas quando nao houver vaga nova
    controle = os.path.join(cfg.dir_saida, NOME_CONTROLE)
    if caminho:
        abrir_no_windows(caminho)
    elif os.path.exists(controle):
        print("Abrindo o controle de candidaturas (nao houve vaga nova para a planilha do dia).")
        abrir_no_windows(controle)
    return 0


def main(argv=None) -> int:
    # Escreve a saida na hora, sem acumular: se travar, o registro.txt mostra ate onde foi
    for fluxo in (sys.stdout, sys.stderr):
        try:
            fluxo.reconfigure(line_buffering=True)  # type: ignore[attr-defined]
        except AttributeError:
            pass
    try:
        return executar(ler_argumentos(argv))
    except (ErroConfiguracao, ErroEmail) as e:
        print(f"\nERRO: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())

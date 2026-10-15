"""Execucao completa do agente, com o Gmail simulado pelo e-mail de exemplo."""
import json
import os
import shutil

from openpyxl import load_workbook

import agente_vagas
import configuracao
from conftest import PASTA_COLETOR


def preparar(tmp_path, monkeypatch, html_exemplo, extra_config=""):
    pasta_agente = tmp_path / "agente"
    pasta_agente.mkdir()
    shutil.copy(os.path.join(PASTA_COLETOR, "palavras_chave.txt"), pasta_agente)
    saida = tmp_path / "saida"
    (pasta_agente / "config.txt").write_text(
        "email = alertas@exemplo.com\nsenha = abcdefghijklmnop\nnota_minima = 4\n"
        f"pasta_saida = {saida}\npasta_adaptador = {tmp_path / 'Adaptador'}\nia_titulos = nao\n"
        + extra_config, encoding="utf-8")
    monkeypatch.setattr(configuracao, "PASTA_AGENTE", str(pasta_agente))
    monkeypatch.setattr(agente_vagas, "ler_emails",
                        lambda cfg, pasta, dias: [{"assunto": "Alerta", "html": html_exemplo}])
    return saida


def test_execucao_completa(tmp_path, monkeypatch, html_exemplo, capsys):
    saida = preparar(tmp_path, monkeypatch, html_exemplo)
    assert agente_vagas.main([]) == 0

    planilhas = list(saida.glob("vagas_*.xlsx"))
    assert len(planilhas) == 1
    ws = load_workbook(planilhas[0]).active
    titulos = [ws.cell(row=r, column=3).value for r in range(2, ws.max_row + 1)]
    assert titulos == ["Analista de Processos e Melhoria Contínua", "Analista de Requisitos Sênior"]

    historico = json.loads((saida / "vagas_vistas.json").read_text(encoding="utf-8"))
    assert set(historico) == {"li-1000000001", "li-1000000002"}
    assert historico["li-1000000001"]["nota"] == 9

    controle = load_workbook(saida / "controle_candidaturas.xlsx")
    assert controle["Candidaturas"].max_row == 3
    assert "Agente de Vagas v1.5" in capsys.readouterr().out


def test_segunda_execucao_nao_repete_vagas_nem_apaga_a_planilha(tmp_path, monkeypatch, html_exemplo, capsys):
    saida = preparar(tmp_path, monkeypatch, html_exemplo)
    agente_vagas.main([])
    planilha = next(saida.glob("vagas_*.xlsx"))
    antes = planilha.read_bytes()

    assert agente_vagas.main([]) == 0
    assert "Nenhuma vaga nova hoje" in capsys.readouterr().out
    assert planilha.read_bytes() == antes


def test_refazer_nao_apaga_o_historico(tmp_path, monkeypatch, html_exemplo):
    saida = preparar(tmp_path, monkeypatch, html_exemplo)
    saida.mkdir()
    (saida / "vagas_vistas.json").write_text(json.dumps({
        "li-999": {"data": "2026-09-01", "titulo": "Vaga antiga", "empresa": "X", "link": ""}}), encoding="utf-8")
    agente_vagas.main(["--refazer"])
    historico = json.loads((saida / "vagas_vistas.json").read_text(encoding="utf-8"))
    assert "li-999" in historico and "li-1000000001" in historico


def test_config_faltando_retorna_erro(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(configuracao, "PASTA_AGENTE", str(tmp_path))
    assert agente_vagas.main([]) == 1
    assert "config.txt" in capsys.readouterr().out


def test_segunda_execucao_com_vagas_novas_nao_sobrescreve_a_da_manha(tmp_path, monkeypatch, html_exemplo):
    saida = preparar(tmp_path, monkeypatch, html_exemplo)
    agente_vagas.main([])
    da_manha = next(saida.glob("vagas_*.xlsx"))
    antes = da_manha.read_bytes()
    # Chega um alerta novo no mesmo dia
    novo = html_exemplo.replace("1000000001", "2000000001").replace("1000000002", "2000000002")
    monkeypatch.setattr(agente_vagas, "ler_emails", lambda cfg, pasta, dias: [{"assunto": "A", "html": novo}])
    agente_vagas.main([])
    assert da_manha.read_bytes() == antes
    assert len(list(saida.glob("vagas_*.xlsx"))) == 2

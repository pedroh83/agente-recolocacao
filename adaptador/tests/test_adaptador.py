"""Testes das regras do adaptador que nao dependem da IA (a IA e simulada)."""
import csv
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import adaptador  # noqa: E402

CV_A = """Fulano de Tal
Analista da Qualidade
EXPERIÊNCIA
Analista da Qualidade — Empresa Alfa (03/2018 - 04/2020)
Acompanhei auditorias internas da ISO 9001.
Conduzi a migração da ISO 9001:2008 para ISO 9001:2015, incluindo recertificação.
Reduzi em 30% as não conformidades recorrentes.
Analista de Processos — Empresa Beta (06/2021 - 07/2022)
Mapeei processos em BPMN e elaborei procedimentos operacionais.
FORMAÇÃO
Administração — Faculdade Exemplo (2012)
IDIOMAS
Inglês — curso básico
""" + "Experiência com PDCA, indicadores e melhoria contínua. " * 5

CV_B = CV_A.replace("Analista da Qualidade\nEXPERIÊNCIA", "Analista de Requisitos\nEXPERIÊNCIA")


def resultado_ia(**mudancas):
    base = {
        "empresa": "Empresa Nova", "cargo": "Analista de Processos", "versao_escolhida": "A",
        "compatibilidade_base": 50, "compatibilidade_adaptado": 80,
        "curriculo": {
            "experiencias": [
                {"cargo": "Analista de Processos", "empresa": "Empresa Beta", "periodo": "06/2021 - 07/2022",
                 "topicos": ["Mapeio processos em BPMN.", "Elaboro procedimentos operacionais."]},
                {"cargo": "Analista da Qualidade", "empresa": "Empresa Alfa", "periodo": "03/2018 - 04/2020",
                 "topicos": ["Reduzi em 30% as não conformidades recorrentes."]},
            ],
            "formacao": ["Administração — Faculdade Exemplo (2012)"],
            "idiomas": ["Inglês — curso básico"],
            "cursos_certificacoes": [],
        },
    }
    base["curriculo"].update(mudancas)
    return base


# ---------------------------------------------------------------- pontuacao local

def test_pontuacao_local_conta_termos_da_vaga_no_curriculo():
    nota, achados, faltando = adaptador.pontuar("Analista de processos com BPMN e PDCA. Desejável SAP.", CV_A)
    assert "bpmn" in achados and "pdca" in achados
    assert "sap" in faltando
    assert 0 < nota < 100


# ---------------------------------------------------------------- correcoes automaticas

def test_corrige_verbos_para_o_passado_em_emprego_encerrado():
    r = resultado_ia()
    trocas = adaptador.corrigir_tempo_verbal(r)
    topicos = r["curriculo"]["experiencias"][0]["topicos"]
    assert topicos == ["Mapeei processos em BPMN.", "Elaborei procedimentos operacionais."]
    assert trocas == 2


def test_nao_mexe_em_emprego_atual():
    r = resultado_ia()
    r["curriculo"]["experiencias"][0]["periodo"] = "06/2021 - atual"
    adaptador.corrigir_tempo_verbal(r)
    assert r["curriculo"]["experiencias"][0]["topicos"][0] == "Mapeio processos em BPMN."


# ---------------------------------------------------------------- veracidade

def test_curriculo_fiel_nao_gera_alerta():
    assert adaptador.verificar_veracidade(resultado_ia(), {"A": CV_A, "B": CV_B}) == []


def test_alerta_para_empresa_inventada():
    r = resultado_ia()
    r["curriculo"]["experiencias"][0]["empresa"] = "Empresa Fantasma"
    alertas = adaptador.verificar_veracidade(r, {"A": CV_A})
    assert any("Fantasma" in a for a in alertas)


def test_alerta_para_numero_inventado():
    r = resultado_ia()
    r["curriculo"]["experiencias"][1]["topicos"] = ["Reduzi em 45% as não conformidades."]
    alertas = adaptador.verificar_veracidade(r, {"A": CV_A})
    assert any("45%" in a for a in alertas)


def test_alerta_para_nivel_de_idioma_inventado():
    r = resultado_ia(idiomas=["Inglês — fluente"])
    alertas = adaptador.verificar_veracidade(r, {"A": CV_A})
    assert any("fluente" in a for a in alertas)


def test_alerta_para_papel_promovido():
    r = resultado_ia()
    r["curriculo"]["experiencias"][1]["topicos"] = ["Liderei auditorias internas da ISO 9001."]
    alertas = adaptador.verificar_veracidade(r, {"A": CV_A})
    assert any("liderei" in a for a in alertas)


# ---------------------------------------------------------------- auxiliares

def test_extrai_json_mesmo_com_cerca_de_codigo():
    assert adaptador.extrair_json('```json\n{"a": 1}\n```') == {"a": 1}
    with pytest.raises(ValueError):
        adaptador.extrair_json("sem json aqui")


def test_auxiliares_de_nome_e_link():
    assert adaptador.limpar_nome("Analista de Negócios Sr.") == "Analista_De_Negocios_Sr"
    assert adaptador.extrair_link("Vaga: https://www.linkedin.com/jobs/view/123/). Fim") == \
        "https://www.linkedin.com/jobs/view/123/"


# ---------------------------------------------------------------- fluxo completo, IA simulada

def test_processar_vaga_gera_docx_relatorio_e_historico(tmp_path, monkeypatch):
    pytest.importorskip("docx")
    monkeypatch.setattr(adaptador, "PASTA_SAIDA", tmp_path / "saida")
    monkeypatch.setattr(adaptador, "ARQ_HISTORICO", tmp_path / "historico.csv")
    vaga = tmp_path / "vaga.txt"
    vaga.write_text("Analista de Processos — Empresa Nova. Requisitos: BPMN, PDCA, indicadores, "
                    "procedimentos operacionais, melhoria contínua e ISO 9001. " * 3, encoding="utf-8")
    cfg = {"IDIOMA": "português do Brasil", "NOME": "Fulano de Tal", "CONTATO": "Cidade - UF"}

    docx = adaptador.processar_vaga(vaga, {"A": CV_A, "B": CV_B}, cfg, chamar=lambda p, c: resultado_ia())

    assert docx.exists() and docx.suffix == ".docx"
    assert len(list((tmp_path / "saida").glob("Relatorio_*.txt"))) == 1
    with open(tmp_path / "historico.csv", encoding="utf-8-sig") as f:
        linhas = list(csv.DictReader(f, delimiter=";"))
    assert linhas[0]["Empresa"] == "Empresa Nova" and linhas[0]["Versão"] == "A"

import datetime as dt
import os

from arquivamento import NOME_PASTA, arrumar_pasta

HOJE = dt.date(2026, 10, 20)


def criar(pasta, *nomes):
    for nome in nomes:
        (pasta / nome).write_bytes(b"x")


def test_move_planilhas_antigas_e_copias_de_correcao(tmp_path):
    criar(tmp_path, "vagas_2026-10-01.xlsx", "vagas_2026-10-12_234004.xlsx",  # mais de 7 dias
          "vagas_2026-10-14.xlsx", "vagas_2026-10-20.xlsx",                   # recentes
          "vagas_vistas_antes_da_correcao.json",
          "controle_candidaturas.xlsx", "controle_candidaturas_backup.xlsx", "vagas_vistas.json")
    movidos = arrumar_pasta(str(tmp_path), dias=7, hoje=HOJE)
    assert sorted(movidos) == ["vagas_2026-10-01.xlsx", "vagas_2026-10-12_234004.xlsx",
                               "vagas_vistas_antes_da_correcao.json"]
    sobrou = sorted(p.name for p in tmp_path.iterdir() if p.is_file())
    assert sobrou == ["controle_candidaturas.xlsx", "controle_candidaturas_backup.xlsx",
                      "vagas_2026-10-14.xlsx", "vagas_2026-10-20.xlsx", "vagas_vistas.json"]
    assert (tmp_path / NOME_PASTA / "vagas_2026-10-01.xlsx").exists()
    assert (tmp_path / NOME_PASTA / "LEIA-ME.txt").exists()


def test_sem_controle_atualizado_as_planilhas_ficam(tmp_path):
    criar(tmp_path, "vagas_2026-10-01.xlsx", "vagas_vistas_antes_da_correcao.json")
    movidos = arrumar_pasta(str(tmp_path), dias=7, hoje=HOJE, controle_atualizado=False)
    assert movidos == ["vagas_vistas_antes_da_correcao.json"]
    assert (tmp_path / "vagas_2026-10-01.xlsx").exists()


def test_zero_dias_desliga_o_arquivamento_de_planilhas(tmp_path):
    criar(tmp_path, "vagas_2020-01-01.xlsx")
    assert arrumar_pasta(str(tmp_path), dias=0, hoje=HOJE) == []


def test_nao_sobrescreve_arquivo_que_ja_esta_na_pasta(tmp_path):
    (tmp_path / NOME_PASTA).mkdir()
    (tmp_path / NOME_PASTA / "vagas_2026-10-01.xlsx").write_bytes(b"antigo")
    criar(tmp_path, "vagas_2026-10-01.xlsx")
    arrumar_pasta(str(tmp_path), dias=7, hoje=HOJE)
    assert (tmp_path / NOME_PASTA / "vagas_2026-10-01.xlsx").read_bytes() == b"antigo"
    assert (tmp_path / NOME_PASTA / "vagas_2026-10-01 (2).xlsx").exists()


def test_arquivo_aberto_no_excel_fica_para_a_proxima(tmp_path, monkeypatch):
    criar(tmp_path, "vagas_2026-10-01.xlsx", "vagas_2026-10-02.xlsx")
    original = os.replace

    def bloqueia_um(origem, destino):
        if "10-01" in str(origem):
            raise PermissionError("em uso")
        return original(origem, destino)

    monkeypatch.setattr(os, "replace", bloqueia_um)
    assert arrumar_pasta(str(tmp_path), dias=7, hoje=HOJE) == ["vagas_2026-10-02.xlsx"]
    assert (tmp_path / "vagas_2026-10-01.xlsx").exists()


def test_ignora_arquivo_temporario_do_excel(tmp_path):
    criar(tmp_path, "~$vagas_2026-10-01.xlsx")
    assert arrumar_pasta(str(tmp_path), dias=7, hoje=HOJE) == []

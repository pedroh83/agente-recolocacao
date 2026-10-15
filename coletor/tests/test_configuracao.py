import pytest

from configuracao import ErroConfiguracao, carregar_config, carregar_palavras_chave


def escrever(tmp_path, nome, texto):
    caminho = tmp_path / nome
    caminho.write_text(texto, encoding="utf-8")
    return str(caminho)


def test_le_config_com_comentarios_e_valores_padrao(tmp_path):
    cfg = carregar_config(escrever(tmp_path, "config.txt", (
        "# comentario\n"
        "email = alertas@exemplo.com\n"
        "senha = abcd efgh ijkl mnop\n"
        "nota_minima = 4\n"
        "chave_que_nao_existe = 1\n")))
    assert cfg.email == "alertas@exemplo.com"
    assert cfg.senha_sem_espacos == "abcdefghijklmnop"
    assert cfg.nota_minima == 4
    assert cfg.dias == 2 and cfg.pasta == "INBOX" and cfg.ia_titulos == "auto"


def test_config_com_numero_invalido_mostra_a_linha(tmp_path):
    with pytest.raises(ErroConfiguracao, match="linha 2"):
        carregar_config(escrever(tmp_path, "config.txt", "email = x\nnota_minima = quatro\n"))


def test_config_faltando(tmp_path):
    with pytest.raises(ErroConfiguracao, match="config.txt"):
        carregar_config(str(tmp_path / "nao_existe.txt"))


def test_palavras_chave_aceitam_acento_e_maiuscula(tmp_path):
    p = carregar_palavras_chave(escrever(tmp_path, "pc.txt", (
        "[A]  # qualidade\n"
        "peso 3: Gestão da Qualidade, ISO 9001\n"
        "peso 1: PDCA,\n"
        "[B]\n"
        "peso 3: Analista de Requisitos\n"
        "[DESCARTE]\n"
        "Estágio\n"
        "[LOCAL]\n"
        "bonus: 3\n"
        "Salvador\n")))
    assert p.curriculos["A"][3] == ["gestao da qualidade", "iso 9001"]
    assert p.curriculos["A"][1] == ["pdca"]
    assert p.descarte == ["estagio"]
    assert p.local == ["salvador"] and p.bonus_local == 3


def test_palavras_chave_com_erro_indicam_a_linha(tmp_path):
    with pytest.raises(ErroConfiguracao, match="linha 2"):
        carregar_palavras_chave(escrever(tmp_path, "pc.txt", "[A]\niso 9001, sgq\n"))


def test_palavras_chave_sem_secao_b(tmp_path):
    with pytest.raises(ErroConfiguracao, match=r"\[B\]"):
        carregar_palavras_chave(escrever(tmp_path, "pc.txt", "[A]\npeso 3: sgq\n"))


def test_arquivo_oficial_carrega(palavras):
    assert "iso 9001" in palavras.curriculos["A"][3]
    assert "analista de requisitos" in palavras.curriculos["B"][3]
    assert "jovem aprendiz" in palavras.descarte


def test_arquivos_salvos_em_ansi_pelo_bloco_de_notas(tmp_path):
    caminho = tmp_path / "pc.txt"
    caminho.write_bytes("[A]\npeso 3: gestão da qualidade\n[B]\npeso 3: análise de negócios\n".encode("cp1252"))
    p = carregar_palavras_chave(str(caminho))
    assert p.curriculos["A"][3] == ["gestao da qualidade"]
    assert p.curriculos["B"][3] == ["analise de negocios"]

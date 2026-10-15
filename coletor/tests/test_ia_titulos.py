from configuracao import Config
from ia_titulos import ia_ligada, interpretar_resposta, reavaliar_titulos


def vagas_de_teste():
    return [
        {"titulo": "Especialista em Sistema de Gestão", "nota": 1, "descartada": "", "curriculo": "A", "termos": "processos"},
        {"titulo": "Motorista de Entregas", "nota": 0, "descartada": "", "curriculo": "?", "termos": ""},
        {"titulo": "Estágio em Processos", "nota": 0, "descartada": "contem 'estagio'", "curriculo": "-", "termos": ""},
        {"titulo": "Analista de Processos", "nota": 7, "descartada": "", "curriculo": "A", "termos": "x"},
    ]


def test_ia_resgata_so_as_vagas_aprovadas():
    vagas = vagas_de_teste()
    enviados = []

    def ia_falsa(prompt):
        enviados.append(prompt)
        return '[{"i": 0, "perfil": "A"}, {"i": 1, "perfil": "N"}]'

    resgatadas = reavaliar_titulos(vagas, Config(), nota_minima=4, chamar=ia_falsa)
    assert resgatadas == 1
    assert vagas[0]["nota"] == 4 and vagas[0]["origem"] == "IA"
    assert vagas[1]["nota"] == 0
    # Descartadas e vagas que ja passaram na nota nao vao para a IA
    assert "Estágio" not in enviados[0] and "Analista de Processos\n" not in enviados[0]


def test_falha_da_ia_nao_interrompe_o_agente():
    vagas = vagas_de_teste()
    avisos = []

    def ia_quebrada(prompt):
        raise RuntimeError("servidor ocupado")

    assert reavaliar_titulos(vagas, Config(), 4, chamar=ia_quebrada, avisar=avisos.append) == 0
    assert vagas[0]["nota"] == 1
    assert "servidor ocupado" in avisos[0]


def test_resposta_com_lixo_e_indices_invalidos():
    texto = '```json\n[{"i": 0, "perfil": "b"}, {"i": 9, "perfil": "A"}, {"x": 1}]\n```'
    assert interpretar_resposta(texto, 2) == {0: "B"}


def test_ia_desligada_sem_chave_ou_por_configuracao():
    chave = {"GEMINI_API_KEY": "A" * 39}
    assert ia_ligada(Config(), chave)
    assert not ia_ligada(Config(), {})
    assert not ia_ligada(Config(), {"GEMINI_API_KEY": "COLE_SUA_CHAVE_AQUI"})
    assert not ia_ligada(Config(ia_titulos="nao"), chave)


def test_sem_chave_nao_chama_ninguem(tmp_path):
    vagas = vagas_de_teste()
    cfg = Config(pasta_adaptador=str(tmp_path))
    assert reavaliar_titulos(vagas, cfg, 4, cfg_ia={}) == 0

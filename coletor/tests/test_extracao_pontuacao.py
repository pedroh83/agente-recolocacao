from extracao import deduzir_empresa_e_local, extrair_vagas, id_da_vaga, link_limpo, separar_titulo
from pontuacao import avaliar, casa_termo
from texto import limpar_html, normalizar


# ---------------------------------------------------------------- texto

def test_normalizar_tira_acento_e_maiuscula():
    assert normalizar("Gestão da QUALIDADE") == "gestao da qualidade"


def test_limpar_html_remove_tag_cortada_pela_metade():
    assert limpar_html("Analista</a><table cellp") == "Analista"


# ---------------------------------------------------------------- extracao

def test_extrai_as_vagas_do_alerta_e_ignora_botoes_e_imagens(html_exemplo):
    vagas = extrair_vagas(html_exemplo)
    titulos = [v["titulo"] for v in vagas]
    assert "Ver vaga" not in titulos and "Candidatar-se" not in titulos
    assert "Ver todas as vagas" not in titulos  # link de busca, nao de vaga
    assert titulos[:4] == [
        "Analista de Processos e Melhoria Contínua",
        "Analista de Requisitos Sênior",
        "Estágio em Qualidade",
        "Especialista em Logística Reversa",
    ]


def test_separa_empresa_e_local(html_exemplo):
    v = extrair_vagas(html_exemplo)[0]
    assert v["empresa"] == "Empresa Alfa Ltda"
    assert v["local"] == "Salvador, Bahia (Híbrido)"


def test_local_sem_cidade():
    assert deduzir_empresa_e_local("Beta Tecnologia S.A. · Brasil (Remoto)", "") == \
        ("Beta Tecnologia S.A.", "Brasil (Remoto)")


def test_link_fica_curto_e_sem_tokens_pessoais(html_exemplo):
    v = extrair_vagas(html_exemplo)[0]
    assert v["link"] == "https://www.linkedin.com/jobs/view/1000000001/"
    assert "otpToken" not in v["link"]


def test_id_estavel_com_ou_sem_comm_no_link():
    a = id_da_vaga("https://www.linkedin.com/comm/jobs/view/123/?trk=x")
    b = id_da_vaga("https://www.linkedin.com/jobs/view/123/")
    assert a == b == "li-123"
    assert link_limpo("https://www.linkedin.com/comm/jobs/view/123/?x=1") == "https://www.linkedin.com/jobs/view/123/"


def test_id_sem_link_usa_titulo_e_empresa():
    assert id_da_vaga("", "Analista", "Empresa") == id_da_vaga("", "ANALISTA", "empresa")


# ---------------------------------------------------------------- pontuacao

def test_palavra_inteira_evita_falso_positivo():
    assert not casa_termo("analista de logistica", "ti")
    assert not casa_termo("vaga popular", "pop")
    assert casa_termo("analista de ti pleno", "ti")
    assert casa_termo("certificacao 5s", "5s")


def test_vaga_de_processos_vai_para_o_curriculo_a(html_exemplo, palavras):
    v = avaliar(extrair_vagas(html_exemplo)[0], palavras)
    assert v["curriculo"] == "A"
    assert v["descartada"] == ""
    # analista de processos (3) + melhoria continua (3) + processos (1) + bonus local (2)
    assert v["nota"] == 9
    assert v["termos"].startswith("melhoria continua")


def test_vaga_de_requisitos_vai_para_o_curriculo_b(html_exemplo, palavras):
    v = avaliar(extrair_vagas(html_exemplo)[1], palavras)
    assert v["curriculo"] == "B"
    assert v["nota"] >= 3 + 2  # analista de requisitos + bonus remoto


def test_estagio_e_descartado_pelo_titulo(html_exemplo, palavras):
    v = avaliar(extrair_vagas(html_exemplo)[2], palavras)
    assert v["descartada"] == "contem 'estagio'"
    assert v["nota"] == 0


def test_vaga_fora_do_perfil_fica_sem_curriculo(html_exemplo, palavras):
    v = avaliar(extrair_vagas(html_exemplo)[3], palavras)
    assert v["curriculo"] == "?"
    assert v["nota"] == 0


def test_link_que_envolve_o_cartao_inteiro_nao_engole_empresa_e_local(html_exemplo):
    v = [v for v in extrair_vagas(html_exemplo) if "1000000006" in v["link"]][0]
    assert v["titulo"] == "Analista de Sistemas Sênior"
    assert v["empresa"] == "Épsilon Tecnologia"
    assert v["local"] == "Salvador, BA"  # sem o salario e sem o selo


def test_separar_titulo_sem_cartao_nao_muda_nada():
    assert separar_titulo("<b>Analista de Processos</b>") == ("Analista de Processos", "")


def test_local_com_modalidade_e_texto_depois():
    assert deduzir_empresa_e_local("Oficial Atakarejo · Salvador, BA (Presencial) Crescimento", "") == \
        ("Oficial Atakarejo", "Salvador, BA (Presencial)")

# POP — Agente de Vagas

**Procedimento Operacional Padrão**
Responsável: Pedro Henrique Simas
Versão: 1.2 — outubro/2026
Versão do programa: 1.5.1

---

## 1. Objetivo

Automatizar o levantamento diário de vagas de emprego, entregando toda manhã
uma planilha com as vagas novas, ordenadas por compatibilidade com os
currículos, indicando qual versão enviar e com o link de candidatura pronto.

Manter um **controle único de candidaturas**, que junta as vagas, os envios e
as adaptações de currículo, com **indicadores** para avaliar e melhorar a busca
(ciclo PDCA).

## 2. Escopo

Aplica-se ao computador pessoal (Windows), à conta de e-mail
`seu-email-de-alertas@gmail.com`, aos alertas de vaga configurados no LinkedIn e ao
histórico do Adaptador de Currículo (`C:\AgenteVagas\Adaptador`).

**Fora do escopo:** o agente não busca vagas nos sites. Ele organiza os alertas
que o LinkedIn envia por e-mail. A qualidade do relatório depende diretamente
da qualidade dos alertas configurados.

## 3. Definições

| Termo | Significado |
|---|---|
| **Alerta de vaga** | Busca salva no LinkedIn que envia resultados por e-mail |
| **Senha de app** | Senha de 16 dígitos gerada pelo Google para programas acessarem a conta |
| **Currículo A** | Perfil Qualidade, Processos e ISO |
| **Currículo B** | Perfil Requisitos, Negócios e Tecnologia |
| **Nota** | Pontuação de compatibilidade calculada pelo agente |
| **Histórico** | Registro das vagas já exibidas, evita repetição |
| **Controle** | Planilha única com todas as vagas, envios, retornos e indicadores |
| **Retorno positivo** | Entrevista, Teste/Case ou Proposta |

## 4. Como funciona

```
Madrugada    LinkedIn processa os alertas e envia os e-mails
                          ↓
             E-mails chegam na caixa de entrada do Gmail
                          ↓
9h           Agendador do Windows dispara o rodar_agente.bat
                          ↓
             1. Conecta no Gmail e lê os e-mails dos últimos 2 dias
                (sem marcá-los como lidos)
             2. Extrai título, empresa, local e link de cada vaga
             3. Descarta as que já estão no histórico e os anúncios
                repetidos (mesmo cargo e empresa nos últimos 30 dias)
             4. Pontua contra as palavras-chave dos currículos A e B
                (palavras_chave.txt). Os títulos de nota baixa recebem
                uma segunda opinião da IA (Gemini)
             5. Gera a planilha do dia ordenada por nota
             6. Atualiza o controle de candidaturas: junta histórico,
                envios marcados nas planilhas do dia e adaptações do
                Adaptador de Currículo
                          ↓
             A planilha do dia abre automaticamente na tela
```

---

## 5. Configuração instalada

### 5.1 Localização dos arquivos

**Pasta do agente:** `C:\AgenteVagas`

| Arquivo | Função |
|---|---|
| `agente_vagas.py` | Programa principal (coordena as etapas) |
| `configuracao.py` | Lê o `config.txt` e o `palavras_chave.txt` |
| `email_gmail.py` | Conexão com o Gmail |
| `extracao.py` | Lê as vagas de dentro dos e-mails |
| `pontuacao.py` | Calcula a nota de cada vaga |
| `ia_titulos.py` | Segunda opinião da IA sobre títulos de nota baixa |
| `planilha_dia.py` | Gera a planilha do dia |
| `controle.py` | Gera o controle de candidaturas e os indicadores |
| `texto.py` | Funções de texto usadas pelos demais |
| `palavras_chave.txt` | **Listas de palavras-chave** — editar aqui para calibrar |
| `config.txt` | **Configuração e senha** — não compartilhar |
| `rodar_agente.bat` | Execução automática (usado pelo Agendador) |
| `testar_agora.bat` | Execução manual, mostra tudo na tela |
| `diagnostico.py` | Verifica a conexão e o conteúdo da conta |
| `inspecionar.py` | Mostra o conteúdo de cada e-mail |
| `registro.txt` | Log das execuções automáticas |
| `backup_v1.4\` | Versão anterior do programa, para voltar atrás se preciso (ver 9.6) |

Todos os arquivos `.py` precisam ficar juntos na mesma pasta.

**Pasta de saída:** `C:\Users\SEU_USUARIO\AgenteVagas`

| Arquivo | Função |
|---|---|
| `vagas_AAAA-MM-DD.xlsx` | Relatório do dia |
| `controle_candidaturas.xlsx` | Controle único de candidaturas e indicadores |
| `controle_candidaturas_backup.xlsx` | Cópia do controle antes da última atualização |
| `vagas_vistas.json` | Histórico de vagas já exibidas |

### 5.2 Conta de e-mail

- Conta: `seu-email-de-alertas@gmail.com`
- Verificação em duas etapas: **ativada** (obrigatória)
- Senha de app: gerenciada em `myaccount.google.com/apppasswords`
- Leitura: caixa de entrada (`INBOX`), sem filtro ou marcador

### 5.3 Agendamento

- Agendador de Tarefas do Windows, tarefa **"Agente de Vagas"**
- Gatilho: diário, 09:00
- Ação: `C:\AgenteVagas\rodar_agente.bat`
- Executar somente quando o usuário estiver conectado
- Condições: desmarcado "somente em fonte de alimentação CA"
- Configurações: marcado "executar após inicialização perdida"

### 5.4 Pré-requisitos instalados

- Python 3.8 ou mais novo, instalado com a opção **Add python.exe to PATH**
- Biblioteca `openpyxl` (`pip install openpyxl`)
- Biblioteca `google-genai`, já instalada pelo Adaptador de Currículo
  (só é usada pela segunda opinião da IA; sem ela o agente funciona normalmente)

### 5.5 Verificar a versão em uso

Ao rodar, a primeira linha exibe a versão:

```
=== Agente de Vagas v1.5.1 ===
```

Conferir sempre após uma atualização. Se o número não mudar, o arquivo antigo
ainda está em uso — o Windows costuma salvar a cópia nova como
`agente_vagas (1).py` sem substituir. Procedimento correto: **apagar o
arquivo antigo primeiro**, depois mover o novo.

### 5.6 Opções do config.txt

| Opção | Padrão | Função |
|---|---|---|
| `email`, `senha` | — | Conta Gmail e senha de app (obrigatórias) |
| `pasta` | `INBOX` | Pasta ou marcador onde os alertas chegam |
| `nota_minima` | `1` (em uso: `4`) | Nota mínima para a vaga entrar no relatório |
| `dias` | `2` | Dias de e-mail para trás a cada execução |
| `ia_titulos` | `auto` | `auto` liga a IA se o Adaptador tiver chave; `nao` desliga |
| `ia_maximo` | `40` | Máximo de títulos enviados à IA por execução |
| `pasta_saida` | `C:\Users\SEU_USUARIO\AgenteVagas` | Onde ficam planilhas, histórico e controle |
| `pasta_adaptador` | `C:\AgenteVagas\Adaptador` | Onde está o Adaptador de Currículo |

Só as opções que você escrever no `config.txt` mudam; as demais usam o padrão.
Um erro de digitação (ex.: `nota_minima = quatro`) faz o agente parar com uma
mensagem que indica a linha do problema.

---

## 6. Rotina diária de uso

1. A planilha do dia abre sozinha por volta das 9h
2. As vagas vêm ordenadas por nota, com as melhores no topo, em verde.
   Vagas com "IA" na coluna Termos foram reconhecidas pela IA, não pelas
   palavras-chave: confira se fazem sentido
3. Para cada vaga de interesse:
   - Consultar a coluna **Currículo** — indica se envia a versão A ou B
   - Clicar em **Candidatar-se** — abre a vaga no LinkedIn
   - Se quiser, adaptar o currículo com o atalho **Adaptar Curriculo**
   - Após se candidatar, escrever "sim" na coluna **Enviado?**
4. Salvar a planilha ao final

**Não é preciso copiar nada para o controle.** Na execução seguinte, o agente
leva o "sim" das planilhas do dia e as adaptações do Adaptador para o
`controle_candidaturas.xlsx`.

**Planilha vazia** significa que não houve vaga nova desde a última execução.
Não é falha.

### 6.1 Rotina semanal (PDCA)

1. Abrir o `controle_candidaturas.xlsx`
2. Na aba **Candidaturas**, atualizar a coluna **Retorno** das empresas que
   responderam (Entrevista, Teste/Case, Proposta, Recusado, Sem retorno)
3. Na aba **Indicadores**, analisar:
   - **Qual currículo traz mais retorno?** — se um deles tiver taxa bem menor,
     revisar esse currículo
   - **Adaptar o currículo faz diferença?** — compara a taxa de retorno das
     candidaturas adaptadas com as não adaptadas
   - **A nota do agente acerta?** — se notas altas não trazem mais retorno,
     recalibrar o `palavras_chave.txt` (8.1)
   - **Últimas 8 semanas** — ritmo de vagas novas e de candidaturas enviadas
4. Registrar as decisões tomadas e ajustar o que for preciso

---

## 7. Operação manual

| Necessidade | Comando |
|---|---|
| Rodar agora, vendo o resultado na tela | `testar_agora.bat` (dois cliques) |
| Rodar mostrando o detalhamento | `python agente_vagas.py --debug` |
| Ver de novo vagas já exibidas | `python agente_vagas.py --refazer` |
| Buscar mais dias para trás (após viagem) | `python agente_vagas.py --dias 10` |
| Rodar sem consultar a IA | `python agente_vagas.py --sem-ia` |
| Só atualizar o controle de candidaturas | `python agente_vagas.py --so-controle` |
| Verificar a conexão com o Gmail | `python diagnostico.py` |
| Ver o conteúdo dos e-mails recebidos | `python inspecionar.py` |

Todos executados a partir do Prompt de Comando, após:

```
cd C:\AgenteVagas
```

---

## 8. Manutenção

### 8.1 Ajustar as palavras-chave

As listas ficam no arquivo **`palavras_chave.txt`**, separadas em seções:
`[A]`, `[B]`, `[DESCARTE]` e `[LOCAL]`. Cada currículo tem termos de
**peso 3** (fortes, quase definem a vaga) e **peso 1** (apoio). A nota final é
a soma dos pesos encontrados, mais o bônus da seção `[LOCAL]` (2 pontos) se a
vaga citar Salvador, Bahia, Petrolina, remoto, home office ou híbrido.
Vagas presenciais ou híbridas em outros estados **não são descartadas**: só
deixam de ganhar o bônus.

**Procedimento:**

1. Abrir `palavras_chave.txt` no Bloco de Notas
2. Acrescentar o termo numa linha `peso 3:` ou `peso 1:` da seção certa,
   separado por vírgula. Pode escrever com acento e maiúscula
3. Salvar com `Ctrl+S` e testar com `python agente_vagas.py --debug --refazer`

Se houver erro de formato, o agente para e mostra a linha do problema
(ex.: `palavras_chave.txt, linha 23: ...`). Não é mais preciso abrir o
`agente_vagas.py`.

**Regras aprendidas na calibragem (importantes):**

- A comparação é por **palavra inteira**. Termos curtos como `qa`, `ti` ou `5s`
  são seguros: não casam dentro de outras palavras.
- As vagas variam a redação. Cadastrar as formas que aparecem de fato:
  `analista de negócios`, `analista de negócio`, `analista negócios`.
- Termo genérico demais no peso 3 contamina o resultado. Exemplo removido:
  `analista de operações` trazia vagas de operação sem relação com qualidade.
- Testar sempre com `--refazer` antes de dar por encerrado.

O conteúdo atual das listas está no próprio `palavras_chave.txt`.

### 8.2 Ajustar o filtro de nota

Editar `config.txt`:

```
nota_minima = 4
```

Valor maior = relatório mais enxuto e exigente. Valor 1 traz quase tudo.

**Referência de calibragem (agosto/2026, 110 vagas avaliadas):**

| Nota mínima | Resultado |
|---|---|
| 1 | Traz tudo, inclusive vagas sem relação |
| 3 | ~60 vagas — inclui alguns falsos positivos |
| **4** | ~58 vagas — **valor em uso**, bom equilíbrio |
| 5 ou mais | Só o núcleo, com risco de perder vaga boa |

**Não calibrar com base numa execução de acúmulo.** A primeira rodada após
vários dias sem execução traz a fila inteira. O volume diário normal é bem
menor (5 a 15 vagas). Apertar a nota com base no acúmulo esvazia os relatórios
seguintes.

A partir da v1.5, a aba **Indicadores** do controle mostra a taxa de retorno
por faixa de nota. Usar esses números para recalibrar.

### 8.3 Revisar os alertas do LinkedIn

Recomendado a cada dois meses, ou sempre que o relatório ficar pobre.

Acessar **LinkedIn → Vagas → Alertas de vaga** e confirmar:
- Todos os alertas continuam listados
- Todos com frequência **diária**
- Destino apontando para o Gmail correto

**Plano de alertas em uso:**

| Bloco | Localidade | Termos |
|---|---|---|
| 1 | Salvador, Bahia | Analista da Qualidade, Coordenador da Qualidade, Analista de Processos, Auditor Interno, Gestão da Qualidade, ISO 9001 |
| 2 | Salvador, Bahia | Analista de Requisitos, Analista de Negócios, Analista de Testes, Analista Funcional, Analista de Sistemas, Product Owner, Quality Assurance |
| 3 | Brasil + filtro Remoto | Mesmos termos do Bloco 2 |
| 4 | Petrolina, Pernambuco | Mesmos termos dos Blocos 1 e 2 |
| 5 | Juazeiro, Bahia | Mesmos termos dos Blocos 1 e 2 |

### 8.4 Limpar o histórico

Para que todas as vagas voltem a aparecer, apagar:

```
C:\Users\SEU_USUARIO\AgenteVagas\vagas_vistas.json
```

O controle de candidaturas não é afetado: as vagas que já estão nele, com os
envios e retornos que você registrou, continuam lá.

Alternativa sem apagar nada: usar `--refazer` na execução.

### 8.5 Segunda opinião da IA nos títulos

A pontuação por palavra-chave só reconhece termos cadastrados. Para não perder
vagas com títulos diferentes (ex.: "Especialista em Excelência Operacional"),
os títulos que ficaram **abaixo da nota mínima** são enviados ao Gemini numa
única consulta por execução. A IA responde se cada um combina com o currículo A,
com o B ou com nenhum. As aprovadas entram no relatório com a nota mínima e a
indicação "IA: título compatível com o perfil".

- Usa a mesma chave gratuita do Adaptador (`config_adaptador.txt`). Uma
  consulta por dia consome muito pouco da cota gratuita.
- Se a IA falhar ou estiver fora do ar, o agente segue só com as
  palavras-chave e registra um aviso.
- Para desligar: escrever `ia_titulos = nao` no `config.txt`.
- Só o título da vaga é enviado à IA. Nenhum dado pessoal sai do computador.

Se a IA aprovar vagas que não fazem sentido, desligar a opção e cadastrar os
termos certos no `palavras_chave.txt`.

### 8.6 Controle de candidaturas

O `controle_candidaturas.xlsx` é refeito a cada execução a partir de quatro
fontes: o histórico de vagas, o "Enviado?" das planilhas do dia, o
`historico.csv` do Adaptador e o próprio controle. Uma vaga que já está no
controle nunca é removida, mesmo que saia das outras fontes.

| Aba | Conteúdo |
|---|---|
| Candidaturas | Uma linha por vaga. Colunas amarelas são suas: Enviado?, Data do envio, Retorno, Observações |
| Indicadores | Funil, taxa de retorno por currículo, efeito da adaptação, retorno por faixa de nota, últimas 8 semanas |
| Como usar | Instruções e exemplo de linha preenchida |

Regras:
- O agente **nunca apaga** o que você escreve nas colunas amarelas.
- Se **Enviado?** for "Sim" e a **Data do envio** estiver em branco, o agente
  preenche com a data da planilha do dia (ou com a data da execução).
- Uma adaptação do Adaptador é ligada à vaga pela empresa e pelo cargo. Se a
  vaga não veio do agente, entra como linha nova com origem "Adaptador".
- Vaga encontrada por fora: acrescentar uma linha no fim da aba Candidaturas,
  com Data, Vaga e Empresa, deixando a coluna ID em branco.
- Não renomear as colunas nem a aba Candidaturas: o agente as procura pelo nome.
- A coluna ID fica oculta. Ela identifica a vaga e não deve ser alterada.

### 8.7 Testes automáticos

O programa tem testes automáticos (pasta `tests` no repositório do GitHub), que
verificam a leitura dos e-mails, a pontuação, o controle e a IA, usando um
e-mail de exemplo com dados fictícios. Eles rodam sozinhos no GitHub a cada
atualização do código.

Quando o LinkedIn mudar o formato dos e-mails (9.2), salvar um alerta real com
`python inspecionar.py --salvar N --anonimizar` e usá-lo como novo exemplo dos
testes antes de ajustar o código.

---

## 9. Tratamento de falhas

### 9.1 Verificação inicial

Diante de qualquer suspeita de falha, nesta ordem:

1. Abrir `C:\AgenteVagas\registro.txt` e ler a última entrada
2. Rodar `python diagnostico.py`
3. Se necessário, rodar `python inspecionar.py`

### 9.2 Tabela de falhas conhecidas

| Sintoma | Causa | Ação |
|---|---|---|
| `python não é reconhecido` | Python instalado sem "Add to PATH" | Reinstalar marcando a opção |
| `No module named 'openpyxl'` | Biblioteca ausente | `pip install openpyxl` |
| `No module named 'configuracao'` (ou outro módulo) | Faltou copiar algum `.py` para a pasta | Conferir a lista da seção 5.1 |
| `Nao achei o arquivo config.txt` | Arquivo ausente ou fora da pasta | Recriar em `C:\AgenteVagas` |
| `Nao achei o arquivo de palavras-chave` | `palavras_chave.txt` ausente | Copiar de volta para `C:\AgenteVagas` |
| `ERRO: palavras_chave.txt, linha N` | Formato errado na edição | Corrigir a linha indicada (ver 8.1) |
| `ERRO: config.txt, linha N` | Valor inválido no config | Corrigir a linha indicada (ver 5.6) |
| `Login recusado pelo Gmail` | Senha de app inválida ou revogada | Ver 9.3 |
| `Nao consegui conectar ao Gmail` | Sem internet ou porta 993 bloqueada | Verificar rede e firewall |
| `0 e-mail(s) encontrado(s)` | Alertas não chegaram ou vão para outro endereço | Ver 9.4 |
| `0 link(s) de vaga` com e-mails presentes | Formato dos e-mails mudou | Rodar `inspecionar.py --salvar N --anonimizar` e solicitar ajuste do código |
| Planilha do dia em branco | Nenhuma vaga nova — comportamento correto | Nenhuma. Usar `--refazer` para conferir |
| Vaga irrelevante com nota alta | Palavra-chave casando indevidamente | Ajustar listas conforme 8.1 |
| Vaga irrelevante marcada "IA" | A IA aprovou um título fora do perfil | Desligar com `ia_titulos = nao` (8.5) |
| `[aviso] A IA nao respondeu` | Gemini fora do ar ou limite gratuito | Nenhuma. O agente seguiu só com as palavras-chave |
| `[aviso] Controle ... aberto no Excel` | Controle aberto na hora da execução | Fechar o arquivo. Ele é atualizado na próxima execução, sem perda |
| Tarefa não executa às 9h | PC desligado ou configuração do Agendador | Ver 9.5 |
| `PermissionError` ao gerar a planilha | Planilha aberta no Excel | O programa salva com outro nome automaticamente |
| Execução trava sem mensagem | Conexão com o Gmail sem resposta | Encerrar com `Ctrl+C`. Há tempo limite de 45s |
| Correção aplicada não surtiu efeito | Arquivo antigo ainda em uso | Conferir a versão na primeira linha (ver 5.5) |
| Envio marcado não aparece no controle | O controle é atualizado na execução seguinte | Rodar `python agente_vagas.py --so-controle` |
| Controle perdeu algum dado | Arquivo editado com colunas renomeadas | Recuperar de `controle_candidaturas_backup.xlsx` |

### 9.3 Recuperar a senha de app

Perder a senha é normal — ela é descartável.

1. Acessar `myaccount.google.com/apppasswords`
2. Clicar na lixeira ao lado de "Agente de Vagas" para revogar a antiga
3. Criar nova com o mesmo nome e **copiar imediatamente** — o Google exibe
   uma única vez
4. Abrir `config.txt` e substituir na linha `senha = `
5. Validar com `python diagnostico.py` — deve retornar `Login: OK`

### 9.4 Alertas não chegam

1. Verificar em **LinkedIn → Vagas → Alertas de vaga** se os alertas existem
2. Verificar em **Configurações → Login e segurança → Endereços de e-mail**
   se o Gmail está **confirmado e definido como principal**
3. Se houver confirmação pendente:
   - Fazer login completo no LinkedIn pelo **navegador do computador**,
     incluindo o código da verificação em duas etapas
   - **Só então** clicar no link de confirmação, na mesma janela
   - A ordem inversa faz a confirmação se perder
4. Alternativa, se a confirmação não completar: criar encaminhamento
   automático no e-mail antigo, de tudo que vem de `linkedin.com` para o
   Gmail do agente. O agente funciona igual, pois o remetente não muda.

### 9.5 Tarefa agendada não executa

1. Abrir o Agendador de Tarefas e localizar "Agente de Vagas"
2. Verificar a coluna **Resultado da última execução** — o esperado é `(0x0)`
3. Botão direito → **Executar** para testar imediatamente
4. Conferir se as configurações da seção 5.3 continuam aplicadas
5. Testar o `.bat` isoladamente com dois cliques em `rodar_agente.bat`

### 9.6 Voltar para a versão anterior (1.4)

Se a versão 1.5.1 apresentar um problema que impeça o uso:

1. Abrir `C:\AgenteVagas\backup_v1.4`
2. Copiar `agente_vagas.py`, `diagnostico.py` e `inspecionar.py` para
   `C:\AgenteVagas`, substituindo os atuais
3. Rodar `testar_agora.bat` e conferir `=== Agente de Vagas v1.4 ===`

A versão 1.4 usa as palavras-chave de dentro do `agente_vagas.py` e ignora os
módulos novos e o controle. O `config.txt` e o histórico servem para as duas.

---

## 10. Segurança

- O `config.txt` contém a senha de app. **Não compartilhar, não enviar por
  e-mail, não armazenar em nuvem pública.**
- Ao solicitar suporte, enviar os arquivos `.py`, o `palavras_chave.txt` ou o
  `registro.txt`, que não contêm dados sensíveis.
- Em caso de suspeita de vazamento, revogar a senha conforme 9.3. A revogação
  é imediata e não afeta a conta.
- O uso de uma conta de e-mail dedicada limita o impacto de um eventual
  vazamento a uma caixa que só recebe alertas de vaga.
- Os links gravados nas planilhas e no histórico são a forma curta
  (`linkedin.com/jobs/view/NÚMERO`). Os links originais dos e-mails trazem
  tokens pessoais de acesso e não são mais guardados.
- Para a IA vão apenas títulos de vagas, sem nenhum dado pessoal.

---

## 11. Limitações conhecidas

1. **O agente depende dos alertas.** Se um alerta não existir ou estiver
   inativo, aquelas vagas nunca aparecerão no relatório.
2. **A extração depende do formato dos e-mails do LinkedIn.** Se a plataforma
   mudar o layout, a leitura pode falhar e exigir ajuste no código. Os testes
   automáticos (8.7) ajudam a detectar e corrigir.
3. **A pontuação é por palavra-chave**, não por compreensão do texto. A segunda
   opinião da IA (8.5) reduz esse problema para os títulos de nota baixa, mas
   o alerta do LinkedIn traz só o título, não a descrição da vaga. Vagas
   marcadas "IA" devem ser conferidas antes da candidatura.
4. **O PC precisa estar ligado.** Se estiver desligado às 9h, a execução
   ocorre na próxima vez que for ligado.
5. **Somente LinkedIn.** O suporte ao Catho e ao Outlook foi removido do código
   na v1.5, por não estar em uso.
6. **Duas execuções no mesmo dia geram dois arquivos.** A primeira cria
   `vagas_AAAA-MM-DD.xlsx`; se outra execução no mesmo dia encontrar vagas novas,
   elas vão para `vagas_AAAA-MM-DD_HHMMSS.xlsx`. A planilha da manhã nunca é
   sobrescrita, e o "Enviado?" das duas vai para o controle.
7. **Ligação das adaptações por semelhança.** O Adaptador nem sempre recebe o
   link da vaga. Quando falta, a ligação com a vaga do agente é feita pela
   empresa e pelo cargo; nomes muito diferentes podem gerar uma linha separada.
8. **Notas das vagas antigas.** Vagas registradas antes da v1.5 tiveram a nota
   recalculada só pelo título, por isso podem aparecer com nota menor no controle.

---

## 12. Histórico de revisões

| Versão | Data | Alteração |
|---|---|---|
| 1.0 | ago/2026 | Emissão inicial. Agente em operação, com agendamento diário às 9h. |
| 1.1 | ago/2026 | Atualizado após a primeira calibragem com dados reais (110 vagas). Inclui as listas de palavras-chave vigentes, a referência de nota mínima, a verificação de versão do programa e cinco falhas novas na seção 9.2. Programa na versão 1.4. |
| 1.2 | out/2026 | Programa na versão 1.5.1: palavras-chave em arquivo próprio, segunda opinião da IA, controle de candidaturas com indicadores, testes automáticos e código reorganizado em módulos. Novas seções 5.6, 6.1, 8.5, 8.6, 8.7 e 9.6; tabela 9.2 e limitações atualizadas. |

---

## 13. Registro de melhorias aplicadas durante a implantação

Mantido como referência para futuras manutenções.

| Problema encontrado | Solução adotada |
|---|---|
| Raspagem dos sites violaria os termos de uso | Uso dos alertas oficiais por e-mail |
| Outlook pessoal deixou de aceitar senha comum | Migração para Gmail com senha de app |
| Filtro do Gmail arquivava sem aplicar marcador | Filtro removido; leitura direta da caixa de entrada |
| Empresa e local vinham juntos e com HTML | Separação pelo caractere `·` e limpeza de tags cortadas |
| Termos curtos casavam dentro de outras palavras (`ti` em "logística") | Comparação por palavra inteira, com delimitadores |
| Senha perdida a cada atualização do programa | Configuração externalizada para o `config.txt` |
| Erros invisíveis na execução automática | Registro gravado em `registro.txt` |
| Registro ficava vazio quando a execução travava | Saída sem buffer (`python -u` e `line_buffering`) |
| Execução travava indefinidamente esperando o Gmail | Tempo limite de 45 segundos na conexão |
| Vagas do perfil não reconhecidas por variação de escrita ("Analista Negócios" sem o "de", "Negócio" no singular, "QA" isolado) | Listas ampliadas com as formas reais encontradas nas vagas |
| Termo genérico no peso 3 trazia vagas fora do alvo | `analista de operacoes` e `funcional` isolado removidos |
| Planilha aberta no Excel interrompia a execução | Gravação com nome alternativo em caso de bloqueio |
| Segunda execução do dia sobrescrevia o relatório com um vazio | Execução sem vagas novas não grava mais arquivo |
| Segunda execução do dia com vagas novas sobrescrevia a planilha da manhã (perdendo o "Enviado?" marcado) | Vagas novas da segunda execução vão para um arquivo com o horário no nome |
| Dúvida sobre qual versão estava instalada | Número de versão exibido na primeira linha |
| `inspecionar.py` falhava no login (lia a senha antes de carregar o config) | Ferramentas de suporte passaram a carregar o `config.txt` pelo mesmo módulo do agente |
| Editar palavras-chave dentro do `.py` causava `SyntaxError` | Listas movidas para `palavras_chave.txt`, com acentos e mensagem de erro indicando a linha |
| Vagas com título fora das listas ficavam de fora | Segunda opinião da IA para os títulos de nota baixa |
| Envios espalhados em planilhas diárias, sem visão do resultado | Controle único de candidaturas com indicadores (PDCA) |
| Adaptações do currículo sem ligação com as vagas | Controle liga o `historico.csv` do Adaptador às vagas por empresa e cargo |
| Links das vagas guardavam tokens pessoais do e-mail | Gravação só do link curto da vaga |
| `--refazer` substituía o histórico inteiro pelas vagas do dia (e, no teste da v1.5, trocava a data das vagas antigas) | Histórico preservado; `--refazer` só ignora o filtro de repetidas e mantém a data em que a vaga apareceu pela primeira vez |
| Leitura marcava os alertas como lidos no Gmail | Leitura em modo somente consulta (`BODY.PEEK`) |
| Código antigo do Catho e do Outlook dificultava a manutenção | Removido; código separado em módulos com uma função cada |
| Mudança no layout dos e-mails só era percebida com o relatório vazio | Testes automáticos com e-mail de exemplo, rodando no GitHub a cada atualização |
| No primeiro teste da v1.5, a IA aprovou consultorias de ERP (SAP, TOTVS, NetSuite) e controle de qualidade de produção | Instrução da IA reescrita com o perfil do candidato, regras de exclusão e exemplos reais de aprovação e recusa |
| A mesma vaga aparecia várias vezes, publicada com números diferentes (88 repetições no histórico) | Anúncio com mesmo cargo e empresa dos últimos 30 dias é tratado como repetido |
| Em parte dos alertas, o título engolia empresa e local, e a empresa ficava vazia (20 vagas no histórico) | Título separado pelos blocos internos do link; salário e selos retirados do local |

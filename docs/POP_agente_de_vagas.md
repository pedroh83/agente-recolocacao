# POP — Agente de Vagas

**Procedimento Operacional Padrão**
Responsável: Pedro Henrique Simas
Versão: 1.1 — agosto/2026
Versão do programa: 1.4

---

## 1. Objetivo

Automatizar o levantamento diário de vagas de emprego, entregando toda manhã
uma planilha com as vagas novas, ordenadas por compatibilidade com os
currículos, indicando qual versão enviar e com o link de candidatura pronto.

## 2. Escopo

Aplica-se ao computador pessoal (Windows), à conta de e-mail
`seu-email-de-alertas@gmail.com` e aos alertas de vaga configurados no LinkedIn.

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

## 4. Como funciona

```
Madrugada    LinkedIn processa os alertas e envia os e-mails
                          ↓
             E-mails chegam na caixa de entrada do Gmail
                          ↓
9h           Agendador do Windows dispara o rodar_agente.bat
                          ↓
             1. Conecta no Gmail e lê os e-mails dos últimos 2 dias
             2. Extrai título, empresa, local e link de cada vaga
             3. Descarta as que já estão no histórico
             4. Pontua contra as palavras-chave dos currículos A e B
             5. Gera a planilha ordenada por nota
             6. Grava as vagas do dia no histórico
                          ↓
             A planilha abre automaticamente na tela
```

---

## 5. Configuração instalada

### 5.1 Localização dos arquivos

**Pasta do agente:** `C:\AgenteVagas`

| Arquivo | Função |
|---|---|
| `agente_vagas.py` | Programa principal |
| `config.txt` | **Configuração e senha** — não compartilhar |
| `rodar_agente.bat` | Execução automática (usado pelo Agendador) |
| `testar_agora.bat` | Execução manual, mostra tudo na tela |
| `diagnostico.py` | Verifica a conexão e o conteúdo da conta |
| `inspecionar.py` | Mostra o conteúdo de cada e-mail |
| `registro.txt` | Log das execuções automáticas |
| `POP_AGENTE_DE_VAGAS.md` | Este documento — **manter cópia fora desta pasta** |

**Pasta de saída:** `C:\Users\SEU_USUARIO\AgenteVagas`

| Arquivo | Função |
|---|---|
| `vagas_AAAA-MM-DD.xlsx` | Relatório do dia |
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

- Python 3.x, instalado com a opção **Add python.exe to PATH**
- Biblioteca `openpyxl` (`pip install openpyxl`)

### 5.5 Verificar a versão em uso

Ao rodar, a primeira linha exibe a versão:

```
=== Agente de Vagas v1.4 ===
```

Conferir sempre após uma atualização. Se o número não mudar, o arquivo antigo
ainda está em uso — o Windows costuma salvar a cópia nova como
`agente_vagas (1).py` sem substituir. Procedimento correto: **apagar o
arquivo antigo primeiro**, depois mover o novo.

---

## 6. Rotina diária de uso

1. A planilha abre sozinha por volta das 9h
2. As vagas vêm ordenadas por nota, com as melhores no topo, em verde
3. Para cada vaga de interesse:
   - Consultar a coluna **Currículo** — indica se envia a versão A ou B
   - Clicar em **Candidatar-se** — abre a vaga no LinkedIn
   - Após se candidatar, escrever "sim" na coluna **Enviado?**
4. Salvar a planilha ao final — ela serve como registro de candidaturas

**Observação:** planilha vazia significa que não houve vaga nova desde a
última execução. Não é falha.

---

## 7. Operação manual

| Necessidade | Comando |
|---|---|
| Rodar agora, vendo o resultado na tela | `testar_agora.bat` (dois cliques) |
| Rodar mostrando o detalhamento | `python agente_vagas.py --debug` |
| Ver de novo vagas já exibidas | `python agente_vagas.py --refazer` |
| Buscar mais dias para trás (após viagem) | `python agente_vagas.py --dias 10` |
| Verificar a conexão com o Gmail | `python diagnostico.py` |
| Ver o conteúdo dos e-mails recebidos | `python inspecionar.py` |

Todos executados a partir do Prompt de Comando, após:

```
cd C:\AgenteVagas
```

---

## 8. Manutenção

### 8.1 Ajustar as palavras-chave

As listas ficam no início do `agente_vagas.py`. Cada currículo tem termos de
**peso 3** (fortes, quase definem a vaga) e **peso 1** (apoio). A nota final é
a soma dos pesos encontrados, mais 2 pontos se a vaga citar Salvador, Bahia,
remoto, home office ou híbrido.

**Procedimento:**

1. Abrir `agente_vagas.py` no Bloco de Notas
2. `Ctrl+F` → procurar `CURRICULO_A` ou `CURRICULO_B`
3. Acrescentar o termo na linha de peso `3:` ou `1:`
4. Escrever em **minúsculo e sem acento** — o programa normaliza o texto antes
   de comparar
5. Manter aspas e vírgulas
6. Salvar com `Ctrl+S` e testar com `python agente_vagas.py --debug --refazer`

**Regras aprendidas na calibragem (importantes):**

- A comparação é por **palavra inteira**. Termos curtos como `qa`, `ti` ou `5s`
  são seguros: não casam dentro de outras palavras.
- As vagas variam a redação. Cadastrar as formas que aparecem de fato:
  `analista de negocios`, `analista de negocio`, `analista negocios`.
- Termo genérico demais no peso 3 contamina o resultado. Exemplo removido:
  `analista de operacoes` trazia vagas de operação sem relação com qualidade.
- Testar sempre com `--refazer` antes de dar por encerrado.

**Conteúdo atual das listas:**

*Currículo A — peso 3:* iso 9001, iso 14001, iso 45001, sistema de gestao da
qualidade, sistema de gestao integrado, sgq, sgi, auditoria interna, auditor
interno, auditora interna, analista da qualidade, analista de qualidade,
coordenador da qualidade, coordenadora da qualidade, gestao da qualidade,
quality management, mapeamento de processos, melhoria continua, controles
internos, analista de processos, analista processos, gestao de processos,
automacao de processos, consultor de processos, consultora de processos,
processos e sistemas, excelencia operacional

*Currículo A — peso 1:* qualidade, auditoria, processos, processo, pdca, nao
conformidade, acao corretiva, indicadores, kpi, bpmn, procedimento
operacional, pop, 5s, lean, meio ambiente, saude e seguranca, sst,
certificacao, conformidade, compliance, integridade, riscos, governanca,
padronizacao, itil

*Currículo B — peso 3:* levantamento de requisitos, analista de requisitos,
analista requisitos, analista de negocios, analista de negocio, analista
negocios, business analyst, consultor de negocios, consultora de negocios,
analise de negocios, analista de implantacao, product owner, product
ownership, analista funcional, analista de sistemas, analista sistemas,
analista de testes, analista de qa, analista qa, qa, quality assurance,
quality engineer, quality analyst, engenheiro de qa, engenheira de qa, teste
manual, testes manuais, homologacao, analista de solucoes

*Currículo B — peso 1:* scrum, agil, kanban, jira, sprint, backlog, user
story, historia de usuario, cerimonias, uat, caso de teste, sql, requisitos,
documentacao, stakeholder, product, dados, inteligencia artificial, automacao,
sistemas, ti, negocio, negocios, implantacao, analytics, digitalizacao,
transformacao digital, no code, low code

*Descarte (vaga eliminada se o título contiver):* estagio, estagiario, jovem
aprendiz, aprendiz, vendedor, vendas externas, motorista, operador de caixa

### 8.2 Ajustar o filtro de nota

Editar `config.txt`:

```
nota_minima = 3
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

### 8.4 Limpar o histórico

Para que todas as vagas voltem a aparecer, apagar:

```
C:\Users\SEU_USUARIO\AgenteVagas\vagas_vistas.json
```

Alternativa sem apagar nada: usar `--refazer` na execução.

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
| `Nao achei o arquivo config.txt` | Arquivo ausente ou fora da pasta | Recriar em `C:\AgenteVagas` |
| `Login recusado pelo Gmail` | Senha de app inválida ou revogada | Ver 9.3 |
| `Nao consegui conectar ao Gmail` | Sem internet ou porta 993 bloqueada | Verificar rede e firewall |
| `0 e-mail(s) encontrado(s)` | Alertas não chegaram ou vão para outro endereço | Ver 9.4 |
| `0 link(s) de vaga` com e-mails presentes | Formato dos e-mails mudou | Rodar `inspecionar.py --salvar N` e solicitar ajuste do código |
| Planilha em branco | Nenhuma vaga nova — comportamento correto | Nenhuma. Usar `--refazer` para conferir |
| `SyntaxError` | Aspa ou vírgula removida na edição | Verificar a linha indicada no erro |
| Vaga irrelevante com nota alta | Palavra-chave casando indevidamente | Ajustar listas conforme 8.1 |
| Tarefa não executa às 9h | PC desligado ou configuração do Agendador | Ver 9.5 |
| `PermissionError` ao gerar a planilha | Planilha aberta no Excel | Fechar o Excel. A partir da v1.2 o programa salva com outro nome automaticamente |
| Execução trava sem mensagem | Conexão com o Gmail sem resposta | Encerrar com `Ctrl+C`. A v1.2 tem tempo limite de 45s |
| Registro com entrada vazia, sem resultado nem erro | Versão anterior à 1.2, que acumulava a saída em memória | Atualizar o programa; o `.bat` usa `python -u` |
| Correção aplicada não surtiu efeito | Arquivo antigo ainda em uso | Conferir a versão na primeira linha (ver 5.5) |
| Relatório do dia ficou vazio após segunda execução | Comportamento anterior à v1.4, que sobrescrevia | Atualizar. Recuperar com `--refazer` |

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

---

## 10. Segurança

- O `config.txt` contém a senha de app. **Não compartilhar, não enviar por
  e-mail, não armazenar em nuvem pública.**
- Ao solicitar suporte, enviar o `agente_vagas.py` ou o `registro.txt`, que
  não contêm dados sensíveis.
- Em caso de suspeita de vazamento, revogar a senha conforme 9.3. A revogação
  é imediata e não afeta a conta.
- O uso de uma conta de e-mail dedicada limita o impacto de um eventual
  vazamento a uma caixa que só recebe alertas de vaga.

---

## 11. Limitações conhecidas

1. **O agente depende dos alertas.** Se um alerta não existir ou estiver
   inativo, aquelas vagas nunca aparecerão no relatório.
2. **A extração depende do formato dos e-mails do LinkedIn.** Se a plataforma
   mudar o layout, a leitura pode falhar e exigir ajuste no código.
3. **A pontuação é por palavra-chave**, não por compreensão do texto. Vagas
   descritas com termos fora das listas receberão nota baixa. As listas foram
   calibradas em agosto/2026 sobre 110 vagas reais e devem ser revisadas
   periodicamente conforme 8.1.
4. **O PC precisa estar ligado.** Se estiver desligado às 9h, a execução
   ocorre na próxima vez que for ligado.
5. **Somente LinkedIn.** O Catho foi retirado do escopo, embora o código
   ainda reconheça seus links caso os alertas sejam criados no futuro.
6. **O relatório do dia usa nome fixo** (`vagas_AAAA-MM-DD.xlsx`). Duas
   execuções no mesmo dia sobrescrevem o arquivo. Ao receber um relatório
   grande que será trabalhado por vários dias, salvar uma cópia com outro nome.

---

## 12. Histórico de revisões

| Versão | Data | Alteração |
|---|---|---|
| 1.0 | ago/2026 | Emissão inicial. Agente em operação, com agendamento diário às 9h. |
| 1.1 | ago/2026 | Atualizado após a primeira calibragem com dados reais (110 vagas). Inclui as listas de palavras-chave vigentes, a referência de nota mínima, a verificação de versão do programa e cinco falhas novas na seção 9.2. Programa na versão 1.4. |

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
| Dúvida sobre qual versão estava instalada | Número de versão exibido na primeira linha |

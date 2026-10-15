# Adaptador de Currículo por Vaga — Passo a passo

> Instalação, uso e revisão · Pedro Henrique Simas · outubro de 2026
> Pasta do programa: `C:\AgenteVagas\Adaptador`

O Adaptador pega a descrição de uma vaga, escolhe entre seus currículos **A (Qualidade & Processos)** e **B (Requisitos & Tecnologia)** o que mais combina e cria um currículo em Word ajustado àquela vaga, mais um relatório. Quem reescreve é a IA gratuita do Google (Gemini), que só pode usar o que já está nos seus currículos.

Você instala **uma vez só** (Parte 1) e depois usa em **3 passos** a cada vaga (Parte 2). Sempre revise antes de enviar (Parte 3).

---

## Sumário

1. [Como funciona](#como-funciona)
2. [Antes de começar](#antes-de-começar)
3. [Parte 1 — Instalar (uma vez só)](#parte-1--instalar-uma-vez-só)
4. [Parte 2 — Usar a cada vaga (3 passos)](#parte-2--usar-a-cada-vaga-3-passos)
5. [Parte 3 — Revisar antes de enviar](#parte-3--revisar-antes-de-enviar)
6. [Parte 4 — Se algo der errado](#parte-4--se-algo-der-errado)
7. [Parte 5 — Manutenção e ajustes](#parte-5--manutenção-e-ajustes)
8. [Extras (opcional)](#extras-opcional)
9. [Cuidados](#cuidados)

---

## Como funciona

1. **Vaga:** você copia a descrição no LinkedIn (Ctrl + C) e abre o atalho **Adaptar Curriculo**.
2. **Pontuação local:** o programa mede quantos % dos termos da vaga aparecem nos currículos A e B.
3. **IA Gemini:** escolhe A ou B e reescreve o currículo para a vaga, seguindo regras de veracidade.
4. **Correções automáticas:** passa para o passado os verbos de empregos que já terminaram.
5. **Verificação:** confere empresas, números, cursos, idiomas e verbos como "liderei" contra os currículos base.
6. **Resultado:** a pasta **saida** abre com o currículo (.docx) e o relatório (.txt). Você revisa e envia.

---

## Antes de começar

- [ ] O computador com Windows onde já roda o Agente de Vagas
- [ ] Seu login do Google (Gmail) para pegar a chave gratuita
- [ ] Seus currículos A e B em Word (.docx). Se estiverem em PDF, abra no Word e use **Arquivo → Salvar como → Documento do Word**
- [ ] O arquivo **Adaptador_Curriculo.zip** (se a pasta `C:\AgenteVagas\Adaptador` ainda não existir)

---

## Parte 1 — Instalar (uma vez só)

### A) Baixar e extrair o pacote

1. Baixe o **Adaptador_Curriculo.zip**. Ele vai para a pasta **Downloads**.
2. Aperte **Windows + E** e clique em **Downloads**.
3. Clique com o **botão direito** no zip → **Extrair tudo...**
4. Clique em **Procurar...**, escolha **Este Computador → Disco Local (C:) → AgenteVagas** → **Selecionar pasta** → **Extrair**.
5. Confira se existe a pasta `C:\AgenteVagas\Adaptador`.

> **Importante:** depois de instalado, não mude a pasta Adaptador de lugar. O atalho da Área de Trabalho aponta para ela.

### B) Rodar o instalador

1. Abra `C:\AgenteVagas\Adaptador` e dê **dois cliques** em **instalar**.
2. Se aparecer a tela azul **"O Windows protegeu o computador"**: **Mais informações → Executar assim mesmo**.
3. Espere aparecer **[1/2] Instalando as bibliotecas... OK**. Em seguida o assistente abre sozinho.

Se aparecer **"O Python não foi encontrado"**: a página de download abre. Instale o Python marcando a caixa **"Add python.exe to PATH"** na primeira tela e rode o **instalar** de novo.

### C) Criar a chave gratuita do Google (Etapa 2 do assistente)

1. Aperte **Enter** na janela preta. O navegador abre o Google AI Studio.
2. Entre com sua conta Google e aceite os termos.
3. Clique em **Criar chave de API** (Create API key).
4. Se o campo do projeto mostrar **"No Cloud Projects Available"** e o botão **Criar chave** ficar cinza, crie um projeto antes:
   1. Clique em **Cancelar**. No menu da esquerda, vá em **Projetos** e use **Criar projeto**, se existir.
   2. Se não existir, abra **console.cloud.google.com/projectcreate**, dê o nome `Adaptador` e clique em **Criar**.
   3. Volte ao AI Studio → **Projetos → Importar projetos**, marque **Adaptador** → **Importar**.
   4. Volte em **Chaves de API**, escolha o projeto **Adaptador** e clique em **Criar chave**.
5. Clique no **ícone de copiar** ao lado da chave (ela pode começar com `AIza` ou `AQ.`).
6. Na janela preta, clique com o **botão direito** (ou Ctrl + V) para colar e aperte **Enter**.
7. Espere a mensagem **"Chave funcionando... Salva!"**.

### D) Seus dados de contato (Etapa 3)

O que aparece entre colchetes `[ ]` é o valor já salvo: aperte **Enter** para manter. Para trocar, digite o valor novo e aperte Enter.

```
  Telefone com DDD, ex.: (71) 99999-9999 [(71) 99999-9999]:
  E-mail [seu.email@exemplo.com]:
```

### E) Seus currículos A e B (Etapa 4)

1. Aperte **Enter**. A pasta **curriculos_base** abre.
2. Copie para ela os dois currículos em Word, com o nome que já têm. Volte à janela preta e aperte **Enter**.
3. Digite o número do currículo **A (Qualidade & Processos)** e aperte **Enter**.
4. Digite o número do currículo **B (Requisitos & Tecnologia)**. Os números ficam iguais nas duas perguntas; o já escolhido aparece marcado "(já escolhido como A)".

### F) Atalho na Área de Trabalho

Quando o assistente perguntar, aperte **Enter** (sim). O ícone **Adaptar Curriculo** aparece na Área de Trabalho (aperte **F5** na Área de Trabalho se não aparecer de imediato).

Se o atalho não aparecer, use um destes jeitos:

- Dê dois cliques em **criar_atalho**, dentro da pasta Adaptador.
- Ou: botão direito em **colar_vaga** → **Mostrar mais opções** → **Enviar para** → **Área de trabalho (criar atalho)**.

> **Área de Trabalho no OneDrive:** se a sua Área de Trabalho estiver no OneDrive, ela fica em `C:\Users\SEU_USUARIO\OneDrive\Área de Trabalho`. É lá que o atalho precisa estar.

### G) Conferir

O assistente termina com a mensagem abaixo. Para conferir tudo a qualquer momento, dê dois cliques em **testar**.

```
------------------------------------------------------------
 TUDO PRONTO!
 Para usar: abra a vaga no LinkedIn, selecione a descrição, Ctrl+C
 e dê dois cliques em 'Adaptar Curriculo' (ou colar_vaga.bat).
------------------------------------------------------------
```

---

## Parte 2 — Usar a cada vaga (3 passos)

1. **Abrir:** abra a vaga no LinkedIn e clique em **Exibir mais** para mostrar a descrição completa.
2. **Copiar:** selecione do título da vaga até o fim da descrição (texto marcado em azul) e aperte **Ctrl + C**.
3. **Adaptar:** dê dois cliques em **Adaptar Curriculo** na Área de Trabalho.

Confira a linha **"Texto copiado começa com:"**. Ela deve mostrar o início da vaga (ex.: "Logo da empresa..."). A janela fica assim:

```
Texto copiado começa com: "Logo da empresa Empresa Exemplo."
Vaga copiada salva em vaga_20261006_003047.txt
[1/1] vaga_20261006_003047.txt
  Pontuação local: A=25%  B=28%
  Consultando a IA (Gemini)...
    ... ainda aguardando a IA (20s)
  OK  Versão A | compatibilidade 35% -> 85%
Concluído: 1 currículo(s) gerado(s)
```

> **Não copie nada entre o passo 2 e o passo 3.** Se você copiar outra coisa (por exemplo, o texto da janela preta para mandar no chat), o Adaptador usa esse texto no lugar da vaga. Para mostrar a janela para alguém, tire um print com **Windows + Shift + S** em vez de copiar o texto. Se o texto copiado for da janela preta, o programa recusa e avisa.

### Vaga que ficou pendente

Se der erro, a vaga fica guardada na pasta `vagas`. Para tentar de novo, dê dois cliques em **adaptar_vagas** (e **não** no atalho, que pegaria o Ctrl+C de novo e criaria uma cópia repetida).

---

## Parte 3 — Revisar antes de enviar

Abra primeiro o **Relatorio** (.txt) e depois o **Curriculo** (.docx) na pasta `saida`.

| Parte do relatório | O que significa | O que fazer |
| --- | --- | --- |
| Versão base escolhida | Se a IA usou o currículo A ou o B | Conferir se faz sentido para a vaga |
| Compatibilidade estimada | Quanto o currículo combina com a vaga, antes e depois | Se o número do adaptado for baixo, pense se vale a candidatura |
| Palavras-chave da vaga | Termos que os filtros automáticos (ATS) procuram | Use alguns na mensagem ao recrutador |
| Lacunas | O que a vaga pede e seu currículo não comprova | Prepare uma resposta honesta para a entrevista |
| Alertas de veracidade | Algo no currículo novo que não está no seu currículo base | **Corrija no Word antes de enviar** |
| Dicas para esta candidatura | Sugestões da IA para a vaga | Use o que fizer sentido |

### Checklist do currículo

- [ ] Cada alerta do relatório foi corrigido (ou confirmado como verdadeiro)
- [ ] Empresas, cargos e datas estão certos
- [ ] Experiências que já terminaram estão no passado ("Conduzi", não "Conduzo")
- [ ] Nada foi "promovido": se você acompanhou, não pode estar "conduzi" ou "liderei"
- [ ] Nenhum nível inventado (ex.: inglês "intermediário")
- [ ] Termos da vaga só aparecem em empregos onde você realmente fez aquilo
- [ ] Telefone, e-mail e LinkedIn no topo estão corretos
- [ ] Se a vaga pedir PDF: **Arquivo → Salvar como → PDF**

### Erros que a IA já cometeu (e o que fazer)

| Erro | Exemplo | Como corrigir |
| --- | --- | --- |
| Termo da vaga no emprego errado | "Monitorei indicadores" num emprego onde isso não foi feito | Voltar à frase do seu currículo base |
| Nível de idioma inventado | "Inglês — intermediário" | Deixar "o que está no currículo base (ex.: o curso feito)" |
| Papel promovido | "Ministrei" quando a base dizia "acompanhei" | Usar o verbo da base (ou atualizar a base, se for verdade) |
| Tempo verbal | "Executo testes" num emprego já encerrado | Passar para o passado (o programa já corrige sozinho) |
| Detalhe importante cortado | "migração da ISO 9001" sem versões | Repor "ISO 9001:2008 para ISO 9001:2015, incluindo recertificação" |

> **Quando o alerta é sobre algo que você fez de verdade:** o currículo gerado não está errado, o currículo base está incompleto. Acrescente a informação em `curriculos_base\curriculo_A.docx` e `curriculo_B.docx` (no Word) para o Adaptador poder usar nas próximas vagas.

Guarde a versão revisada com **CORRIGIDO** no nome, para saber qual enviar.

---

## Parte 4 — Se algo der errado

| Apareceu na tela | O que fazer |
| --- | --- |
| O Python não foi encontrado | Instale o Python marcando **Add python.exe to PATH** e rode o **instalar** de novo |
| Não consegui instalar as bibliotecas | Confira a internet e rode o **instalar** de novo |
| A chave não funcionou / Acesso negado pela API | Rode o **instalar**, responda "s" para trocar a chave e cole uma nova |
| Servidor do Google ocupado / continua ocupado. Trocando de modelo | Normal: o programa tenta outro modelo sozinho. Só espere |
| ... ainda aguardando a IA (20s) | Normal: a IA está escrevendo. Após 2 minutos sem resposta, ele troca de modelo |
| O Google está sobrecarregado agora e nenhum modelo respondeu | Espere alguns minutos e dê dois cliques em **adaptar_vagas** |
| Limite gratuito atingido | Não feche: ele espera e tenta de novo. Se continuar, tente mais tarde |
| O texto copiado NÃO é uma vaga | Volte ao LinkedIn, copie a descrição (Ctrl + C) e abra o atalho de novo |
| A área de transferência está vazia | Você não copiou a vaga. Copie e abra o atalho de novo |
| texto muito curto ... Pulando | Faltou clicar em **Exibir mais** antes de copiar |
| Nenhum currículo base encontrado | Rode o **instalar** e escolha de novo os arquivos A e B |
| Permission denied | O currículo ou o historico.csv está aberto no Word/Excel. Feche e tente de novo |
| Janela abre e fecha / atalho não abre | A pasta Adaptador mudou de lugar. Rode **criar_atalho** de novo |

---

## Parte 5 — Manutenção e ajustes

### Arquivos da pasta Adaptador

| Arquivo / pasta | Para que serve |
| --- | --- |
| instalar | Instala as bibliotecas e abre o assistente (chave, contato, currículos, atalho) |
| Adaptar Curriculo (atalho) / colar_vaga | Adapta a vaga que está no Ctrl + C |
| adaptar_vagas | Processa as vagas guardadas na pasta vagas (pendentes ou salvas à mão) |
| testar | Confere configuração, currículos e conexão com o Google |
| criar_atalho | Recria o atalho na Área de Trabalho |
| config_adaptador.txt | Chave, modelo, nome, contato e idioma |
| curriculos_base | Currículos A e B que a IA usa como fonte da verdade |
| vagas / vagas\processadas | Vagas a processar / já processadas |
| saida | Currículos e relatórios gerados |
| historico.csv | Registro de todas as adaptações (abre no Excel) |

### Ajustes no config_adaptador.txt

Abra com **botão direito → Abrir com → Bloco de Notas**, mude só o texto depois do `=`, salve com **Ctrl + S**.

- **Telefone, e-mail ou LinkedIn:** linha `CONTATO=`, separando as informações com ` | `.
- **Currículo em inglês:** troque para `IDIOMA=inglês` (só se você sustentar a entrevista em inglês).
- **Mais pausa entre vagas:** `PAUSA_ENTRE_VAGAS=10`.

### Atualizar os currículos base

Sempre que fizer um curso, projeto ou emprego novo, atualize `curriculo_A.docx` e `curriculo_B.docx` na pasta `curriculos_base`. Escreva empregos encerrados no passado. O Adaptador só usa o que estiver lá.

---

## Extras (opcional)

### Rotina junto com o Agente de Vagas

1. Às 9h, o Agente gera a planilha do dia com o link de cada vaga e a versão sugerida (A ou B).
2. Você escolhe as vagas que valem a candidatura e faz os 3 passos da Parte 2 para cada uma.
3. Na execução seguinte, o Agente registra cada adaptação no `controle_candidaturas.xlsx`, ligada à vaga correspondente. A aba **Indicadores** mostra se as candidaturas com currículo adaptado trazem mais retorno.

### Várias vagas de uma vez

1. Abra a pasta `Adaptador\vagas`.
2. Para cada vaga: botão direito → **Novo → Documento de Texto**, cole o link na primeira linha e a descrição embaixo, salve.
3. Dê dois cliques em **adaptar_vagas**.

### Rodar sozinho todo dia (Agendador de Tarefas)

1. Menu Iniciar → **Agendador de Tarefas** → **Criar Tarefa Básica**, nome `Adaptador de Currículo`.
2. **Diariamente** → escolha o horário → **Iniciar um programa**.
3. Programa: `C:\AgenteVagas\Adaptador\adaptar_vagas_agendado.bat`. Iniciar em: `C:\AgenteVagas\Adaptador`.
4. Concluir. O andamento fica em `log_adaptador.txt`.

---

## Cuidados

- **Nunca envie sem revisar.** A verificação automática pega muita coisa, mas não tudo.
- **Lacuna não é para esconder:** prepare uma resposta honesta para a entrevista.
- **Chave do Google:** não mostre para ninguém. Se vazar, apague no AI Studio e cadastre outra pelo **instalar**.
- **Privacidade:** no plano gratuito, o Google pode usar o conteúdo enviado para melhorar os produtos. Deixe CPF, RG e endereço completo fora dos currículos base.
- **Leia a vaga:** confira local (presencial ou remoto), nível e salário antes de se candidatar.

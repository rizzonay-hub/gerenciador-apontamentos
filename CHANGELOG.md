# Histórico — Apontamentos de horas

## 29/10/2026 — treinamento em vídeo

**Novo**
- Pasta `treinamento\`: roteiro corporativo em linguagem simples (`roteiro-video.md`), perguntas de fixação em formulário (`formulario-perguntas.md`), montador `montar.py` (capturas reais + narração por IA + ffmpeg) e as capturas com login de funcionário (`capturas-funcionario\`). Os vídeos (`Trilha1-Apontamentos-de-horas.mp4`, `Trilha2-Gerenciamento-do-Time.mp4`), os áudios e o ffmpeg ficam fora do repositório.
- `DemoMain.vb`: `/captura <pasta> funcionario` gera as telas sem os botões de gestor e a tela "acima de 8h".
- `build-demo.rsp`: inclui `ModNiboUsuarios.vb` da entrega de Colaboradores, exigido pelo `UiMecano.vb` atual.

## 23/10/2026 — sem trava; aviso ao gestor

**Mudou** (decisão dos usuários do GM)
- A **trava de 10 h saiu**: `BancoHoras.TRAVA_ATIVA = False`. Acima de 8h00 ou de 10h00 a pessoa é avisada na tela (subtítulo vermelho e mensagem ao gravar) e grava.

**Novo**
- **Aviso ao gestor**: ao gravar um dia acima do limite, `ModApontamento.RegistrarAviso` grava uma linha na `TbTimeAviso` para cada gestor da pessoa (uma por pessoa, dia e gestor; lançar de novo atualiza e volta a "não visto"). Quem é o gestor (`AvisoHoras.EscolherGestores`): líder das equipes ativas da pessoa (TbEquipe.LIDERID, quando existir), senão gestor do departamento, senão diretoria, senão qualquer gestor.
- Botão **Avisos (N)** na lista (só gestor e diretoria; vermelho com pendências) e aba **Avisos** no Gerenciamento do Time (período, "só os que ainda não vi", "Marcar todos como vistos", Exportar). A diretoria vê os de todos.
- Script `preprod\2026-10-23-Avisos-de-Horas.sql`; demo com `avisos.txt`; captura 12; teste sem interface.

## 20/10/2026 — Meus grupos de empresas

**Corrigido**
- Barra de ações da lista em janela estreita: os botões agora descem para quantas linhas a largura pedir, sem sobrepor uns aos outros nem a busca, e a linha divisória acompanha a altura nova (antes ficava desenhada no lugar antigo, cruzando os botões).

**Novo**
- **Grupos de empresas por pessoa** (`frmGruposEmpresas`, `GrupoEmpresas`): quantos grupos quiser, cada um com nome e empresas; montar, renomear e excluir. Botão "Meus grupos" na barra da lista.
- No lançamento com várias empresas, o campo **Grupo** passa as empresas do grupo para a direita de uma vez; "Meus grupos…" abre a janela já com as empresas selecionadas como grupo novo.
- **Primeiro script SQL da entrega**: `preprod\2026-10-20-Grupos-de-Empresas.sql` cria a `TbTimeGrupoEmpresa` (USUARIOID, GRUPO, EMPRESAID). `ModApontamento.GruposDisponiveis` checa a tabela; sem ela, botão e campo somem e nada quebra.
- Demo: `grupos.txt` com dois grupos fictícios; captura `03c-meus-grupos.png`; teste sem interface.

## 19/10/2026 — decisões da diretoria

**Novo**
- **Aviso acima de 8h00**: na janela de lançamento, quando o dia passa de 8h00 o subtítulo fica vermelho ("· acima de 8h00") e, ao gravar, aparece uma mensagem na tela ("Hoje ficou com 9h00, acima de 8h00"). Grava normalmente; a trava de 10h00 continua recusando.
- **Banco de horas adiado**: `BancoHoras.ATIVO = False`. O cartão fica na tela sem número ("em breve"), o clique não abre o extrato e nenhuma mensagem fala em banco. A conta e o extrato ficam prontos para ligar com a chave.
- **Lançar em nome de outra pessoa**, só para gestor e diretoria: campo Colaborador na janela de lançamento (vem com o próprio nome). A trava de 10 h e o subtítulo usam as horas do dia da pessoa escolhida (`ModApontamento.HorasNoDiaDe`); o rodapé diz "Em nome de …" e o lançamento não entra na lista nem nos cartões de quem gravou.

**Decidido** (README, "Decididas em 19/10/2026")
- O banco de horas não alimenta o "Banco de Horas (Saldos)" do RH.
- "Várias empresas" continua abrindo em "cada empresa fica com o tempo inteiro".
- Retrabalho sai de vez.
- O item Consulta do menu Time aponta para a tela nova.
- Motivos do período ficam os dois de hoje; Atestado já é atividade interna.

## 16/10/2026 (noite) — menos texto na tela

- Pasta `preprod\` com os sete `.vb` que entram no Gerenciador (os únicos arquivos que sobem); a demonstração e os `.txt` ficam na raiz.
- Retirados os textos explicativos de todas as telas: subtítulos dos cabeçalhos (Apontamentos, Anotações por período, Gerenciamento), frases de ajuda ("Digite parte do nome…", "Lista do Outros.txt…", "Grava um dia de cada vez…", regra do banco de horas, regra das horas diretas e indiretas, "clique no dia", "clique para ver o extrato", "domingos não entram") e os parênteses das opções ("com código de empresa", "grava uma linha por empresa", "gravado como 1,50 h").
- O subtítulo do lançamento ficou só com os números: "Hoje: 7h00 · faltam 1h00 · com este: 9h00 (1h00 no banco)".
- Capturas refeitas.

## 16/10/2026 (noite) — Gerenciamento do Time completo

**Novo**
- As três abas que estavam "em breve" no `frmGerenciamentoTime`, cada uma com as procedures e a escolha entre elas da tela antiga correspondente:
  - **Sintética** (`frmTimeConsultaAcumulado`): horas por empresa ou atividade, colaborador e mês; filtros de colaborador, empresa ou atividade (a lista acompanha o modo), grupo de empresas e "Mostrar"; "Somar todos os colaboradores por empresa" (o "Acumular por Empresa"). `ModApontamento.ConsultaSintetica`, `Time_GetByEmpresaAcumuladoGeralComSem/Geral/SC/SCTE/CC/CCTE`.
  - **Por funcionário** (`frmTimeFuncionario`, "Consulta Geral"): diretas, indiretas, total e percentuais por pessoa, com departamento e função; duplo clique abre a Sintética da pessoa. `ConsultaPorFuncionario`, `Time_GetByDataTudoHorasCustoDireto/Indireto`.
  - **Por usuário e data** (`frmTimeConsultausuario`): total por pessoa no intervalo, esperado (8h00 dia útil, 4h00 sábado), diferença e situação; quem não lançou entra com 0h00; vermelho fora do esperado. `ConsultaPorUsuarioData`, `Time_GetByEmpresaUsuario`.
- `frmGerenciamentoTime.Abrir(aba)`: os quatro itens do menu Gerenciamento abrem a mesma tela, cada um na sua aba. Exportar CSV em todas as abas, com totais.
- Modelos em `Apontamento.vb`: `FiltroSintetica`, `LinhaSintetica` (com `AcumularPorEmpresa`), `LinhaFuncionario`, `LinhaUsuarioData` (com `HorasEsperadas` e a regra do vermelho). `ColaboradorResumo.Funcao` (TbSenha.FUNCAO).
- Demo: as mesmas consultas sobre o txt; capturas 09, 10 e 11; teste sem interface com 42 verificações cruzando as três abas com a Analítica.

**Não copiado das telas antigas, de propósito** (totais iguais; detalhes no IMPLANTACAO.md): a conversão de horas que trocava o número na consulta por usuário; os nomes fixos no código (diretoria escondida, seis usuários fora da lista); as atividades sem código juntas numa linha só no "Acumular por Empresa".

## 16/10/2026 (tarde) — hierarquia do menu Time dentro da tela

**Novo**
- **Lançar período** (`frmApontamentoPeriodo`): o "Anotações Período" de hoje no visual novo. Gestor ou diretoria escolhe a pessoa, o intervalo e o motivo; a janela mostra antes quantos dias e horas vai gravar (8h00 por dia útil, "Sabado" 4h00, domingo nada, como o `frmTimePeriodo`) e pula os dias em que a pessoa já tem apontamento.
- **Gerenciamento do Time** (`frmGerenciamentoTime`): uma tela com uma aba por consulta de gestão. A **Analítica** está pronta, com as mesmas procedures `Time_GetByEmpresa*` e a mesma escolha entre elas que o `frmTimeConsulta`; Sintética, Por funcionário e Por usuário e data aparecem como "em breve".
- Dois botões na barra da lista, só para gestor e diretoria: "Lançar período" e "Gerenciamento".
- `ModApontamento`: `ListarDe`, `Colaboradores` (TbSenha), `ConsultaAnalitica`, `PodeGerenciar`. Demo: `colaboradores.txt` fictício e lançamentos de mais duas pessoas; a pessoa logada na demo passou a ser gestor.

**Mudou**
- Banco de horas: sábado conta 4h00 de jornada (antes 0), igual à meta mensal e ao "Sabado" do período; só o que passar disso é crédito.

## 16/10/2026 — pronta para pré-produção

**Novo**
- **Rateio entre empresas**: em "Lançar para várias empresas", a opção "Como gravar o tempo?" com *cada empresa fica com o tempo inteiro* (padrão, como o Time de hoje) ou *dividir o tempo entre as empresas*. Divisão em centésimos de hora para a soma fechar exata; recusa quando não dá 0,01 h por empresa. As frases mostram os números reais da pessoa.
- **Trava de 10 h por dia** (`BancoHoras.LIMITE_DIA`): o que já está no dia mais o que vai ser gravado não passa de 10h00. Vale para novo, alteração e várias empresas. O subtítulo da janela mostra o efeito antes de gravar.
- **Banco de horas** (`BancoHoras`, em `Apontamento.vb`): o que passa de 8h00 num dia útil (e todo o tempo de sábado e domingo) é crédito, desde 01/10/2026. Calculado a partir da `TbTime`, sem tabela nova. Quinto cartão BANCO DE HORAS na tela e extrato `frmBancoHoras` com exportação CSV.
- Demo: dados gerados com dois dias acima de 8 h; captura `06-banco-de-horas.png`.

**Corrigido**
- Cartão MENOS DE 8 H: o "+N" dos dias que não cabiam era desenhado em cima do aviso do rodapé.
- Rodapé da janela de lançamento: o texto de atalhos passava por baixo do botão Cancelar.
- "1 apontamentos gravados" quando só uma empresa estava na lista.

**Visual e textos**
- Números grandes dos cartões em Segoe UI negrito (a fonte da marca, Neulis Neue, fica só nos títulos). O "h" segue minúsculo.
- Textos das opções e avisos reescritos em frases curtas, com os números da pessoa, sem decimais na tela.

**Documentação**
- `README.md`, `IMPLANTACAO.md` (passo a passo e roteiro de teste de pré-produção), `CHANGELOG.md` e `.gitignore`.

## 25/09/2026 — primeira entrega (demonstração)

- Tela `frmApontamentos` no visual novo: período, busca, quatro cartões (HOJE, ESTE MÊS com gráfico, MÊS ANTERIOR, MENOS DE 8 H), lista, exportar CSV.
- Janela `frmApontamentoDet`: cliente com filtro conforme digita, atividade interna, tempo em horas e minutos com atalhos, "Salvar e novo", várias empresas (mesmo tempo para cada).
- `ResumoHoras` reúne a regra da trava de horas do `frmMain` num lugar só.
- `ModApontamento` (SQL, mesmas procedures do Time) e `ModApontamentoDemo` (txt, sem banco).
- Saíram: campo Retrabalho (oculto desde 2022), CNPJ digitável (era só exibição), o "Inserir antes de preencher".
- Demo `GM-Apontamentos-Demo.exe` e a mesma tela embutida na demo da tela principal (v12).

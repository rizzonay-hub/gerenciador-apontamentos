# Apontamentos de horas (o "Time") no visual novo do Gerenciador

Entrega de 25/09/2026 para o Gerenciador Mêcano (projeto `Grupo_Mecano`), ampliada em 16/10/2026 (rateio, trava de 10 h, banco de horas). Este é o documento técnico detalhado; a visão geral está em [README.md](README.md), o passo a passo de pré-produção em [IMPLANTACAO.md](IMPLANTACAO.md) e as mudanças por data em [CHANGELOG.md](CHANGELOG.md).

Execute **`GM-Apontamentos-Demo.exe`** para ver a proposta. Não precisa de banco nem de rede: as empresas são **fictícias** (`clientes.txt`), as atividades internas são a cópia do `Outros.txt` da rede, as metas mensais são as reais da `TbHorasMes` (`horasmes.txt`) e os lançamentos são **gerados na primeira abertura** em volta da data de hoje (`apontamentos.txt`), com dois dias pendentes de propósito. O que você gravar na demo fica nesse txt; para recomeçar, apague-o.

A mesma tela também está **embutida na demo da tela principal** (`..\TO-DO_Proposta-Tela-Principal\GM-Tema-Completo-Demo.exe`, versão 12): item "Apontamentos" da seção ROTINA e atalho "Horas" no Acesso rápido.

## 1. O que muda para quem usa

Hoje o Time são duas janelas dos anos 2000: **Cadastro de Tempo** (`frmTimeDet`, menu Apontamentos › Time › Anotações) para lançar, e **Consulta do Time** (`frmTime`, › Consulta) para ver a lista. A proposta junta as duas em uma tela, no padrão das telas novas (login, Colaboradores, Explorer):

| Tela nova | O que tem |
|---|---|
| **Apontamentos de horas** (`frmApontamentos`) | Cabeçalho escuro com logo; período (hoje, semana, mês, mês anterior, 30 dias ou datas livres); busca por empresa, código ou atividade; **quatro cartões** no alto; a lista dos lançamentos; barra de status com contagem e total do período. |
| **Cartões** | HOJE (horas de 8 h, com barra) · ESTE MÊS (apontadas de meta, faltam X, e o **gráfico dia a dia** da proposta da tela principal, com hoje em vermelho e a linha das 8 h) · MÊS ANTERIOR (apontadas de meta: "fechado" em verde ou "faltam X" em vermelho) · **DIAS COM MENOS DE 8 H** nos últimos 20 dias, um chip por dia; clicar no dia abre o lançamento já naquela data. |
| **Novo apontamento** (`frmApontamentoDet`) | Cliente: **a lista filtra conforme você digita**, por qualquer trecho do nome ou pelo código (várias palavras valem: "aurora const"); ↓ ↑ escolhe, **Tab, Enter ou clique seleciona**; ao escolher aparece CNPJ e grupo. **Ou** atividade interna (Outros.txt). Data; tempo em **horas + minutos** com atalhos 15 min / 30 min / 1 h / 2 h / 4 h / 8 h; "Salvar" e "Salvar e novo" (Ctrl+Enter) para lançar o dia inteiro sem fechar a janela. |
| **Várias empresas** | A caixa "Lançar para várias empresas" abre as duas listas (disponíveis › selecionadas, com localizar e marcar todas) e grava uma linha por empresa, como o "SELECIONAR VÁRIAS EMPRESAS" de hoje. Abaixo das listas, "Como gravar o tempo?" com duas opções: **Cada empresa fica com o tempo inteiro** (o comportamento do Time de hoje, que continua sendo o padrão) ou **Dividir o tempo entre as empresas**. Assim que a pessoa informa o tempo e as empresas, as duas frases passam a mostrar os números dela: "Cada empresa fica com 2h00 (4 empresas = 8h00 no total)" e "Dividir as 2h00 entre as 4 empresas (0h30 para cada uma)". O texto acima da lista da direita e a linha abaixo do tempo repetem o resultado, sem decimais. |
| **Trava de 10 h por dia** | O que a pessoa já tem no dia mais o que vai gravar não pode passar de **10h00**. A tela avisa antes ("Hoje: você já tem 8h00 · com este lançamento fica 9h00, 1h00 vai para o banco de horas") e, se passar, não grava: "Você já tem 9h00 hoje. O limite é 10h00 por dia, então cabem mais 1h00." Vale também para alteração e para várias empresas (conta o total das linhas). |
| **Banco de horas** | Quinto cartão, **BANCO DE HORAS**: saldo do que passou de 8h00 por dia desde 01/10/2026 (sábado e domingo entram inteiros). Clicar abre o **extrato** (`frmBancoHoras`): saldo, um chip por mês, a lista dos dias (data, apontado, jornada, quanto foi para o banco) e Exportar CSV. Só leitura: a conta é refeita a partir da TbTime a cada abertura, então editar ou excluir um apontamento corrige o saldo sozinho. |
| **Editar / excluir** | Duplo clique ou Enter abre o lançamento; Del ou o botão "Excluir" pede confirmação; botão direito tem também "Novo apontamento nesta data". |
| **Exportar** | CSV da lista filtrada, com total, abrindo direto no Excel. |

Capturas em `capturas\`: `01-apontamentos.png`, `02-novo-apontamento.png`, `02b-busca-cliente.png` (a lista filtrando com "aurora"), `03-varias-empresas.png`, `03b-dividir-tempo.png` (2 h divididas entre 4 empresas), `04-editar-apontamento.png`, `05-lancar-dia-pendente.png`, `06-banco-de-horas.png` (extrato).

### Trava de 10 h e banco de horas (16/10/2026)

- **Regra**: ninguém aponta mais de 10h00 num dia (`BancoHoras.LIMITE_DIA`). O que passar da jornada (8h00 em dia útil; 0 em sábado e domingo, então o fim de semana entra inteiro) é crédito no banco de horas. A conta começa em **01/10/2026** (`BancoHoras.INICIO_VIGENCIA`, a confirmar com a diretoria); apontamentos anteriores não entram.
- **Nada novo no banco de dados**: o saldo é calculado a partir da TbTime (`Time_GetByData`) toda vez que a tela abre, em `BancoHoras.Calcular` (Apontamento.vb). Por isso editar ou excluir um apontamento corrige o saldo sozinho, e não existe risco de a conta divergir dos lançamentos.
- **Já existe um "Banco de Horas (Saldos)" em produção** (menu Controle do Ponto, telas `frmEntradaSaidaBancoHoras*`, procedures `EntradaSaidaBancoHorasSaldo_*`): é o lançamento manual do RH por competência (usuário, mês, saldo em horas decimais, pago, demitido), sem ligação com a TbTime. Esta entrega **não grava nele**. Se a diretoria quiser que o crédito dos apontamentos alimente aquele saldo mensal, é um passo à parte (uma linha por competência, via `EntradaSaidaBancoHorasSaldo_Insert`), a decidir.
- **O que a tela mostra**: no lançamento, o subtítulo diz quanto a pessoa já tem no dia, quanto falta para 8h00 e o que o lançamento faz ("fica 9h00, 1h00 vai para o banco de horas" ou "passa do limite de 10h00 por dia"). Ao gravar, a mensagem repete "1h00 foi para o banco de horas". Na lista, o cartão BANCO DE HORAS e o extrato.
- **Teste sem interface** (demo): dia com 10h00 + 1h00 recusado; 8h00 + 1h00 gravado com aviso do banco; 9h00 + 3h00 recusado; 9h00 + 1h00 fecha em 10h00; alterar uma linha de 1h00 para 2h00 num dia de 10h00 recusado; saldo recalculado bate com os dias.
- **Fica de fora, por decisão a tomar**: débitos no banco (dia com menos de 8 h hoje é cobrado pela trava de horas, não desconta do banco), compensação/pagamento do saldo (continua no módulo do RH) e feriados (quando o script `2026-09-23-Perfis-e-Trava-Horas.sql` subir, `BancoHoras.JornadaDoDia` passa a ler `TbFeriado` e `JORNADADIARIA`).

**Fontes (decisão de 16/10/2026):** a fonte da marca (Neulis Neue) fica só nos títulos das telas, como nas outras telas novas. Os números grandes dos cartões ("0h00", "138h00") usam Segoe UI em negrito, a mesma dos chips e do corpo, e o "h" continua minúsculo.

### Como o rateio grava (16/10/2026)

- A TbTime guarda `TEMPO` com duas casas, então a divisão é feita em **centésimos de hora** e a soma das linhas é sempre exatamente o tempo digitado: 2,00 h em 3 empresas grava 0,67 + 0,67 + 0,66 (a sobra fica com as primeiras da lista, em ordem de nome). Sem isso, 3 × 0,67 = 2,01 e o dia fecharia com um centésimo a mais. A conta está em `Apontamento.Ratear`.
- Quando as partes diferem por um centésimo, a tela mostra "cerca de 0h40 para cada uma"; quando são iguais, "0h30 para cada uma". Os decimais (0,66 / 0,67) ficam só no banco e nesta nota.
- O Time não grava tempo zerado, então a tela não deixa dividir quando não dá nem 0,01 h para cada empresa (ex.: 5 min entre 20 empresas): avisa "O tempo é pouco para dividir entre 20 empresas: algumas ficariam com zero" e pede para aumentar o tempo ou tirar empresas.
- Cada linha gravada é um apontamento comum da TbTime, sem vínculo com as outras: editar ou excluir depois é linha a linha, como hoje. O cartão HOJE e a trava de horas somam o total digitado, não o total vezes a quantidade de empresas.

## 2. O que NÃO muda (de propósito)

- **Banco**: nada. A tela grava na `TbTime` com as mesmas procedures que o Time usa hoje (`Time_GetByData`, `Time_Get`, `Time_Insert`, `Time_Update`, `Time_Delete`) e lê a meta na `TbHorasMes` (`HorasMes_GetList`). `EMPRESA` recebe o nome do cliente ou o texto da atividade; `EMPRESAID` recebe o TRADUTOR ou fica vazio; `TEMPO` continua decimal com duas casas (1 h 30 min = 1,50). Tudo o que lê a TbTime (consultas de gestão, Custo, relatórios) continua igual.
- **Regras**: data não pode ser futura; tempo não pode ser zero; ou cliente ou atividade, nunca os dois; "administrativo" do Outros.txt só para gestor e diretoria; cada pessoa vê só os próprios apontamentos.
- **Trava de horas**: continua no `frmMain`. Depois de gravar ou excluir, a tela chama `Opl2.menuacesso()` como o `frmTimeDet` faz hoje, então os menus liberam (ou travam) na hora. Os cartões usam a **mesma conta** do `frmMain` (`VerificarHorasMensais` / `VerificarDiasIndividuais`: 20 dias corridos até ontem, sem sábado e domingo, menos de 8 h; mês anterior contra a `TbHorasMes`; sem cobrança no mês da admissão), agora num lugar só: `ResumoHoras.Calcular` em `Apontamento.vb`.
- **Consultas de gestão** (`frmTimeConsulta`, `...Acumulado`, `...Funcionario`, `frmTimePeriodo`): ficam como estão. São outra entrega.

## 3. O que saiu

- **Retrabalho**: o campo está `Visible = False` no Designer desde 2022 e não há nenhum lançamento "Retrabalho - ..." desde 2025. Lançamentos antigos com esse prefixo abrem normalmente (o texto aparece como atividade).
- **CNPJ/CPF digitável** e cód./nome do grupo: eram só exibição (a TbTime não grava). Viraram a linha de informação abaixo do cliente escolhido.
- A necessidade de clicar "Inserir" antes de preencher, e a janela de edição que não carregava o registro (hoje é preciso clicar Inserir e Gravar para alterar). Agora abrir um lançamento já mostra o que está gravado.

## 4. Arquivos

| Arquivo | Vai para o projeto? | O que é |
|---|---|---|
| `Apontamento.vb` | sim: `Grupo_Mecano\Cadastro\` | Modelos (`Apontamento`, `ClienteResumo`, `DiaPendente`, `ResumoHoras` com a regra da trava). |
| `ModApontamento.vb` | sim: `Grupo_Mecano\App_Code\` | Acesso a dados **SQL** (procedures do Time, TbClientes, Outros.txt, `AposGravar` → `Opl2.menuacesso()`). |
| `frmApontamentos.vb` | sim: `Grupo_Mecano\Cadastro\` | Tela principal (lista, cartões, período, busca, exportar). Tem `Abrir()` e `Embutir()`. |
| `frmApontamentoDet.vb` | sim: `Grupo_Mecano\Cadastro\` | Lançamento novo / alterar, inclusive várias empresas (mesmo tempo ou rateio) e a trava de 10 h. |
| `frmBancoHoras.vb` | sim: `Grupo_Mecano\Cadastro\` | Extrato do banco de horas (só leitura), aberto pelo cartão BANCO DE HORAS. |
| `ModApontamentoDemo.vb` | **não** | Versão de demonstração do `ModApontamento` (txt, sem banco). Mesmos membros públicos. |
| `DemoMain.vb`, `build-demo.rsp`, `Compilar-Demo.bat` | não | Ponto de entrada e compilação da demo (`vbc` do .NET 4.8, sem Visual Studio). |
| `clientes.txt`, `outros.txt`, `horasmes.txt`, `apontamentos.txt` | não | Dados da demo. |
| `Neulis_Neue_Regular.otf`, `LOGO_PRINCIPAL_2.png` | já estão na entrega da tela de login | Fonte e logo da marca, lidos pelo `TemaMecano`. |

Dependências já entregues antes: `TemaMecano.vb` (tela de login) e `UiMecano.vb` + `Colaborador.vb` (Colaboradores; usa `BotaoMecano`, `FormasMecano` e `Colaborador.SemAcento`). A demo lê esses três direto de `..\TO-DO_Entrega-Cadastro-Colaboradores`, sem cópia.

`ModApontamento.vb` (SQL) foi compilado junto com as telas usando dublês das globais do Gerenciador (`Usersys`, `UserFuncao`, `UserAdmissao`, `My.Settings.Myconnection`, `Opl2.menuacesso`). O projeto real ainda precisa ser compilado no Visual Studio.

## 5. Para colocar no Gerenciador

1. Incluir no `.vbproj` os cinco arquivos marcados "sim" acima (mais `TemaMecano.vb`, `UiMecano.vb` e `Colaborador.vb`, se a entrega de Colaboradores ainda não tiver subido).
2. `frmMain`: no item **Apontamentos › Time › Anotações** trocar `New frmTimeDet(0)...Show()` por `frmApontamentos.Abrir()`; o item **Consulta** (`frmTime`) pode apontar para a mesma tela ou sair do menu. Na tela principal nova, o item "Apontamentos" da barra lateral e o atalho "Horas" já chamam `frmApontamentos` embutida (`Embutir()`), como está na demo v12.
3. Conferir na máquina do usuário que `I:\Sistemas\Nibo\Clientes\Outros.txt` abre (a tela tenta também `\\mec-srv-00\Dados\...`); sem o arquivo ela avisa e só permite clientes.
4. Compilar no Visual Studio e testar com um usuário comum: lançar, alterar, excluir, várias empresas, e confirmar que o menu libera depois de fechar um dia pendente (mesma verificação do item 7 do TO-DO de Colaboradores).
5. Quando o script `2026-09-23-Perfis-e-Trava-Horas.sql` subir (jornada por pessoa, isentos, feriados), trocar as duas constantes de `ResumoHoras` (`JORNADA`, `DIAS_VERIFICADOS`) pela leitura de `JORNADADIARIA` / `TbFeriado`, ou chamar `Acesso_VerificarHoras` e só exibir o resultado. Até lá a tela usa 8 h e 20 dias, iguais ao `frmMain` de hoje.
6. Testar em 125 % e 150 % de DPI (as capturas desta pasta foram feitas em 150 %).

## 6. Pontos para decidir

- **Retrabalho**: confirmar que pode sair de vez (hoje já está oculto).
- **Consulta do Time** (menu › Consulta): sai do menu ou vira o mesmo item da tela nova?
- **Cartão "Horas do mês" na tela Início**: a proposta da tela principal previa esse cartão à direita do Início (nota "NOVO · HORAS DO MÊS" do canvas). O componente que desenha os cartões (`PainelResumoHoras`, em `frmApontamentos.vb`) já serve para isso; falta decidir se entra no Início junto com o calendário do Nibo ou no lugar dele.
- **Lançar para outra pessoa** (gestor preenchendo o time de alguém, como o `frmTimePeriodo` faz por período): fica fora desta tela ou entra como opção só para gestor?
- **Banco de horas**: confirmar a data de início da conta (hoje 01/10/2026), se sábado e domingo entram inteiros, e se o crédito deve alimentar o "Banco de Horas (Saldos)" do RH no fechamento do mês.
- **Modo "várias empresas"**: o padrão ao abrir é "cada empresa fica com o tempo inteiro" (comportamento do Time de hoje). Trocar o padrão para "dividir" é uma linha.

# Implantação — Apontamentos de horas (pré-produção e produção)

Servidor de banco: `MEC-SRV-04\SQLEXPRESS`. Produção = `Grupo_Mecano`. Pré-produção (teste) = `Grupo_Mecano_Testes`.
Projeto: `Grupo_Mecano` (Visual Basic, .NET Framework 4.8). A cópia de testes do projeto está em `..\Grupo_Mecano_Testes`.

**Esta entrega não tem script SQL.** A tela usa as procedures que o Time já usa (`Time_GetByData`, `Time_Get`, `Time_Insert`, `Time_Update`, `Time_Delete`, `HorasMes_GetList`) e lê `TbClientes`. Nada é criado nem alterado no banco. Voltar atrás é só apontar o menu de volta para a tela antiga.

---

## Passo 0 — Pré-requisitos

- [ ] `TemaMecano.vb` (entrega da tela de login) já está no projeto.
- [ ] `UiMecano.vb` e `Colaborador.vb` (entrega de Colaboradores) já estão no projeto. Se não, incluir os dois junto com esta entrega: a tela usa `BotaoMecano`, `FormasMecano` e `Colaborador.SemAcento`.
- [ ] `Neulis_Neue_Regular.otf` e `LOGO_PRINCIPAL_2.png` chegam ao `bin` do Gerenciador (já acontece desde a tela de login; `TemaMecano` cai para Segoe UI se a fonte faltar).
- [ ] Na máquina de teste, `I:\Sistemas\Nibo\Clientes\Outros.txt` abre (a tela tenta também `\\mec-srv-00\Dados\Sistemas\Nibo\Clientes\Outros.txt`). Sem o arquivo ela avisa e só permite clientes.

## Passo 1 — Arquivos no projeto

A pasta `preprod\` tem só os sete arquivos que entram no Gerenciador, nada da demonstração. Copiar para o projeto de **testes** primeiro (`Grupo_Mecano_Testes\Grupo_Mecano\`):

| De (este repositório) | Para |
|---|---|
| `preprod\Apontamento.vb` | `Cadastro\Apontamento.vb` |
| `preprod\frmApontamentos.vb` | `Cadastro\frmApontamentos.vb` |
| `preprod\frmApontamentoDet.vb` | `Cadastro\frmApontamentoDet.vb` |
| `preprod\frmBancoHoras.vb` | `Cadastro\frmBancoHoras.vb` |
| `preprod\frmApontamentoPeriodo.vb` | `Cadastro\frmApontamentoPeriodo.vb` |
| `preprod\frmGerenciamentoTime.vb` | `Cadastro\frmGerenciamentoTime.vb` |
| `preprod\ModApontamento.vb` | `App_Code\ModApontamento.vb` |

- [ ] Incluir os sete no `Grupo_Mecano.vbproj` (Adicionar › Item existente). São classes criadas por código, **sem Designer e sem .resx**.
- [ ] **Não** incluir `ModApontamentoDemo.vb`, `DemoMain.vb` nem os `.txt`: são só da demonstração.

## Passo 2 — Menu (`Base\frmMain.vb`)

- [ ] Em **Apontamentos › Time › Anotações**, trocar `New frmTimeDet(0)…Show()` por `frmApontamentos.Abrir()`.
- [ ] Em **Apontamentos › Time › Consulta** (`frmTime`): apontar para `frmApontamentos.Abrir()` também, ou retirar do menu (decisão em aberto; em pré-produção, apontar para a tela nova).
- [ ] Em **Apontamentos › Time › Anotações Período**, trocar `AbrirFormulario(Of frmTimePeriodo)()` por abrir `New frmApontamentoPeriodo()` com `ShowDialog()`.
- [ ] Em **Apontamentos › Time › Gerenciamento**, os quatro itens passam a abrir a mesma tela, cada um na sua aba:
  - Consulta Analítica: `AbrirFormulario(Of frmTimeConsulta)()` → `frmGerenciamentoTime.Abrir()`
  - Sintética: `frmTimeConsultaAcumulado` → `frmGerenciamentoTime.Abrir(frmGerenciamentoTime.ABA_SINTETICA)`
  - Geral: `frmTimeFuncionario` → `frmGerenciamentoTime.Abrir(frmGerenciamentoTime.ABA_FUNCIONARIO)`
  - Acumulada por Usuário e Data: `frmTimeConsultausuario` → `frmGerenciamentoTime.Abrir(frmGerenciamentoTime.ABA_USUARIO_DATA)`
- [ ] Deixar `frmTimeDet`, `frmTime`, `frmTimePeriodo`, `frmTimeConsulta`, `frmTimeConsultaAcumulado`, `frmTimeFuncionario` e `frmTimeConsultausuario` no projeto. São o caminho de volta e servem para a comparação do roteiro abaixo.
- [ ] Na tela principal nova (quando subir), o item "Apontamentos" da barra lateral e o atalho "Horas" chamam `frmApontamentos` com `Embutir()`.

## Passo 3 — Compilar

- [ ] Compilar no Visual Studio com `Option Strict On` (os arquivos já vêm assim). Erros esperados só se faltar o passo 0.
- [ ] Conferir que a conexão de testes aponta para `Grupo_Mecano_Testes` (`My.Settings.Myconnection`).

## Passo 4 — Roteiro de teste em pré-produção

Fazer com um **usuário comum** (não gestor) e depois com um **gestor**. Em cada item, conferir na `TbTime` que a linha gravada tem `EMPRESA`, `EMPRESAID` (TRADUTOR ou vazio), `TEMPO` com duas casas, `USUARIOID` e `DATA` corretos.

**Básico**
- [ ] Abrir a tela: cartões carregam, lista do período "Este mês" aparece, barra de status mostra contagem e total.
- [ ] Lançar 1 h para um cliente, hoje. Aparece na lista e no cartão HOJE.
- [ ] Lançar uma atividade interna. "Administrativo" não aparece para usuário comum; aparece para gestor/diretoria.
- [ ] Editar um lançamento (duplo clique): abre com os dados gravados; mudar o tempo e salvar.
- [ ] Excluir com Del: pede confirmação; some da lista e do cartão.
- [ ] Data futura é recusada. Tempo zero é recusado.
- [ ] Exportar gera o CSV e abre no Excel com acentos corretos.

**Várias empresas**
- [ ] Marcar "Lançar para várias empresas", passar 4 empresas, 2 h, **"cada empresa fica com o tempo inteiro"**: quatro linhas de 2,00; dia soma 8 h.
- [ ] Mesmas 4 empresas, 2 h, **"dividir"**: quatro linhas de 0,50; dia soma 2 h.
- [ ] 3 empresas, 2 h, dividir: 0,67 + 0,67 + 0,66 (soma 2,00). A tela mostra "cerca de 0h40 para cada uma".
- [ ] 20 empresas, 5 min, dividir: recusado ("algumas ficariam com zero").

**Trava de 10 h**
- [ ] Num dia com 8 h, lançar mais 1 h: grava e a mensagem diz "1h00 foi para o banco de horas".
- [ ] No mesmo dia (9 h), lançar 3 h: recusado, "cabem mais 1h00".
- [ ] Lançar 1 h: grava; dia em 10 h. Lançar mais 15 min: recusado, "não dá para lançar mais nesse dia".
- [ ] Várias empresas num dia com 9 h, 4 empresas, 1 h "cada": recusado (4 h no total). Com "dividir": grava (1 h no total).
- [ ] Editar uma linha de 1 h para 2 h num dia de 10 h: recusado.

**Banco de horas**
- [ ] Cartão BANCO DE HORAS mostra o saldo dos dias acima de 8 h desde 01/10/2026.
- [ ] Clicar abre o extrato; dias, jornada e "para o banco" batem com a TbTime. Exportar funciona.
- [ ] Excluir uma linha que gerou crédito: ao reabrir a tela, o saldo cai sozinho.
- [ ] Lançar 4 h num sábado: entra inteiro no banco.

**Gestão (com um gestor; com um analista os dois botões não aparecem)**
- [ ] "Lançar período": escolher outra pessoa, duas semanas, Férias / Afastamento. A janela diz "10 dias úteis × 8h00 e 2 sábados × 4h00 = 88h00". Gravar e conferir na TbTime: 12 linhas no USUARIOID dela, EMPRESA = "Ferias / Afastamento" e "Sabado", EMPRESAID vazio, TEMPO 8,00 e 4,00.
- [ ] Repetir o mesmo período: a janela avisa que os dias já têm apontamento e não grava nada.
- [ ] "Gerenciamento", aba Analítica: mesmo período e filtros na tela antiga (Consulta Analítica) e na nova; a contagem de linhas e o total de horas têm de ser iguais nos três modos (Tudo / só clientes / só atividades internas), com "Todos" e com um colaborador escolhido, com "Todas as empresas" e com uma empresa.
- [ ] Exportar de cada aba abre no Excel com acentos corretos e a linha de total.

**Gerenciamento: as outras três abas, comparadas com as telas antigas**

Em cada comparação, usar o mesmo período nas duas telas. A tela nova mostra horas como "8h30"; a antiga, como "8,50". Os totais são os mesmos.

- [ ] **Sintética × Consulta dados Acumulados (frmTimeConsultaAcumulado)**, com "Acumular por Empresa" desmarcado nas duas: mesma quantidade de linhas e mesmo "TOTAL DE HORAS" nos três modos (Geral / Com Código / Sem Código), com "Todos" e com um usuário, com "Todas as empresas" e com uma empresa, e com um grupo de empresas escolhido.
- [ ] Sintética com "Somar todos os colaboradores por empresa" (o "Acumular por Empresa", marcado): o total é o mesmo da tela antiga. **Diferença esperada**: a tela antiga junta todas as atividades sem código numa linha só (a chave dela é o código, vazio); a nova deixa uma linha por atividade. Nos clientes as linhas são iguais.
- [ ] **Por funcionário × Time Funcionário (frmTimeFuncionario, "Consulta Geral")**: na antiga a busca é por competência (um mês); na nova, informar do dia 1 ao último dia do mesmo mês. Diretas, indiretas e totais por pessoa têm de bater, e os totais gerais de horas diretas e indiretas também. **Diferenças esperadas**: a tela antiga esconde quatro pessoas com nome fixo no código (a diretoria); a nova mostra todo mundo que tem horas. Os percentuais saem inteiros (90%) em vez de 90,00.
- [ ] Por funcionário: duplo clique numa pessoa abre a Sintética dela no mesmo período, com as horas somadas por empresa e mês (a antiga abre o `frmTimeConsultaAcumuladoFuncionario`).
- [ ] **Por usuário e data × Consulta dados no Time (frmTimeConsultausuario)**: mesma lista de pessoas e mesmas horas. **Atenção na comparação**: a tela antiga, antes de mostrar, "converte" valores com fração acima de 0,60 como se fossem minutos (8,75 vira 9,15), então para quem tem fração alta o número da grade antiga não é a soma real. Conferir contra a soma da pessoa na Analítica (ou na coluna "Horas (decimal)" do CSV). Diferenças esperadas: a antiga não lista seis nomes fixos no código (Marcelo, Junior, Teste, Cassia, Ingrid, Thiago); a nova lista todos os ativos da TbSenha menos o ADMINISTRADOR, e os desligados só se tiverem horas no período.
- [ ] Por usuário e data: o "Esperado" é 8h00 por dia útil, 4h00 no sábado e nada no domingo (a mesma conta da tela antiga); fica em vermelho quem está abaixo ou mais de 0,15 h acima, como hoje.

**Trava de horas do menu principal**
- [ ] Com um dia pendente (menos de 8 h) nos últimos 20 dias úteis, o menu principal fica travado como hoje; o cartão MENOS DE 8 H mostra o dia.
- [ ] Clicar no chip do dia, completar as 8 h, salvar: ao voltar, o menu libera (a tela chama `Opl2.menuacesso()` depois de gravar).

**Visual**
- [ ] Testar em 100 %, 125 % e 150 % de DPI: nada cortado, botões do rodapé inteiros, "+N" do cartão de dias pendentes ao lado dos chips.

## Passo 5 — Produção

Só depois do roteiro acima aprovado em pré-produção e das decisões em aberto respondidas (ver README).

- [ ] Repetir os passos 1 a 3 no projeto de produção (`Grupo_Mecano\Grupo_Mecano\`).
- [ ] Publicar o Gerenciador como de costume.
- [ ] Avisar a equipe: nova tela, teto de 10 h por dia, banco de horas na própria tela.

## Voltar atrás

Sem nada no banco, voltar é trocar o menu de volta para `New frmTimeDet(0)…Show()` e `frmTime`. Os lançamentos feitos pela tela nova são linhas comuns da `TbTime` e continuam aparecendo nas telas antigas e nos relatórios.

## Depois (não bloqueia a pré-produção)

- Quando o script `2026-09-23-Perfis-e-Trava-Horas.sql` (entrega de Colaboradores) subir: trocar as constantes `ResumoHoras.JORNADA` e `DIAS_VERIFICADOS` pela leitura de `JORNADADIARIA` e `TbFeriado`, e `BancoHoras.JornadaDoDia` passa a considerar feriados.
- Se a diretoria quiser, alimentar o "Banco de Horas (Saldos)" do RH no fechamento do mês com o crédito calculado (uma linha por competência, `EntradaSaidaBancoHorasSaldo_Insert`).

# Implantação — Apontamentos de horas (pré-produção e produção)

Servidor de banco: `MEC-SRV-04\SQLEXPRESS`. Produção = `Grupo_Mecano`. Pré-produção (teste) = `Grupo_Mecano_Testes`.
Projeto: `Grupo_Mecano` (Visual Basic, .NET Framework 4.8). A cópia de testes do projeto está em `..\Grupo_Mecano_Testes`.

**Um script SQL, só para os grupos de empresas.** A tela usa as procedures que o Time já usa (`Time_GetByData`, `Time_Get`, `Time_Insert`, `Time_Update`, `Time_Delete`, `HorasMes_GetList`) e lê `TbClientes`. A única novidade no banco é a tabela `TbTimeGrupoEmpresa` ("Meus grupos de empresas"), criada por `preprod\2026-10-20-Grupos-de-Empresas.sql`; sem ela o botão e o campo de grupos somem e o resto funciona. Voltar atrás é só apontar o menu de volta para a tela antiga (a tabela pode ficar).

---

## Passo 0 — Pré-requisitos

- [ ] `TemaMecano.vb` (entrega da tela de login) já está no projeto.
- [ ] `UiMecano.vb` e `Colaborador.vb` (entrega de Colaboradores) já estão no projeto. Se não, incluir os dois junto com esta entrega: a tela usa `BotaoMecano`, `FormasMecano` e `Colaborador.SemAcento`.
- [ ] `Neulis_Neue_Regular.otf` e `LOGO_PRINCIPAL_2.png` chegam ao `bin` do Gerenciador (já acontece desde a tela de login; `TemaMecano` cai para Segoe UI se a fonte faltar).
- [ ] Na máquina de teste, `I:\Sistemas\Nibo\Clientes\Outros.txt` abre (a tela tenta também `\\mec-srv-00\Dados\Sistemas\Nibo\Clientes\Outros.txt`). Sem o arquivo ela avisa e só permite clientes.

## Passo 0b — Script SQL (grupos de empresas)

- [ ] No SQL Server Management Studio, conectado em `MEC-SRV-04\SQLEXPRESS`, abrir `preprod\2026-10-20-Grupos-de-Empresas.sql`, escolher o banco **Grupo_Mecano_Testes** e executar. Ele imprime "TbTimeGrupoEmpresa criada." (ou "já existia").
- [ ] Em produção, o mesmo script no banco `Grupo_Mecano`, no passo 5.

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
| `preprod\frmGruposEmpresas.vb` | `Cadastro\frmGruposEmpresas.vb` |
| `preprod\ModApontamento.vb` | `App_Code\ModApontamento.vb` |

- [ ] Incluir os oito no `Grupo_Mecano.vbproj` (Adicionar › Item existente). São classes criadas por código, **sem Designer e sem .resx**.
- [ ] **Não** incluir `ModApontamentoDemo.vb`, `DemoMain.vb` nem os `.txt`: são só da demonstração.

## Passo 2 — Menu (`Base\frmMain.vb`)

Tudo deste passo é em **um arquivo só**, `Base\frmMain.vb`, dentro do bloco `#Region "Menus — Arquivo / Gerenciamento"` (por volta das linhas 830 a 905). Nada muda no `frmMain.Designer.vb`: os itens do menu continuam os mesmos, com os mesmos nomes; só troca o que cada clique abre. Os nomes dos `Sub` e a cláusula `Handles` ficam como estão, porque é o `Handles` que liga o item do menu ao código.

O menu **Apontamentos › Time** tem hoje estes itens, nesta ordem:

| Item do menu | `Sub` no frmMain.vb | Abre hoje | Passa a abrir |
|---|---|---|---|
| Anotações | `AnotaçõesToolStripMenuItem_Click` | `frmTimeDet(0)` | `frmApontamentos.Abrir()` |
| Anotações Período | `AnotaçõesPeríodoToolStripMenuItem_Click` | `frmTimePeriodo` | `frmApontamentoPeriodo` (janela modal) |
| Consulta | `ConsultaToolStripMenuItem1_Click` | `frmTime` | `frmApontamentos.Abrir()` (a lista é a mesma tela) |
| Gerenciamento › Consulta Analítica | `ConsultaToolStripMenuItem6_Click` | `frmTimeConsulta` | `frmGerenciamentoTime.Abrir()` |
| Gerenciamento › Consulta Sintética | `ConsultaAcumuladaToolStripMenuItem_Click` | `frmTimeConsultaAcumulado` | `frmGerenciamentoTime.Abrir(frmGerenciamentoTime.ABA_SINTETICA)` |
| Gerenciamento › Consulta Geral | `ToolStripMenuItem29_Click` | `frmTimeFuncionario` | `frmGerenciamentoTime.Abrir(frmGerenciamentoTime.ABA_FUNCIONARIO)` |
| Gerenciamento › Consulta Analítica Acumulada por Usuário e Data | `ConsultaAnalíticaAcumuladaPoUsuárioEDataToolStripMenuItem_Click` | `frmTimeConsultausuario` | `frmGerenciamentoTime.Abrir(frmGerenciamentoTime.ABA_USUARIO_DATA)` |
| Gerenciamento › Controle do Ponto | — | `frmEntradaSaidaConsulta` | **não mexer** (fora da entrega) |

### 2.1 — O que está lá hoje

Os sete `Sub` estão espalhados pelo bloco (entre eles há itens de outros menus; esses ficam como estão). São estes, exatamente como estão em produção:

```vb
Private Sub AnotaçõesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnotaçõesToolStripMenuItem.Click
    Dim f As New frmTimeDet(0)
    f.MaximizeBox = True : f.MinimizeBox = True : f.ShowInTaskbar = True : f.Show()
End Sub

Private Sub ConsultaToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ConsultaToolStripMenuItem1.Click
    AbrirFormulario(Of frmTime)()
End Sub

Private Sub ConsultaToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ConsultaToolStripMenuItem6.Click
    AbrirFormulario(Of frmTimeConsulta)()
End Sub

Private Sub ConsultaAcumuladaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsultaAcumuladaToolStripMenuItem.Click
    AbrirFormulario(Of frmTimeConsultaAcumulado)()
End Sub

Private Sub ConsultaAnalíticaAcumuladaPoUsuárioEDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsultaAnalíticaAcumuladaPoUsuárioEDataToolStripMenuItem.Click
    AbrirFormulario(Of frmTimeConsultausuario)()
End Sub

Private Sub ToolStripMenuItem29_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem29.Click
    AbrirFormulario(Of frmTimeFuncionario)()
End Sub

Private Sub AnotaçõesPeríodoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnotaçõesPeríodoToolStripMenuItem.Click
    AbrirFormulario(Of frmTimePeriodo)()
End Sub
```

### 2.2 — Como fica

Trocar **só o corpo** de cada um dos sete `Sub` (a linha `Private Sub … Handles …` e o `End Sub` ficam iguais):

```vb
' Apontamentos > Time > Anotações: a tela nova de apontamentos (lista, cartões e lançamento).
Private Sub AnotaçõesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnotaçõesToolStripMenuItem.Click
    frmApontamentos.Abrir()
End Sub

' Apontamentos > Time > Consulta: a lista já está na tela nova; abre a mesma tela.
Private Sub ConsultaToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ConsultaToolStripMenuItem1.Click
    frmApontamentos.Abrir()
End Sub

' Apontamentos > Time > Gerenciamento > Consulta Analítica
Private Sub ConsultaToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ConsultaToolStripMenuItem6.Click
    frmGerenciamentoTime.Abrir(frmGerenciamentoTime.ABA_ANALITICA)
End Sub

' Apontamentos > Time > Gerenciamento > Consulta Sintética
Private Sub ConsultaAcumuladaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsultaAcumuladaToolStripMenuItem.Click
    frmGerenciamentoTime.Abrir(frmGerenciamentoTime.ABA_SINTETICA)
End Sub

' Apontamentos > Time > Gerenciamento > Consulta Analítica Acumulada por Usuário e Data
Private Sub ConsultaAnalíticaAcumuladaPoUsuárioEDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsultaAnalíticaAcumuladaPoUsuárioEDataToolStripMenuItem.Click
    frmGerenciamentoTime.Abrir(frmGerenciamentoTime.ABA_USUARIO_DATA)
End Sub

' Apontamentos > Time > Gerenciamento > Consulta Geral
Private Sub ToolStripMenuItem29_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem29.Click
    frmGerenciamentoTime.Abrir(frmGerenciamentoTime.ABA_FUNCIONARIO)
End Sub

' Apontamentos > Time > Anotações Período: janela modal, em nome de outra pessoa (gestor e diretoria).
Private Sub AnotaçõesPeríodoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnotaçõesPeríodoToolStripMenuItem.Click
    Using f As New frmApontamentoPeriodo()
        f.ShowDialog(Me)
    End Using
End Sub
```

Por que cada forma de abrir:

- `frmApontamentos.Abrir()` e `frmGerenciamentoTime.Abrir(aba)` são `Shared`: abrem a tela uma vez só e, se ela já estiver aberta, trazem para a frente (na aba pedida, no caso do Gerenciamento). Não precisam do `AbrirFormulario(Of …)()`, que criaria uma janela nova a cada clique.
- `frmApontamentoPeriodo` é uma janela de confirmação, por isso `ShowDialog(Me)` dentro de `Using`, como a própria tela de apontamentos faz pelo botão "Lançar período".
- `MaximizeBox`, `MinimizeBox` e `ShowInTaskbar` já são verdadeiros por padrão num Form; as três atribuições que o item Anotações fazia não precisam ser repetidas.

### 2.3 — Conferir antes de compilar

- [ ] Os sete `Sub` continuam com o mesmo nome e o mesmo `Handles` (se o Visual Studio mostrar "evento sem manipulador" no Designer, foi mudado o nome por engano).
- [ ] Nenhuma linha do `frmMain.Designer.vb` foi alterada.
- [ ] O item **Controle do Ponto** continua abrindo `frmEntradaSaidaConsulta`.
- [ ] As telas antigas (`frmTimeDet`, `frmTime`, `frmTimePeriodo`, `frmTimeConsulta`, `frmTimeConsultaAcumulado`, `frmTimeFuncionario`, `frmTimeConsultausuario`) **continuam no projeto** e no `.vbproj`. Não apagar: são o caminho de volta e servem para a comparação do roteiro do passo 4. Elas só deixam de estar no menu.
- [ ] `ModApontamento.AposGravar()` chama `Opl2.menuacesso()` depois de gravar ou excluir; esse `Sub` é público no `frmMain.vb` (classe `Opl2`) e já existe. Nada a fazer, só saber que é daí que a trava de horas é reavaliada.

### 2.4 — Voltar atrás

Se precisar voltar para as telas antigas, é só devolver os sete corpos de 2.1. Nada no banco muda num sentido nem no outro.

### 2.5 — Depois, na tela principal nova

Quando a proposta da tela principal subir, o item "Apontamentos" da barra lateral e o atalho "Horas" chamam `frmApontamentos` com `Embutir()`; não é parte deste passo.

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

**Meus grupos de empresas**
- [ ] Botão "Meus grupos" na barra da lista abre a janela. Criar um grupo com 3 empresas e gravar: na TbTimeGrupoEmpresa aparecem 3 linhas com o USUARIOID da pessoa.
- [ ] Renomear o grupo e gravar: as linhas trocam de nome, sem duplicar. Tirar uma empresa e gravar: fica com 2 linhas.
- [ ] "Novo apontamento" › "Lançar para várias empresas": o campo Grupo lista o grupo; escolher passa as empresas para a direita. Gravar 1 h "para cada": uma linha por empresa na TbTime.
- [ ] Com empresas selecionadas, "Meus grupos…" abre já com elas num grupo novo.
- [ ] Excluir o grupo: some do campo Grupo e da tabela.
- [ ] Com outro usuário, os grupos do primeiro não aparecem.
- [ ] (Opcional) Apagar a tabela num banco de teste: o botão e o campo somem, o resto funciona.

**Trava de 10 h**
- [ ] Num dia com 8 h, lançar mais 1 h: o subtítulo fica vermelho ("acima de 8h00"), grava, e aparece a mensagem "Hoje ficou com 9h00, acima de 8h00". A linha está na TbTime.
- [ ] No mesmo dia (9 h), lançar 3 h: recusado, "cabem mais 1h00".
- [ ] Lançar 1 h: grava; dia em 10 h. Lançar mais 15 min: recusado, "não dá para lançar mais nesse dia".
- [ ] Várias empresas num dia com 9 h, 4 empresas, 1 h "cada": recusado (4 h no total). Com "dividir": grava (1 h no total).
- [ ] Editar uma linha de 1 h para 2 h num dia de 10 h: recusado.

**Banco de horas (adiado em 19/10/2026)**
- [ ] Cartão BANCO DE HORAS aparece sem número, com "em breve"; clicar nele não abre nada e o cursor não vira mão.
- [ ] Nenhuma mensagem da tela fala em "banco de horas".
- [ ] (Só quando `BancoHoras.ATIVO` for True) Cartão mostra o saldo dos dias acima de 8 h desde 01/10/2026; clicar abre o extrato e os números batem com a TbTime; excluir uma linha que gerou crédito faz o saldo cair; 4 h num sábado entram inteiras.

**Gestão (com um gestor; com um analista os dois botões não aparecem)**
- [ ] "Lançar período": escolher outra pessoa, duas semanas, Férias / Afastamento. A janela diz "10 dias úteis × 8h00 e 2 sábados × 4h00 = 88h00". Gravar e conferir na TbTime: 12 linhas no USUARIOID dela, EMPRESA = "Ferias / Afastamento" e "Sabado", EMPRESAID vazio, TEMPO 8,00 e 4,00.
- [ ] Repetir o mesmo período: a janela avisa que os dias já têm apontamento e não grava nada.
- [ ] "Novo apontamento" como gestor: aparece o campo **Colaborador** com o próprio nome. Escolher outra pessoa, lançar 1 h: a linha vai para o USUARIOID dela na TbTime, não aparece na lista de quem gravou e o rodapé diz "Em nome de …". Com um analista o campo não aparece.
- [ ] Como gestor, escolher uma pessoa que já tem 10 h num dia e tentar lançar mais: recusado com o nome dela na mensagem.
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

- [ ] Rodar `preprod\2026-10-20-Grupos-de-Empresas.sql` no banco `Grupo_Mecano`.
- [ ] Repetir os passos 1 a 3 no projeto de produção (`Grupo_Mecano\Grupo_Mecano\`).
- [ ] Publicar o Gerenciador como de costume.
- [ ] Avisar a equipe: nova tela, aviso acima de 8 h, teto de 10 h por dia.

## Voltar atrás

Voltar é trocar o menu de volta para `New frmTimeDet(0)…Show()` e `frmTime`. Os lançamentos feitos pela tela nova são linhas comuns da `TbTime` e continuam aparecendo nas telas antigas e nos relatórios.

## Depois (não bloqueia a pré-produção)

- Quando o script `2026-09-23-Perfis-e-Trava-Horas.sql` (entrega de Colaboradores) subir: trocar as constantes `ResumoHoras.JORNADA` e `DIAS_VERIFICADOS` pela leitura de `JORNADADIARIA` e `TbFeriado`, e `BancoHoras.JornadaDoDia` passa a considerar feriados.

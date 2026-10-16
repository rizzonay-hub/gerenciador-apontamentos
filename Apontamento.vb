Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq

' ============================================================================
'  Apontamento - modelos do módulo de APONTAMENTOS DE HORAS (o "Time").
'
'  Um apontamento é uma linha da TbTime: quanto tempo (TEMPO, em horas com
'  duas casas) o usuário (USUARIOID) dedicou em uma data (DATA) a um cliente
'  (EMPRESA = nome, EMPRESAID = código TRADUTOR) ou a uma atividade interna
'  (EMPRESA = texto do Outros.txt, EMPRESAID vazio).
'
'  ResumoHoras reúne, num lugar só, a regra que hoje está espalhada entre
'  frmTimeDet.SOMADODIA e frmMain.VerificarHorasMensais/VerificarDiasIndividuais:
'  horas de hoje, do mês e do mês anterior contra a meta da TbHorasMes, e os
'  dias úteis dos últimos 20 dias com menos de 8 h (os que travam o menu).
'
'  Destino no projeto: Grupo_Mecano\Cadastro\Apontamento.vb
' ============================================================================

Public Class Apontamento
    Public Property Id As Integer
    Public Property Data As Date = Date.Today
    ''' <summary>Nome do cliente ou texto da atividade interna (TbTime.EMPRESA, até 100 caracteres).</summary>
    Public Property Empresa As String = ""
    ''' <summary>Código TRADUTOR do cliente (TbTime.EMPRESAID, até 7). Vazio = atividade interna.</summary>
    Public Property EmpresaId As String = ""
    ''' <summary>Horas com duas casas (TbTime.TEMPO): 1,50 = 1 h 30 min.</summary>
    Public Property Horas As Decimal
    Public Property UsuarioId As Integer

    Public ReadOnly Property EhCliente As Boolean
        Get
            Return If(EmpresaId, "").Trim() <> ""
        End Get
    End Property

    Public ReadOnly Property TipoTexto As String
        Get
            Return If(EhCliente, "Cliente", "Interno")
        End Get
    End Property

    Public ReadOnly Property HorasTexto As String
        Get
            Return FormatarHoras(Horas)
        End Get
    End Property

    Public Function Clonar() As Apontamento
        Return New Apontamento With {.Id = Id, .Data = Data, .Empresa = Empresa, .EmpresaId = EmpresaId, .Horas = Horas, .UsuarioId = UsuarioId}
    End Function

    ' ---------------------------------------------------------------- horas
    Public Shared ReadOnly PtBr As CultureInfo = CultureInfo.GetCultureInfo("pt-BR")

    ''' <summary>Minutos inteiros correspondentes ao decimal gravado (1,50 → 90; 0,28 → 17).</summary>
    Public Shared Function Minutos(horas As Decimal) As Integer
        Return CInt(Math.Round(horas * 60D, MidpointRounding.AwayFromZero))
    End Function

    ''' <summary>Decimal com duas casas a partir de minutos (90 → 1,50; 17 → 0,28), como o Time sempre gravou.</summary>
    Public Shared Function DeMinutos(minutos As Integer) As Decimal
        Return Math.Round(CDec(minutos) / 60D, 2, MidpointRounding.AwayFromZero)
    End Function

    ''' <summary>
    ''' Divide um tempo em partes iguais, em centésimos de hora (a unidade que a TbTime guarda),
    ''' de modo que a soma das partes seja exatamente o total: 2,00 h em 3 → 0,67, 0,67, 0,66.
    ''' A sobra vai para as primeiras partes. Com 0,01 h ou menos por parte, devolve lista vazia:
    ''' o Time não grava tempo zerado.
    ''' </summary>
    Public Shared Function Ratear(horas As Decimal, partes As Integer) As List(Of Decimal)
        Dim lista As New List(Of Decimal)()
        If partes <= 0 Then Return lista
        Dim centesimos = CInt(Math.Round(horas * 100D, MidpointRounding.AwayFromZero))
        If centesimos < partes Then Return lista
        Dim base = centesimos \ partes
        Dim sobra = centesimos Mod partes
        For i = 0 To partes - 1
            lista.Add(CDec(base + If(i < sobra, 1, 0)) / 100D)
        Next
        Return lista
    End Function

    ''' <summary>"1h30", "0h17", "8h00", "12h".</summary>
    Public Shared Function FormatarHoras(horas As Decimal) As String
        Dim m = Minutos(horas)
        Dim negativo = m < 0
        m = Math.Abs(m)
        Dim h = m \ 60
        Dim r = m Mod 60
        Dim texto = h.ToString() & "h" & r.ToString("00")
        Return If(negativo, "-" & texto, texto)
    End Function

    ''' <summary>Decimal como o Time grava, com vírgula: "1,50".</summary>
    Public Shared Function FormatarDecimal(horas As Decimal) As String
        Return horas.ToString("0.00", PtBr)
    End Function

    ''' <summary>"ter 22/09/2026"</summary>
    Public Shared Function FormatarData(d As Date) As String
        Return PtBr.DateTimeFormat.GetAbbreviatedDayName(d.DayOfWeek).TrimEnd("."c) & " " & d.ToString("dd/MM/yyyy")
    End Function

    ''' <summary>"setembro/2026"</summary>
    Public Shared Function NomeCompetencia(competencia As String) As String
        Dim d As Date
        If Date.TryParseExact("01/" & competencia, "dd/MM/yyyy", PtBr, DateTimeStyles.None, d) Then
            Return d.ToString("MMMM/yyyy", PtBr)
        End If
        Return competencia
    End Function

    ''' <summary>Competência no formato da TbHorasMes: "MM/yyyy".</summary>
    Public Shared Function Competencia(d As Date) As String
        Return d.ToString("MM/yyyy")
    End Function
End Class

''' <summary>Cliente como o apontamento precisa: código, nome e o que a ficha mostra ao escolher.</summary>
Public Class ClienteResumo
    Public Property Tradutor As String = ""
    Public Property Nome As String = ""
    Public Property Cnpj As String = ""
    Public Property CodGrupo As Integer
    Public Property Grupo As String = ""
    Public Property Ativo As Boolean = True

    ''' <summary>Como aparece na lista: "0665  DIVERTI EVENTOS F&amp;B LTDA" (mesmo formato do Time antigo).</summary>
    Public ReadOnly Property Rotulo As String
        Get
            Return Tradutor.PadLeft(4, "0"c) & "  " & Nome & If(Ativo, "", "   (não cliente)")
        End Get
    End Property

    Public ReadOnly Property CnpjFormatado As String
        Get
            Dim digitos = New String(If(Cnpj, "").Where(Function(ch) Char.IsDigit(ch)).ToArray())
            If digitos.Length = 14 Then
                Return String.Format("{0}.{1}.{2}/{3}-{4}", digitos.Substring(0, 2), digitos.Substring(2, 3), digitos.Substring(5, 3), digitos.Substring(8, 4), digitos.Substring(12, 2))
            ElseIf digitos.Length = 11 Then
                Return String.Format("{0}.{1}.{2}-{3}", digitos.Substring(0, 3), digitos.Substring(3, 3), digitos.Substring(6, 3), digitos.Substring(9, 2))
            End If
            Return If(Cnpj, "").Trim()
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return Rotulo
    End Function
End Class

''' <summary>Dia útil dos últimos 20 dias com menos de 8 h apontadas (o que trava o menu no frmMain).</summary>
Public Class DiaPendente
    Public Property Data As Date
    Public Property Horas As Decimal

    Public ReadOnly Property Faltam As Decimal
        Get
            Return Math.Max(0D, ResumoHoras.JORNADA - Horas)
        End Get
    End Property
End Class

''' <summary>
''' Totais que a tela mostra e a regra da trava de horas, calculados a partir da
''' lista de apontamentos do usuário. A mesma função serve ao SQL e à demonstração.
''' </summary>
Public Class ResumoHoras
    ''' <summary>Jornada padrão. O script 2026-09-23-Perfis-e-Trava-Horas.sql traz JORNADADIARIA por pessoa e TbFeriado; quando ele subir, ler daqui.</summary>
    Public Const JORNADA As Decimal = 8D
    ''' <summary>Janela da verificação de dias (frmMain.VerificarDiasIndividuais): 20 dias corridos, até ontem, sem sábado e domingo.</summary>
    Public Const DIAS_VERIFICADOS As Integer = 20

    Public Property Hoje As Date
    Public Property HorasHoje As Decimal
    Public Property Competencia As String = ""
    Public Property MesApontado As Decimal
    Public Property MesMeta As Decimal
    Public Property CompetenciaAnterior As String = ""
    Public Property MesAnteriorApontado As Decimal
    Public Property MesAnteriorMeta As Decimal
    ''' <summary>False quando a pessoa foi admitida no mês atual (o frmMain não cobra o mês anterior nesse caso).</summary>
    Public Property MesAnteriorExigido As Boolean = True
    Public Property DiasPendentes As New List(Of DiaPendente)()
    ''' <summary>Horas por dia do mês atual (1..31), para o gráfico dia a dia.</summary>
    Public Property PorDiaDoMes As New Dictionary(Of Integer, Decimal)()

    Public ReadOnly Property MesAApontar As Decimal
        Get
            Return Math.Max(0D, MesMeta - MesApontado)
        End Get
    End Property

    Public ReadOnly Property MesAnteriorAApontar As Decimal
        Get
            Return Math.Max(0D, MesAnteriorMeta - MesAnteriorApontado)
        End Get
    End Property

    ''' <summary>Situação que, no frmMain de hoje, desabilita os menus.</summary>
    Public ReadOnly Property TravaMenu As Boolean
        Get
            Return DiasPendentes.Count > 0 OrElse (MesAnteriorExigido AndAlso MesAnteriorAApontar > 0D)
        End Get
    End Property

    Public Shared Function Calcular(itens As IEnumerable(Of Apontamento), meta As Func(Of String, Decimal), admissao As Date, hoje As Date) As ResumoHoras
        Dim r As New ResumoHoras With {.Hoje = hoje.Date}
        Dim lista = If(itens, New List(Of Apontamento)()).ToList()
        Dim mesAtual = New Date(hoje.Year, hoje.Month, 1)
        Dim mesAnterior = mesAtual.AddMonths(-1)

        r.Competencia = Apontamento.Competencia(mesAtual)
        r.CompetenciaAnterior = Apontamento.Competencia(mesAnterior)
        r.MesMeta = meta(r.Competencia)
        r.MesAnteriorMeta = meta(r.CompetenciaAnterior)
        r.MesAnteriorExigido = Not (admissao.Year = hoje.Year AndAlso admissao.Month = hoje.Month)

        For Each a In lista
            Dim d = a.Data.Date
            If d = hoje.Date Then r.HorasHoje += a.Horas
            If d.Year = mesAtual.Year AndAlso d.Month = mesAtual.Month Then
                r.MesApontado += a.Horas
                Dim acum As Decimal = 0D
                r.PorDiaDoMes.TryGetValue(d.Day, acum)
                r.PorDiaDoMes(d.Day) = acum + a.Horas
            ElseIf d.Year = mesAnterior.Year AndAlso d.Month = mesAnterior.Month Then
                r.MesAnteriorApontado += a.Horas
            End If
        Next

        ' Dias individuais: mesma varredura do frmMain (do dia seguinte a "hoje - N" até ontem).
        Dim diasAdm = CInt(Math.Min(CInt((hoje.Date - admissao.Date).TotalDays), DIAS_VERIFICADOS))
        Dim verdia = hoje.Date.AddDays(-diasAdm)
        For aa = 1 To diasAdm - 1
            verdia = verdia.AddDays(1)
            If verdia.DayOfWeek = DayOfWeek.Saturday OrElse verdia.DayOfWeek = DayOfWeek.Sunday Then Continue For
            Dim dia = verdia
            Dim total = lista.Where(Function(a) a.Data.Date = dia).Sum(Function(a) a.Horas)
            If total < JORNADA Then r.DiasPendentes.Add(New DiaPendente With {.Data = dia, .Horas = total})
        Next
        r.DiasPendentes = r.DiasPendentes.OrderBy(Function(p) p.Data).ToList()
        Return r
    End Function

    ''' <summary>Primeiro dia que o resumo precisa ler do banco: o menor entre o início do mês anterior e a janela de dias.</summary>
    Public Shared Function InicioNecessario(hoje As Date) As Date
        Dim mesAnterior = New Date(hoje.Year, hoje.Month, 1).AddMonths(-1)
        Dim janela = hoje.Date.AddDays(-DIAS_VERIFICADOS)
        Return If(janela < mesAnterior, janela, mesAnterior)
    End Function
End Class

''' <summary>Dia em que a pessoa apontou mais do que a jornada: a diferença é crédito no banco de horas.</summary>
Public Class DiaBanco
    Public Property Data As Date
    Public Property Apontado As Decimal
    ''' <summary>8 h em dia útil; 0 em sábado e domingo (todo o tempo do fim de semana é crédito).</summary>
    Public Property Jornada As Decimal

    Public ReadOnly Property Extra As Decimal
        Get
            Return Math.Max(0D, Apontado - Jornada)
        End Get
    End Property
End Class

''' <summary>Crédito de um mês do banco de horas ("MM/yyyy", horas, quantos dias).</summary>
Public Class CompetenciaBanco
    Public Property Competencia As String = ""
    Public Property Extra As Decimal
    Public Property Dias As Integer
End Class

''' <summary>
''' Banco de horas derivado dos apontamentos. Regra (16/10/2026): ninguém aponta mais
''' de LIMITE_DIA (10 h) num dia; o que passar da jornada (8 h em dia útil, 0 no fim de
''' semana) é crédito no banco. Nada é gravado a mais na TbTime nem em outra tabela:
''' a conta é refeita a partir dos apontamentos a cada abertura, então editar ou
''' excluir um lançamento corrige o saldo sozinho. O "Banco de Horas (Saldos)" do menu
''' Controle do Ponto (lançamento manual do RH por competência, procedures
''' EntradaSaidaBancoHorasSaldo_*) continua separado e não é tocado.
''' </summary>
Public Class BancoHoras
    ''' <summary>Teto de apontamento por dia. Acima disso a tela não grava.</summary>
    Public Const LIMITE_DIA As Decimal = 10D
    ''' <summary>Primeiro dia que entra na conta: início da vigência da regra (a confirmar com a diretoria).</summary>
    Public Shared ReadOnly INICIO_VIGENCIA As New Date(2026, 10, 1)

    Public Property Inicio As Date = INICIO_VIGENCIA
    Public Property Hoje As Date = Date.Today
    ''' <summary>Só os dias com crédito, do mais recente para o mais antigo.</summary>
    Public Property Dias As New List(Of DiaBanco)()

    Public ReadOnly Property Saldo As Decimal
        Get
            Return Dias.Sum(Function(d) d.Extra)
        End Get
    End Property

    ''' <summary>Crédito por mês, do mais recente para o mais antigo.</summary>
    Public Function PorCompetencia() As List(Of CompetenciaBanco)
        Return Dias.GroupBy(Function(d) New Date(d.Data.Year, d.Data.Month, 1)).
            OrderByDescending(Function(gr) gr.Key).
            Select(Function(gr) New CompetenciaBanco With {.Competencia = Apontamento.Competencia(gr.Key), .Extra = gr.Sum(Function(d) d.Extra), .Dias = gr.Count()}).
            ToList()
    End Function

    ''' <summary>Jornada padrão. Sábado conta 4 h: é o que a meta da TbHorasMes e o "Anotações por período" usam ("Sabado", 4 h). Domingo 0.</summary>
    Public Const JORNADA_SABADO As Decimal = 4D

    Public Shared Function JornadaDoDia(d As Date) As Decimal
        If d.DayOfWeek = DayOfWeek.Sunday Then Return 0D
        If d.DayOfWeek = DayOfWeek.Saturday Then Return JORNADA_SABADO
        Return ResumoHoras.JORNADA
    End Function

    ''' <summary>Quanto ainda cabe no dia, dado o que já está apontado.</summary>
    Public Shared Function CabeNoDia(jaApontado As Decimal) As Decimal
        Return Math.Max(0D, LIMITE_DIA - jaApontado)
    End Function

    ''' <summary>Quanto de um lançamento novo vira crédito no banco, dado o que já há no dia (7 h + 2 h num dia útil → 1 h).</summary>
    Public Shared Function ExtraDoLancamento(data As Date, jaApontado As Decimal, novo As Decimal) As Decimal
        Dim j = JornadaDoDia(data)
        Return Math.Max(0D, jaApontado + novo - j) - Math.Max(0D, jaApontado - j)
    End Function

    Public Shared Function Calcular(itens As IEnumerable(Of Apontamento), inicio As Date, hoje As Date) As BancoHoras
        Dim b As New BancoHoras With {.Inicio = inicio.Date, .Hoje = hoje.Date}
        Dim lista = If(itens, New List(Of Apontamento)())
        b.Dias = lista.Where(Function(a) a.Data.Date >= inicio.Date AndAlso a.Data.Date <= hoje.Date).
            GroupBy(Function(a) a.Data.Date).
            Select(Function(gr) New DiaBanco With {.Data = gr.Key, .Apontado = gr.Sum(Function(a) a.Horas), .Jornada = JornadaDoDia(gr.Key)}).
            Where(Function(d) d.Extra > 0D).
            OrderByDescending(Function(d) d.Data).
            ToList()
        Return b
    End Function
End Class

' ============================================================================
'  Gestão do Time (menu Apontamentos › Time › Anotações Período e › Gerenciamento)
' ============================================================================

''' <summary>Colaborador como as telas de gestão precisam (TbSenha): código, nome, departamento.</summary>
Public Class ColaboradorResumo
    Public Property Id As Integer
    Public Property Nome As String = ""
    Public Property Depto As String = ""
    Public Property Desligado As Boolean

    Public ReadOnly Property Rotulo As String
        Get
            Return Nome & If(Desligado, "   (desligado)", "")
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return Rotulo
    End Function
End Class

''' <summary>Um dia que o "Anotações por período" vai gravar (ou pular).</summary>
Public Class DiaPeriodo
    Public Property Data As Date
    ''' <summary>Texto gravado em TbTime.EMPRESA: o motivo, ou "Sabado".</summary>
    Public Property Texto As String = ""
    Public Property Horas As Decimal
    ''' <summary>Horas que a pessoa já tem nesse dia. Com qualquer coisa já lançada, o dia é pulado.</summary>
    Public Property JaTem As Decimal

    Public ReadOnly Property Pular As Boolean
        Get
            Return JaTem > 0D
        End Get
    End Property
End Class

''' <summary>
''' Regra do frmTimePeriodo de hoje, num lugar só: para cada dia do intervalo, domingo não
''' grava; sábado grava "Sabado" com 4 h; dia útil grava o motivo com 8 h. Tudo em nome do
''' colaborador escolhido, com EMPRESAID vazio (atividade interna), via Time_Insert.
''' Novidade: dias em que a pessoa já tem apontamento são pulados (hoje duplicava).
''' </summary>
Public Class PeriodoAfastamento
    ''' <summary>Textos exatamente como o Time grava hoje; relatórios podem filtrar por eles.</summary>
    Public Shared ReadOnly Motivos As String() = {"Ferias / Afastamento", "Licenca maternidade"}
    Public Const TEXTO_SABADO As String = "Sabado"

    Public Shared Function Montar(de As Date, ate As Date, motivo As String, jaApontado As Func(Of Date, Decimal)) As List(Of DiaPeriodo)
        Dim lista As New List(Of DiaPeriodo)()
        Dim d = de.Date
        While d <= ate.Date
            If d.DayOfWeek <> DayOfWeek.Sunday Then
                Dim item As New DiaPeriodo With {.Data = d}
                If d.DayOfWeek = DayOfWeek.Saturday Then
                    item.Texto = TEXTO_SABADO
                    item.Horas = BancoHoras.JORNADA_SABADO
                Else
                    item.Texto = motivo
                    item.Horas = ResumoHoras.JORNADA
                End If
                item.JaTem = If(jaApontado Is Nothing, 0D, jaApontado(d))
                lista.Add(item)
            End If
            d = d.AddDays(1)
        End While
        Return lista
    End Function

    ''' <summary>"15 dias úteis × 8h00 e 2 sábados × 4h00 = 128h00".</summary>
    Public Shared Function Descrever(dias As IEnumerable(Of DiaPeriodo)) As String
        Dim gravar = dias.Where(Function(x) Not x.Pular).ToList()
        If gravar.Count = 0 Then Return "Nenhum dia para gravar."
        Dim uteis = gravar.Where(Function(x) x.Data.DayOfWeek <> DayOfWeek.Saturday).Count()
        Dim sabados = gravar.Count - uteis
        Dim partes As New List(Of String)()
        If uteis > 0 Then partes.Add(uteis & If(uteis = 1, " dia útil × ", " dias úteis × ") & Apontamento.FormatarHoras(ResumoHoras.JORNADA))
        If sabados > 0 Then partes.Add(sabados & If(sabados = 1, " sábado × ", " sábados × ") & Apontamento.FormatarHoras(BancoHoras.JORNADA_SABADO))
        Return String.Join(" e ", partes) & " = " & Apontamento.FormatarHoras(gravar.Sum(Function(x) x.Horas))
    End Function
End Class

''' <summary>Modos da Consulta Analítica, como os três botões de opção da tela de hoje.</summary>
Public Enum ModoAnalitica
    ''' <summary>Só linhas com código de empresa (clientes): Time_GetByEmpresaCC / CCTE.</summary>
    ComCodigoEmpresa = 0
    ''' <summary>Só linhas sem código (atividades internas, "Outros"): Time_GetByEmpresaSC / SCTE.</summary>
    SemCodigoEmpresa = 1
    ''' <summary>Tudo: Time_GetByEmpresaGeral.</summary>
    Geral = 2
End Enum

''' <summary>Filtro da Consulta Analítica. Usuário e empresa viajam como NOME, que é o que as procedures esperam.</summary>
Public Class FiltroAnalitica
    Public Const TODOS_USUARIOS As String = "TODOS OS USUARIOS"
    Public Const TODAS_EMPRESAS As String = "TODAS AS EMPRESAS"

    Public Property De As Date = New Date(Date.Today.Year, Date.Today.Month, 1)
    Public Property Ate As Date = Date.Today
    ''' <summary>TbSenha.USUARIO ou TODOS_USUARIOS.</summary>
    Public Property Usuario As String = TODOS_USUARIOS
    Public Property Modo As ModoAnalitica = ModoAnalitica.Geral
    ''' <summary>TbClientes.NOME ou TODAS_EMPRESAS.</summary>
    Public Property Empresa As String = TODAS_EMPRESAS
    ''' <summary>TRADUTOR da empresa escolhida (modo Com código); vazio para todas.</summary>
    Public Property EmpresaId As String = ""

    Public ReadOnly Property TodosUsuarios As Boolean
        Get
            Return Usuario = TODOS_USUARIOS
        End Get
    End Property

    Public ReadOnly Property TodasEmpresas As Boolean
        Get
            Return Empresa = TODAS_EMPRESAS
        End Get
    End Property
End Class

''' <summary>Uma linha da Consulta Analítica (Time_GetByEmpresa*): o apontamento mais quem fez e de que departamento.</summary>
Public Class LinhaAnalitica
    Public Property Id As Integer
    Public Property Data As Date
    Public Property Empresa As String = ""
    Public Property EmpresaId As String = ""
    Public Property Horas As Decimal
    Public Property UsuarioId As Integer
    Public Property Usuario As String = ""
    Public Property Depto As String = ""
    ''' <summary>custo_por_hora da procedure. A tela de hoje não o exibe (o "TOTAL DO CUSTO" dela é sempre 0); fica aqui para o futuro.</summary>
    Public Property CustoHora As Decimal
    Public Property Grupo As String = ""

    Public ReadOnly Property EhCliente As Boolean
        Get
            Return If(EmpresaId, "").Trim() <> ""
        End Get
    End Property

    Public ReadOnly Property HorasTexto As String
        Get
            Return Apontamento.FormatarHoras(Horas)
        End Get
    End Property
End Class

''' <summary>Período da lista principal.</summary>
Public Enum PeriodoLista
    Hoje = 0
    EstaSemana = 1
    EsteMes = 2
    MesAnterior = 3
    Ultimos30Dias = 4
    Personalizado = 5
End Enum

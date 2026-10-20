Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Windows.Forms

' Ponto de entrada da DEMONSTRAÇÃO (GM-Apontamentos-Demo.exe).
' Na implantação real este arquivo NÃO é usado: a tela é aberta de dentro do
' Gerenciador (frmApontamentos.Abrir() no item Apontamentos > Time).
'
' Sem argumentos abre a tela de apontamentos. Argumentos:
'   /captura <pasta>    gera PNGs das telas sem interação, para conferir layout

''' <summary>Globais da sessão que o Gerenciador fornece (App_Code\Main.vb). Na demo você é uma analista do Fiscal.</summary>
Public Module GlobaisDemo
    Public Usersys As Integer = 999
    ' gestor, para a demo mostrar também "Lançar período" e "Gerenciamento" (um analista não vê esses botões)
    Public UserName As String = "GESTOR DE DEMONSTRAÇÃO"
    Public UserFuncao As String = "GESTOR FISCAL"
    Public UserDepto As String = "FISCAL"
    Public UserAdmissao As Date = New Date(2020, 1, 6)
End Module

Public Module DemoMain

    Private Declare Function SetProcessDPIAware Lib "user32" () As Boolean

    Public Sub Main(args As String())
        Try
            SetProcessDPIAware()
        Catch
        End Try
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)

        Dim opcao = If(args IsNot Nothing AndAlso args.Length >= 1, args(0).ToLowerInvariant(), "")
        Select Case opcao
            Case "/captura"
                If args.Length >= 2 Then Capturar(args(1))
            Case Else
                Application.Run(New frmApontamentos())
        End Select
    End Sub

    Private Sub Processar(vezes As Integer)
        For i = 1 To vezes
            Application.DoEvents()
            Threading.Thread.Sleep(60)
        Next
    End Sub

    Private Sub SalvarCaptura(f As Form, caminho As String)
        Using bmp As New Bitmap(f.Width, f.Height)
            f.DrawToBitmap(bmp, New Rectangle(0, 0, f.Width, f.Height))
            bmp.Save(caminho, ImageFormat.Png)
        End Using
    End Sub

    ''' <summary>Captura pela tela (pixels reais), para telas com controle sobreposto: o DrawToBitmap desenha os filhos fora da ordem.</summary>
    Private Sub SalvarCapturaTela(f As Form, caminho As String)
        f.TopMost = True
        f.Refresh()
        Processar(4)
        Using bmp As New Bitmap(f.Width, f.Height)
            Using g = Graphics.FromImage(bmp)
                g.CopyFromScreen(f.Location, Point.Empty, f.Size)
            End Using
            bmp.Save(caminho, ImageFormat.Png)
        End Using
        f.TopMost = False
    End Sub

    Private Sub Capturar(pasta As String)
        Directory.CreateDirectory(pasta)

        Dim lista As New frmApontamentos()
        lista.StartPosition = FormStartPosition.Manual
        lista.Location = New Point(0, 0)
        lista.Show()
        Processar(12)
        SalvarCaptura(lista, Path.Combine(pasta, "01-apontamentos.png"))

        Dim resumo = ModApontamento.CalcularResumo()

        Dim novo As New frmApontamentoDet(0, Nothing, resumo)
        novo.StartPosition = FormStartPosition.Manual
        novo.Location = New Point(0, 0)
        novo.Show()
        Processar(8)
        SalvarCaptura(novo, Path.Combine(pasta, "02-novo-apontamento.png"))
        novo.DigitarCliente("aurora")
        Processar(6)
        SalvarCapturaTela(novo, Path.Combine(pasta, "02b-busca-cliente.png"))
        novo.DigitarCliente("")
        novo.DefinirVarias(True)
        Processar(6)
        novo.EscolherGrupo("Minha carteira")
        Processar(4)
        SalvarCaptura(novo, Path.Combine(pasta, "03-varias-empresas.png"))
        novo.DemonstrarRateio(4, 120)
        Processar(6)
        SalvarCaptura(novo, Path.Combine(pasta, "03b-dividir-tempo.png"))
        novo.Close()

        Dim primeiro = ModApontamento.Listar(Date.Today.AddDays(-10), Date.Today)
        If primeiro.Count > 0 Then
            Dim editar As New frmApontamentoDet(primeiro(primeiro.Count - 1).Id, Nothing, resumo)
            editar.StartPosition = FormStartPosition.Manual
            editar.Location = New Point(0, 0)
            editar.Show()
            Processar(8)
            SalvarCaptura(editar, Path.Combine(pasta, "04-editar-apontamento.png"))
            editar.Close()
        End If

        Dim pendente = resumo.DiasPendentes.Count > 0
        If pendente Then
            Dim dia As New frmApontamentoDet(0, resumo.DiasPendentes(0).Data, resumo)
            dia.StartPosition = FormStartPosition.Manual
            dia.Location = New Point(0, 0)
            dia.Show()
            Processar(8)
            SalvarCaptura(dia, Path.Combine(pasta, "05-lancar-dia-pendente.png"))
            dia.Close()
        End If

        Dim grupos As New frmGruposEmpresas()
        grupos.StartPosition = FormStartPosition.Manual
        grupos.Location = New Point(0, 0)
        grupos.Show()
        Processar(8)
        SalvarCaptura(grupos, Path.Combine(pasta, "03c-meus-grupos.png"))
        grupos.Close()

        Dim banco As New frmBancoHoras(ModApontamento.CalcularBanco())
        banco.StartPosition = FormStartPosition.Manual
        banco.Location = New Point(0, 0)
        banco.Show()
        Processar(8)
        SalvarCaptura(banco, Path.Combine(pasta, "06-banco-de-horas.png"))
        banco.Close()

        Dim periodo As New frmApontamentoPeriodo()
        periodo.StartPosition = FormStartPosition.Manual
        periodo.Location = New Point(0, 0)
        periodo.Show()
        Processar(8)
        periodo.DemonstrarPeriodo()
        Processar(6)
        SalvarCaptura(periodo, Path.Combine(pasta, "07-anotacoes-periodo.png"))
        periodo.Close()

        Dim ger As New frmGerenciamentoTime()
        ger.StartPosition = FormStartPosition.Manual
        ger.Location = New Point(0, 0)
        ger.Show()
        Processar(8)
        ger.ConsultarAgora()
        Processar(10)
        SalvarCaptura(ger, Path.Combine(pasta, "08-gerenciamento-analitica.png"))
        ger.MostrarAba(frmGerenciamentoTime.ABA_SINTETICA)
        ger.ConsultarAgora()
        Processar(10)
        SalvarCaptura(ger, Path.Combine(pasta, "09-gerenciamento-sintetica.png"))
        ger.MostrarAba(frmGerenciamentoTime.ABA_FUNCIONARIO)
        ger.ConsultarAgora()
        Processar(10)
        SalvarCaptura(ger, Path.Combine(pasta, "10-gerenciamento-por-funcionario.png"))
        ger.MostrarAba(frmGerenciamentoTime.ABA_USUARIO_DATA)
        ger.ConsultarAgora()
        Processar(10)
        SalvarCaptura(ger, Path.Combine(pasta, "11-gerenciamento-por-usuario-data.png"))
        ger.Close()

        lista.Close()
    End Sub

End Module

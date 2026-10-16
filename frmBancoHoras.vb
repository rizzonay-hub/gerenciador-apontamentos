Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms

' ============================================================================
'  frmBancoHoras - extrato do banco de horas da pessoa logada.
'
'  Só leitura. A conta vem de BancoHoras.Calcular (Apontamento.vb): em cada dia,
'  o que passou da jornada (8 h em dia útil; sábado e domingo inteiros) é crédito.
'  Nada é gravado: editar ou excluir um apontamento corrige o saldo sozinho.
'  Aberta pelo cartão "BANCO DE HORAS" da frmApontamentos.
'
'  Não confundir com o "Banco de Horas (Saldos)" do menu Controle do Ponto, que é
'  o lançamento manual do RH por competência e continua como está.
'
'  Destino no projeto: Grupo_Mecano\Cadastro\frmBancoHoras.vb
' ============================================================================
Public Class frmBancoHoras
    Inherits Form

    Private ReadOnly _banco As BancoHoras
    Private _escala As Single = 1.0F

    Private pnlTopo As Panel
    Private pnlRodape As Panel
    Private grade As DataGridView
    Private lblVazio As Label
    Private lblTotal As Label
    Private btnExportar As BotaoMecano
    Private btnFechar As BotaoMecano

    Private ReadOnly _verde As Color = Color.FromArgb(22, 101, 52)
    Private ReadOnly _verdeFundo As Color = Color.FromArgb(222, 244, 231)

    Private Const LARGURA As Integer = 780
    Private Const ALT_TOPO As Integer = 132
    Private Const ALT_RODAPE As Integer = 64

    Private Function Esc(v As Integer) As Integer
        Return CInt(Math.Round(v * _escala))
    End Function

    Public Sub New(banco As BancoHoras)
        _banco = If(banco, New BancoHoras())
        Using g = Me.CreateGraphics()
            _escala = g.DpiX / 96.0F
        End Using

        Me.AutoScaleMode = AutoScaleMode.None
        Me.Text = "Banco de horas"
        Me.BackColor = TemaMecano.Branco
        Me.Font = TemaMecano.FonteCorpo(9.75F)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.ClientSize = New Size(Esc(LARGURA), Esc(560))
        Me.KeyPreview = True
        Me.DoubleBuffered = True

        MontarTopo()
        MontarRodape()
        MontarGrade()

        Me.Controls.Add(grade)
        Me.Controls.Add(lblVazio)
        Me.Controls.Add(pnlTopo)
        Me.Controls.Add(pnlRodape)
        grade.BringToFront()
        lblVazio.BringToFront()

        Preencher()
        ' só leitura: nenhuma linha fica "escolhida" ao abrir
        AddHandler grade.DataBindingComplete, Sub() grade.ClearSelection()
        AddHandler Me.Shown, Sub()
                                 grade.ClearSelection()
                                 btnFechar.Focus()
                             End Sub
        AddHandler Me.KeyDown, Sub(s, e)
                                   If e.KeyCode = Keys.Escape Then Me.Close()
                               End Sub
    End Sub

    ' ---------------------------------------------------------------- topo
    Private Sub MontarTopo()
        pnlTopo = New Panel With {.Dock = DockStyle.Top, .Height = Esc(ALT_TOPO), .BackColor = TemaMecano.Branco}
        AddHandler pnlTopo.Paint, AddressOf DesenharTopo
        AddHandler pnlTopo.Resize, Sub() pnlTopo.Invalidate()
    End Sub

    Private Sub DesenharTopo(sender As Object, e As PaintEventArgs)
        Dim g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
        Using p As New Pen(TemaMecano.BordaCampo)
            g.DrawLine(p, 0, pnlTopo.Height - 1, pnlTopo.Width, pnlTopo.Height - 1)
        End Using
        Dim x = Esc(24)
        Dim w = pnlTopo.Width - Esc(48)
        Dim flags = TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPrefix

        Using f = TemaMecano.FonteTitulo(15.0F)
            TextRenderer.DrawText(g, "Banco de horas", f, New Rectangle(x, Esc(14), w, Esc(30)), TemaMecano.TextoCampo, flags)
        End Using

        Dim saldo = _banco.Saldo
        Using f = TemaMecano.FonteCorpo(11.0F, FontStyle.Bold)
            TextRenderer.DrawText(g, "Saldo: " & If(saldo > 0D, "+", "") & Apontamento.FormatarHoras(saldo), f, New Rectangle(x, Esc(46), w, Esc(24)),
                                  If(saldo > 0D, _verde, TemaMecano.TextoCampo), flags)
        End Using

        Using f = TemaMecano.FonteCorpo(8.75F)
            TextRenderer.DrawText(g, "desde " & _banco.Inicio.ToString("dd/MM/yyyy"), f, New Rectangle(x, Esc(72), w, Esc(18)), TemaMecano.CinzaEscuro, flags)
        End Using

        ' um chip por mês com crédito
        Dim meses = _banco.PorCompetencia()
        If meses.Count > 0 Then
            Using fChip = TemaMecano.FonteCorpo(8.25F, FontStyle.Bold)
                Dim cx = x
                Dim cy = Esc(96)
                For Each m In meses
                    Dim texto = Apontamento.NomeCompetencia(m.Competencia) & "  +" & Apontamento.FormatarHoras(m.Extra)
                    Dim largura = TextRenderer.MeasureText(g, texto, fChip).Width + Esc(11)
                    If cx + largura > pnlTopo.Width - Esc(24) Then Exit For
                    Dim usado = FormasMecano.DesenharChip(g, texto, cx, cy, Esc(22), fChip, _verdeFundo, _verde)
                    cx += usado + Esc(6)
                Next
            End Using
        End If
    End Sub

    ' ---------------------------------------------------------------- rodapé
    Private Sub MontarRodape()
        pnlRodape = New Panel With {.Dock = DockStyle.Bottom, .Height = Esc(ALT_RODAPE), .BackColor = TemaMecano.Branco}
        AddHandler pnlRodape.Paint, Sub(s, e)
                                        Using p As New Pen(TemaMecano.BordaCampo)
                                            e.Graphics.DrawLine(p, 0, 0, pnlRodape.Width, 0)
                                        End Using
                                    End Sub

        btnFechar = BotaoMecano.Primario("Fechar", Esc(120), Esc(36))
        btnFechar.Top = Esc(14)
        btnFechar.Left = Esc(LARGURA) - Esc(24) - btnFechar.Width
        AddHandler btnFechar.Click, Sub() Me.Close()

        btnExportar = BotaoMecano.Secundario("Exportar", Esc(110), Esc(36))
        btnExportar.Top = Esc(14)
        btnExportar.Left = btnFechar.Left - Esc(10) - btnExportar.Width
        AddHandler btnExportar.Click, Sub() Exportar()

        lblTotal = New Label With {
            .AutoSize = False, .Left = Esc(24), .Top = Esc(14), .Width = btnExportar.Left - Esc(24) - Esc(12), .Height = Esc(36), .AutoEllipsis = True,
            .TextAlign = ContentAlignment.MiddleLeft, .ForeColor = TemaMecano.CinzaEscuro, .Font = TemaMecano.FonteCorpo(9.0F)}

        pnlRodape.Controls.AddRange(New Control() {btnFechar, btnExportar, lblTotal})
    End Sub

    ' ---------------------------------------------------------------- grade
    Private Sub MontarGrade()
        grade = New DataGridView With {.Dock = DockStyle.Fill}
        With grade
            .BackgroundColor = TemaMecano.Branco
            .BorderStyle = BorderStyle.None
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            .GridColor = Color.FromArgb(232, 234, 237)
            .RowHeadersVisible = False
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .ReadOnly = True
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .MultiSelect = False
            .AutoGenerateColumns = True
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            .EnableHeadersVisualStyles = False
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = Esc(36)
            .ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(38, 41, 45)
            .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            .ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(38, 41, 45)
            .ColumnHeadersDefaultCellStyle.Font = TemaMecano.FonteCorpo(9.0F, FontStyle.Bold)
            .ColumnHeadersDefaultCellStyle.Padding = New Padding(Esc(8), 0, 0, 0)
            .DefaultCellStyle.Font = TemaMecano.FonteCorpo(9.5F)
            .DefaultCellStyle.ForeColor = TemaMecano.TextoCampo
            .DefaultCellStyle.BackColor = TemaMecano.Branco
            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(222, 244, 231)
            .DefaultCellStyle.SelectionForeColor = TemaMecano.TextoCampo
            .DefaultCellStyle.Padding = New Padding(Esc(8), 0, Esc(4), 0)
            .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 251, 252)
            .RowTemplate.Height = Esc(32)
            .StandardTab = True
            .TabStop = False
        End With

        lblVazio = New Label With {
            .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleCenter, .Visible = False,
            .ForeColor = TemaMecano.CinzaClaro, .Font = TemaMecano.FonteCorpo(10.0F), .BackColor = TemaMecano.Branco,
            .Text = "Nenhum dia acima de " & Apontamento.FormatarHoras(ResumoHoras.JORNADA) & " desde " & _banco.Inicio.ToString("dd/MM/yyyy") & "."}
    End Sub

    Private Sub Preencher()
        Dim dt As New DataTable()
        dt.Columns.Add("Data", GetType(String))
        dt.Columns.Add("Dia", GetType(String))
        dt.Columns.Add("Apontado", GetType(String))
        dt.Columns.Add("Jornada", GetType(String))
        dt.Columns.Add("Para o banco", GetType(String))
        For Each d In _banco.Dias
            dt.Rows.Add(d.Data.ToString("dd/MM/yyyy"), NomeDia(d.Data), Apontamento.FormatarHoras(d.Apontado),
                        If(d.Jornada > 0D, Apontamento.FormatarHoras(d.Jornada), "fim de semana"), "+" & Apontamento.FormatarHoras(d.Extra))
        Next
        grade.DataSource = dt
        If grade.Columns.Count > 0 Then
            grade.Columns("Data").FillWeight = 18
            grade.Columns("Dia").FillWeight = 10
            grade.Columns("Dia").DefaultCellStyle.ForeColor = TemaMecano.CinzaClaro
            grade.Columns("Apontado").FillWeight = 18
            grade.Columns("Jornada").FillWeight = 22
            grade.Columns("Para o banco").FillWeight = 20
            For Each nome In New String() {"Apontado", "Jornada", "Para o banco"}
                grade.Columns(nome).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                grade.Columns(nome).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
            Next
            grade.Columns("Para o banco").DefaultCellStyle.Font = TemaMecano.FonteCorpo(9.5F, FontStyle.Bold)
            grade.Columns("Para o banco").DefaultCellStyle.ForeColor = _verde
        End If
        grade.ClearSelection()
        lblVazio.Visible = (_banco.Dias.Count = 0)
        btnExportar.Enabled = (_banco.Dias.Count > 0)

        Dim n = _banco.Dias.Count
        lblTotal.Text = If(n = 0, "", n & If(n = 1, " dia", " dias") & "  ·  total " & If(_banco.Saldo > 0D, "+", "") & Apontamento.FormatarHoras(_banco.Saldo))
    End Sub

    Private Shared Function NomeDia(d As Date) As String
        Return Apontamento.PtBr.DateTimeFormat.GetAbbreviatedDayName(d.DayOfWeek).TrimEnd("."c)
    End Function

    ' ---------------------------------------------------------------- exportar
    ''' <summary>CSV com ponto e vírgula, para abrir direto no Excel.</summary>
    Private Sub Exportar()
        If _banco.Dias.Count = 0 Then Return
        Using dlg As New SaveFileDialog With {.Filter = "CSV (*.csv)|*.csv", .FileName = "banco-de-horas-" & Date.Today.ToString("yyyy-MM-dd") & ".csv"}
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim sb As New StringBuilder()
            sb.AppendLine("Data;Dia;Apontado (h);Jornada (h);Para o banco (h)")
            For Each d In _banco.Dias.OrderBy(Function(x) x.Data)
                sb.AppendLine(d.Data.ToString("dd/MM/yyyy") & ";" & NomeDia(d.Data) & ";" & Apontamento.FormatarDecimal(d.Apontado) & ";" &
                              Apontamento.FormatarDecimal(d.Jornada) & ";" & Apontamento.FormatarDecimal(d.Extra))
            Next
            sb.AppendLine(";;;Total;" & Apontamento.FormatarDecimal(_banco.Saldo))
            Try
                File.WriteAllText(dlg.FileName, sb.ToString(), New UTF8Encoding(True))
                Process.Start(dlg.FileName)
            Catch ex As Exception
                MessageBox.Show(Me, "Não foi possível gravar o arquivo." & vbCrLf & ex.Message, "Banco de horas", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

End Class

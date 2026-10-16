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
'  frmGerenciamentoTime - "Gerenciamento" do menu Apontamentos › Time, numa tela
'  só, com uma aba por consulta de gestão:
'
'    Analítica          (hoje frmTimeConsulta)           ← pronta
'    Sintética          (hoje frmTimeConsultaAcumulado)  ← em breve
'    Por funcionário    (hoje frmTimeFuncionario, "Consulta Geral") ← em breve
'    Por usuário e data (hoje frmTimeConsultausuario)    ← em breve
'
'  A aba Analítica usa exatamente as procedures e a mesma escolha entre elas que
'  o frmTimeConsulta faz hoje (ModApontamento.ConsultaAnalitica), então os
'  números batem com a tela antiga. Só leitura; exporta CSV.
'  Só gestor e diretoria (ModApontamento.PodeGerenciar).
'
'  Destino no projeto: Grupo_Mecano\Cadastro\frmGerenciamentoTime.vb
' ============================================================================
Public Class frmGerenciamentoTime
    Inherits Form

    Private Shared _instancia As frmGerenciamentoTime = Nothing

    ''' <summary>Abre a tela (ou traz para a frente a que já está aberta).</summary>
    Public Shared Sub Abrir()
        If _instancia Is Nothing OrElse _instancia.IsDisposed Then
            _instancia = New frmGerenciamentoTime()
            _instancia.Show()
        Else
            If _instancia.WindowState = FormWindowState.Minimized Then _instancia.WindowState = FormWindowState.Normal
            _instancia.BringToFront()
            _instancia.Activate()
        End If
    End Sub

    Private _escala As Single = 1.0F

    Private Function Esc(v As Integer) As Integer
        Return CInt(Math.Round(v * _escala))
    End Function

    ' ---------- dados ----------
    Private _linhas As New List(Of LinhaAnalitica)()
    Private _clientes As New List(Of ClienteResumo)()
    Private _colaboradores As New List(Of ColaboradorResumo)()
    Private _filtroAtual As FiltroAnalitica = Nothing

    ' ---------- controles ----------
    Private pnlHeader As Panel
    Private pnlAbas As Panel
    Private pnlFiltros As Panel
    Private dtpDe As DateTimePicker
    Private dtpAte As DateTimePicker
    Private cboColaborador As ComboBox
    Private cboEmpresa As ComboBox
    Private rbCom As RadioButton
    Private rbSem As RadioButton
    Private rbGeral As RadioButton
    Private btnConsultar As BotaoMecano
    Private btnExportar As BotaoMecano
    Private grade As DataGridView
    Private lblVazio As Label
    Private barraStatus As StatusStrip
    Private lblStatus As ToolStripStatusLabel
    Private lblDica As ToolStripStatusLabel

    Private ReadOnly _abas As String() = {"Analítica", "Sintética", "Por funcionário", "Por usuário e data"}
    Private ReadOnly _abasRects As New List(Of Rectangle)()
    Private _abaAtiva As Integer = 0

    Private Const TODOS As String = "Todos os colaboradores"
    Private Const TODAS As String = "Todas as empresas"

    Public Sub New()
        Using g = Me.CreateGraphics()
            _escala = g.DpiX / 96.0F
        End Using

        Me.AutoScaleMode = AutoScaleMode.None
        Me.Text = "Gerenciador Mêcano — Gerenciamento do Time"
        Me.BackColor = TemaMecano.FundoTela
        Me.Font = TemaMecano.FonteCorpo(9.75F)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.ClientSize = New Size(Esc(1240), Esc(720))
        Me.MinimumSize = New Size(Esc(980), Esc(560))
        Me.KeyPreview = True
        Me.DoubleBuffered = True

        MontarGrade()
        MontarFiltros()
        MontarAbas()
        MontarHeader()
        MontarBarraStatus()

        Me.Controls.Add(grade)
        Me.Controls.Add(lblVazio)
        Me.Controls.Add(pnlFiltros)
        Me.Controls.Add(pnlAbas)
        Me.Controls.Add(pnlHeader)
        Me.Controls.Add(barraStatus)
        grade.BringToFront()
        lblVazio.BringToFront()

        AddHandler Me.Shown, Sub() CarregarListas()
        AddHandler Me.KeyDown, Sub(s, e)
                                   If e.KeyCode = Keys.Escape Then
                                       Me.Close()
                                   ElseIf e.KeyCode = Keys.F5 Then
                                       e.Handled = True
                                       Consultar()
                                   End If
                               End Sub
    End Sub

    ' =================================================================
    '  CABEÇALHO (faixa escura, padrão das telas novas)
    ' =================================================================
    Private Sub MontarHeader()
        pnlHeader = New Panel With {.Dock = DockStyle.Top, .Height = Esc(64), .BackColor = Color.FromArgb(18, 18, 18)}
        AddHandler pnlHeader.Paint, Sub(s, e)
                                        Dim g = e.Graphics
                                        g.SmoothingMode = SmoothingMode.AntiAlias
                                        g.InterpolationMode = InterpolationMode.HighQualityBicubic
                                        g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
                                        Dim x = Esc(24)
                                        Dim logo = TemaMecano.LogoPrincipal
                                        If logo IsNot Nothing Then
                                            Dim h = Esc(36)
                                            Dim w = CInt(logo.Width * (h / CSng(logo.Height)))
                                            g.DrawImage(logo, New Rectangle(x, (pnlHeader.Height - h) \ 2, w, h))
                                            x += w + Esc(20)
                                        End If
                                        Using f = TemaMecano.FonteTitulo(13.0F)
                                            TextRenderer.DrawText(g, "Gerenciamento do Time", f, New Point(x, Esc(10)), Color.White, TextFormatFlags.NoPadding)
                                        End Using
                                        Using f = TemaMecano.FonteCorpo(9.0F)
                                            TextRenderer.DrawText(g, "Consultas de gestão sobre os apontamentos de toda a equipe  ·  só leitura", f, New Point(x, Esc(36)), Color.FromArgb(170, 176, 180), TextFormatFlags.NoPadding)
                                        End Using
                                    End Sub
        AddHandler pnlHeader.Resize, Sub() pnlHeader.Invalidate()
    End Sub

    ' =================================================================
    '  ABAS (uma por consulta; só a Analítica está pronta)
    ' =================================================================
    Private Sub MontarAbas()
        pnlAbas = New Panel With {.Dock = DockStyle.Top, .Height = Esc(44), .BackColor = TemaMecano.Branco}
        AddHandler pnlAbas.Paint, AddressOf DesenharAbas
        AddHandler pnlAbas.Resize, Sub() pnlAbas.Invalidate()
        AddHandler pnlAbas.MouseMove, Sub(s, e) pnlAbas.Cursor = If(_abasRects.Any(Function(r) r.Contains(e.Location)), Cursors.Hand, Cursors.Default)
        AddHandler pnlAbas.MouseClick, Sub(s, e)
                                           For i = 0 To _abasRects.Count - 1
                                               If _abasRects(i).Contains(e.Location) Then
                                                   EscolherAba(i)
                                                   Return
                                               End If
                                           Next
                                       End Sub
    End Sub

    Private Sub DesenharAbas(sender As Object, e As PaintEventArgs)
        Dim g = e.Graphics
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
        Using p As New Pen(TemaMecano.BordaCampo)
            g.DrawLine(p, 0, pnlAbas.Height - 1, pnlAbas.Width, pnlAbas.Height - 1)
        End Using
        _abasRects.Clear()
        Dim x = Esc(24)
        For i = 0 To _abas.Length - 1
            Dim ativa = (i = _abaAtiva)
            Dim pronta = (i = 0)
            Dim texto = _abas(i) & If(pronta, "", "  (em breve)")
            Using f = TemaMecano.FonteCorpo(9.75F, If(ativa, FontStyle.Bold, FontStyle.Regular))
                Dim w = TextRenderer.MeasureText(g, texto, f).Width + Esc(24)
                Dim r As New Rectangle(x, 0, w, pnlAbas.Height - 1)
                _abasRects.Add(r)
                Dim cor = If(ativa, TemaMecano.VermelhoPrimario, If(pronta, TemaMecano.CinzaEscuro, TemaMecano.CinzaClaro))
                TextRenderer.DrawText(g, texto, f, r, cor, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPrefix)
                If ativa Then
                    Using b As New SolidBrush(TemaMecano.VermelhoPrimario)
                        g.FillRectangle(b, New Rectangle(x + Esc(8), pnlAbas.Height - Esc(4), w - Esc(16), Esc(3)))
                    End Using
                End If
                x += w + Esc(4)
            End Using
        Next
    End Sub

    Private Sub EscolherAba(i As Integer)
        If i = 0 Then
            _abaAtiva = 0
            pnlAbas.Invalidate()
            Return
        End If
        MostrarMensagem("""" & _abas(i) & """ ainda não está nesta tela: por enquanto use o menu Apontamentos › Time › Gerenciamento.")
    End Sub

    ' =================================================================
    '  FILTROS DA ANALÍTICA
    ' =================================================================
    Private Sub MontarFiltros()
        pnlFiltros = New Panel With {.Dock = DockStyle.Top, .Height = Esc(118), .BackColor = TemaMecano.FundoTela}
        Dim y1 = Esc(14), y2 = Esc(66)
        Dim x = Esc(24)

        pnlFiltros.Controls.Add(Rotulo("Período", x, y1))
        dtpDe = New DateTimePicker With {.Left = x, .Top = y1 + Esc(20), .Width = Esc(118), .Format = DateTimePickerFormat.Short, .Font = TemaMecano.FonteCorpo(9.75F),
                                         .Value = New Date(Date.Today.Year, Date.Today.Month, 1)}
        Dim lblA = New Label With {.Text = "a", .AutoSize = True, .Left = dtpDe.Right + Esc(6), .Top = y1 + Esc(24), .ForeColor = TemaMecano.CinzaEscuro}
        dtpAte = New DateTimePicker With {.Left = lblA.Left + Esc(16), .Top = y1 + Esc(20), .Width = Esc(118), .Format = DateTimePickerFormat.Short, .Font = TemaMecano.FonteCorpo(9.75F), .Value = Date.Today}
        x = dtpAte.Right + Esc(20)

        pnlFiltros.Controls.Add(Rotulo("Colaborador", x, y1))
        cboColaborador = New ComboBox With {.Left = x, .Top = y1 + Esc(20), .Width = Esc(280), .DropDownStyle = ComboBoxStyle.DropDownList, .Font = TemaMecano.FonteCorpo(9.75F), .FlatStyle = FlatStyle.System, .DropDownHeight = Esc(320)}
        x = cboColaborador.Right + Esc(20)

        pnlFiltros.Controls.Add(Rotulo("Empresa", x, y1))
        cboEmpresa = New ComboBox With {.Left = x, .Top = y1 + Esc(20), .Width = Esc(380), .DropDownStyle = ComboBoxStyle.DropDownList, .Font = TemaMecano.FonteCorpo(9.75F), .FlatStyle = FlatStyle.System, .DropDownHeight = Esc(360)}

        ' modo (os três botões de opção da tela de hoje), em palavras simples
        pnlFiltros.Controls.Add(Rotulo("Mostrar", Esc(24), y2))
        rbGeral = New RadioButton With {.Text = "Tudo", .AutoSize = True, .Left = Esc(24), .Top = y2 + Esc(20), .Checked = True, .Font = TemaMecano.FonteCorpo(9.5F)}
        rbCom = New RadioButton With {.Text = "Só clientes (com código de empresa)", .AutoSize = True, .Left = Esc(90), .Top = y2 + Esc(20), .Font = TemaMecano.FonteCorpo(9.5F)}
        rbSem = New RadioButton With {.Text = "Só atividades internas (sem código)", .AutoSize = True, .Left = Esc(360), .Top = y2 + Esc(20), .Font = TemaMecano.FonteCorpo(9.5F)}

        btnConsultar = BotaoMecano.Primario("Consultar", Esc(130), Esc(36))
        btnConsultar.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnConsultar.Top = y2 + Esc(10)
        AddHandler btnConsultar.Click, Sub() Consultar()
        btnExportar = BotaoMecano.Secundario("Exportar", Esc(110), Esc(36))
        btnExportar.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnExportar.Top = y2 + Esc(10)
        btnExportar.Enabled = False
        AddHandler btnExportar.Click, Sub() Exportar()

        pnlFiltros.Controls.AddRange(New Control() {dtpDe, lblA, dtpAte, cboColaborador, cboEmpresa, rbGeral, rbCom, rbSem, btnConsultar, btnExportar})
        AddHandler pnlFiltros.Resize, Sub() PosicionarBotoes()
        AddHandler pnlFiltros.Paint, Sub(s, e)
                                         Using p As New Pen(TemaMecano.BordaCampo)
                                             e.Graphics.DrawLine(p, 0, pnlFiltros.Height - 1, pnlFiltros.Width, pnlFiltros.Height - 1)
                                         End Using
                                     End Sub
        PosicionarBotoes()
    End Sub

    Private Sub PosicionarBotoes()
        Dim direita = pnlFiltros.Width - Esc(24)
        btnConsultar.Left = direita - btnConsultar.Width
        btnExportar.Left = btnConsultar.Left - Esc(8) - btnExportar.Width
    End Sub

    Private Function Rotulo(texto As String, x As Integer, y As Integer) As Label
        Return New Label With {.Text = texto, .AutoSize = True, .Left = x, .Top = y, .ForeColor = TemaMecano.CinzaEscuro, .Font = TemaMecano.FonteCorpo(8.75F)}
    End Function

    ' =================================================================
    '  GRADE
    ' =================================================================
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
            .ColumnHeadersHeight = Esc(38)
            .ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(38, 41, 45)
            .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            .ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(38, 41, 45)
            .ColumnHeadersDefaultCellStyle.Font = TemaMecano.FonteCorpo(9.0F, FontStyle.Bold)
            .ColumnHeadersDefaultCellStyle.Padding = New Padding(Esc(8), 0, 0, 0)
            .DefaultCellStyle.Font = TemaMecano.FonteCorpo(9.5F)
            .DefaultCellStyle.ForeColor = TemaMecano.TextoCampo
            .DefaultCellStyle.BackColor = TemaMecano.Branco
            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(252, 232, 232)
            .DefaultCellStyle.SelectionForeColor = TemaMecano.TextoCampo
            .DefaultCellStyle.Padding = New Padding(Esc(8), 0, Esc(4), 0)
            .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 251, 252)
            .RowTemplate.Height = Esc(32)
            .StandardTab = True
        End With
        AddHandler grade.DataBindingComplete, Sub() grade.ClearSelection()

        lblVazio = New Label With {
            .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleCenter, .Visible = True,
            .ForeColor = TemaMecano.CinzaClaro, .Font = TemaMecano.FonteCorpo(10.0F), .BackColor = TemaMecano.Branco,
            .Text = "Escolha o período e os filtros e clique em Consultar."}
    End Sub

    Private Sub MontarBarraStatus()
        barraStatus = New StatusStrip With {.SizingGrip = False, .BackColor = TemaMecano.Branco, .Font = TemaMecano.FonteCorpo(9.0F)}
        lblStatus = New ToolStripStatusLabel("") With {.ForeColor = TemaMecano.CinzaEscuro}
        lblDica = New ToolStripStatusLabel("F5 consulta  ·  Esc fecha") With {.Spring = True, .TextAlign = ContentAlignment.MiddleRight, .ForeColor = TemaMecano.CinzaClaro}
        barraStatus.Items.AddRange(New ToolStripItem() {lblStatus, lblDica})
    End Sub

    Private Sub MostrarMensagem(texto As String)
        lblStatus.Text = texto
    End Sub

    ' =================================================================
    '  DADOS
    ' =================================================================
    Private Sub CarregarListas()
        If Not ModApontamento.PodeGerenciar() Then
            MessageBox.Show(Me, "Só gestor e diretoria abrem o Gerenciamento do Time.", "Gerenciamento do Time", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
            Return
        End If
        Me.Cursor = Cursors.WaitCursor
        Try
            _colaboradores = ModApontamento.Colaboradores()
            _clientes = ModApontamento.Clientes()
        Catch ex As Exception
            MessageBox.Show(Me, "Não foi possível ler colaboradores e empresas." & vbCrLf & ex.Message, "Gerenciamento do Time", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
        cboColaborador.BeginUpdate()
        cboColaborador.Items.Clear()
        cboColaborador.Items.Add(TODOS)
        For Each c In _colaboradores
            cboColaborador.Items.Add(c)
        Next
        cboColaborador.EndUpdate()
        cboColaborador.SelectedIndex = 0

        cboEmpresa.BeginUpdate()
        cboEmpresa.Items.Clear()
        cboEmpresa.Items.Add(TODAS)
        For Each c In _clientes
            cboEmpresa.Items.Add(c)
        Next
        cboEmpresa.EndUpdate()
        cboEmpresa.SelectedIndex = 0
    End Sub

    Private Function FiltroDaTela() As FiltroAnalitica
        Dim f As New FiltroAnalitica With {.De = dtpDe.Value.Date, .Ate = dtpAte.Value.Date}
        Dim colab = TryCast(cboColaborador.SelectedItem, ColaboradorResumo)
        f.Usuario = If(colab Is Nothing, FiltroAnalitica.TODOS_USUARIOS, colab.Nome)
        Dim emp = TryCast(cboEmpresa.SelectedItem, ClienteResumo)
        f.Empresa = If(emp Is Nothing, FiltroAnalitica.TODAS_EMPRESAS, emp.Nome)
        f.EmpresaId = If(emp Is Nothing, "", emp.Tradutor)
        f.Modo = If(rbCom.Checked, ModoAnalitica.ComCodigoEmpresa, If(rbSem.Checked, ModoAnalitica.SemCodigoEmpresa, ModoAnalitica.Geral))
        Return f
    End Function

    ''' <summary>Só para a captura de tela da demonstração.</summary>
    Public Sub ConsultarAgora()
        Consultar()
    End Sub

    Private Sub Consultar()
        If dtpAte.Value.Date < dtpDe.Value.Date Then
            MessageBox.Show(Me, "A data final é anterior à inicial.", "Gerenciamento do Time", MessageBoxButtons.OK, MessageBoxIcon.Information)
            dtpAte.Focus() : Return
        End If
        Dim f = FiltroDaTela()
        Me.Cursor = Cursors.WaitCursor
        Try
            _linhas = ModApontamento.ConsultaAnalitica(f)
            _filtroAtual = f
        Catch ex As Exception
            _linhas = New List(Of LinhaAnalitica)()
            MessageBox.Show(Me, "Não foi possível consultar." & vbCrLf & ex.Message, "Gerenciamento do Time", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
        Preencher()
    End Sub

    Private Sub Preencher()
        Dim dt As New DataTable()
        dt.Columns.Add("Data", GetType(Date))
        dt.Columns.Add("Dia", GetType(String))
        dt.Columns.Add("Colaborador", GetType(String))
        dt.Columns.Add("Depto", GetType(String))
        dt.Columns.Add("Empresa", GetType(String))
        dt.Columns.Add("Código", GetType(String))
        dt.Columns.Add("Horas", GetType(String))
        For Each l In _linhas
            dt.Rows.Add(l.Data, Apontamento.PtBr.DateTimeFormat.GetAbbreviatedDayName(l.Data.DayOfWeek).TrimEnd("."c), l.Usuario, l.Depto, l.Empresa, If(l.EhCliente, l.EmpresaId, ""), l.HorasTexto)
        Next
        grade.SuspendLayout()
        grade.DataSource = dt
        If grade.Columns.Count > 0 Then
            grade.Columns("Data").FillWeight = 9
            grade.Columns("Data").DefaultCellStyle.Format = "dd/MM/yyyy"
            grade.Columns("Dia").FillWeight = 5
            grade.Columns("Dia").DefaultCellStyle.ForeColor = TemaMecano.CinzaClaro
            grade.Columns("Colaborador").FillWeight = 20
            grade.Columns("Depto").FillWeight = 10
            grade.Columns("Empresa").FillWeight = 36
            grade.Columns("Código").FillWeight = 7
            grade.Columns("Horas").FillWeight = 7
            grade.Columns("Horas").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            grade.Columns("Horas").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
            grade.Columns("Horas").DefaultCellStyle.Font = TemaMecano.FonteCorpo(9.5F, FontStyle.Bold)
        End If
        grade.ResumeLayout()
        grade.ClearSelection()

        lblVazio.Visible = (_linhas.Count = 0)
        lblVazio.Text = "Nenhum apontamento com esses filtros."
        btnExportar.Enabled = (_linhas.Count > 0)
        Dim f = _filtroAtual
        Dim quem = If(f Is Nothing OrElse f.TodosUsuarios, "toda a equipe", f.Usuario)
        MostrarMensagem(_linhas.Count & If(_linhas.Count = 1, " apontamento", " apontamentos") & "  ·  " & Apontamento.FormatarHoras(_linhas.Sum(Function(l) l.Horas)) & " no total  ·  " &
                        quem & If(f Is Nothing, "", "  ·  " & f.De.ToString("dd/MM") & " a " & f.Ate.ToString("dd/MM/yyyy")))
    End Sub

    ' =================================================================
    '  EXPORTAR
    ' =================================================================
    Private Sub Exportar()
        If _linhas.Count = 0 Then Return
        Using dlg As New SaveFileDialog With {.Filter = "CSV (*.csv)|*.csv", .FileName = "time-analitica-" & dtpDe.Value.ToString("yyyy-MM-dd") & "-a-" & dtpAte.Value.ToString("yyyy-MM-dd") & ".csv"}
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim sb As New StringBuilder()
            sb.AppendLine("Data;Dia;Colaborador;Depto;Empresa;Código;Horas;Horas (decimal)")
            For Each l In _linhas.OrderBy(Function(x) x.Data).ThenBy(Function(x) x.Usuario)
                sb.AppendLine(l.Data.ToString("dd/MM/yyyy") & ";" & Apontamento.PtBr.DateTimeFormat.GetAbbreviatedDayName(l.Data.DayOfWeek).TrimEnd("."c) & ";" &
                              l.Usuario & ";" & l.Depto & ";" & l.Empresa.Replace(";", ",") & ";" & If(l.EhCliente, l.EmpresaId, "") & ";" &
                              l.HorasTexto & ";" & Apontamento.FormatarDecimal(l.Horas))
            Next
            sb.AppendLine(";;;;;;Total;" & Apontamento.FormatarDecimal(_linhas.Sum(Function(l) l.Horas)))
            Try
                File.WriteAllText(dlg.FileName, sb.ToString(), New UTF8Encoding(True))
                Process.Start(dlg.FileName)
            Catch ex As Exception
                MessageBox.Show(Me, "Não foi possível gravar o arquivo." & vbCrLf & ex.Message, "Gerenciamento do Time", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

End Class

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
Imports System.Reflection
Imports System.Text
Imports System.Windows.Forms

' ============================================================================
'  frmApontamentos - APONTAMENTOS DE HORAS (o "Time"), tela principal.
'
'  Junta em uma janela o que hoje são duas (Apontamentos > Time > Anotações =
'  frmTimeDet, e > Consulta = frmTime): a lista dos lançamentos do usuário
'  por período, os totais do dia/mês/mês anterior contra a meta da TbHorasMes
'  e os dias pendentes que travam o menu, e o botão para lançar.
'
'  Cada pessoa vê SÓ os próprios apontamentos, como no Time de hoje. As
'  consultas de gestão (frmTimeConsulta, Acumulado, por usuário...) não mudam.
'
'  Montada em código, sem Designer, no visual das telas novas (TemaMecano,
'  UiMecano). Abre com frmApontamentos.Abrir() ou embutida na tela principal
'  (Embutir). Dados: ModApontamento.
'
'  Destino no projeto: Grupo_Mecano\Cadastro\frmApontamentos.vb
' ============================================================================
Public Class frmApontamentos
    Inherits Form

    Private Shared _instancia As frmApontamentos = Nothing

    ''' <summary>Abre a tela (ou traz para a frente a que já está aberta).</summary>
    Public Shared Sub Abrir()
        If _instancia Is Nothing OrElse _instancia.IsDisposed Then
            _instancia = New frmApontamentos()
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
    Private _itens As New List(Of Apontamento)()
    Private _visiveis As New List(Of Apontamento)()
    Private _resumo As ResumoHoras = Nothing
    Private _banco As BancoHoras = Nothing
    Private _textoContagem As String = ""
    Private _embutido As Boolean = False
    Private _selecionadoId As Integer = 0

    ' ---------- controles ----------
    Private pnlHeader As Panel
    Private pnlToolbar As Panel
    Private cboPeriodo As ComboBox
    Private dtpDe As DateTimePicker
    Private dtpAte As DateTimePicker
    Private lblAte As Label
    Private txtBusca As TextBox
    Private btnNovo As BotaoMecano
    Private btnExportar As BotaoMecano
    Private btnExcluir As BotaoMecano
    Private btnPeriodo As BotaoMecano          ' gestão: "Anotações por período" (só gestor e diretoria)
    Private btnGerenciamento As BotaoMecano    ' gestão: as consultas do menu Gerenciamento
    Private pnlResumo As PainelResumoHoras
    Private grade As DataGridView
    Private lblVazio As Label
    Private menuGrade As ContextMenuStrip
    Private barraStatus As StatusStrip
    Private lblStatus As ToolStripStatusLabel
    Private lblAviso As ToolStripStatusLabel
    Private lblDica As ToolStripStatusLabel
    Private WithEvents _timerBusca As New Timer()
    Private WithEvents _timerMensagem As New Timer()

    Private Const TEXTO_BUSCA As String = "Buscar por empresa, código ou atividade"
    Private Const EM_SETCUEBANNER As Integer = &H1501

    <Runtime.InteropServices.DllImport("user32.dll", CharSet:=Runtime.InteropServices.CharSet.Unicode)>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As String) As IntPtr
    End Function

    Public Sub New()
        Using g = Me.CreateGraphics()
            _escala = g.DpiX / 96.0F
        End Using

        Me.AutoScaleMode = AutoScaleMode.None
        Me.Text = "Gerenciador Mêcano — Apontamentos de horas"
        Me.BackColor = TemaMecano.FundoTela
        Me.Font = TemaMecano.FonteCorpo(9.75F)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.ClientSize = New Size(Esc(1280), Esc(760))
        Me.MinimumSize = New Size(Esc(1000), Esc(600))
        Me.KeyPreview = True
        Me.DoubleBuffered = True

        _timerBusca.Interval = 220
        _timerMensagem.Interval = 6000

        MontarGrade()
        MontarResumo()
        MontarToolbar()
        MontarHeader()
        MontarBarraStatus()

        Me.Controls.Add(grade)
        Me.Controls.Add(pnlResumo)
        Me.Controls.Add(pnlToolbar)
        Me.Controls.Add(pnlHeader)
        Me.Controls.Add(barraStatus)
        grade.BringToFront()
        lblVazio.BringToFront()

        AddHandler Me.Shown, Sub()
                                 Carregar()
                                 grade.Focus()
                             End Sub
        AddHandler Me.KeyDown, AddressOf AoTeclar
    End Sub

    ''' <summary>Embutida na tela principal do Gerenciador: sem o cabeçalho escuro próprio e Esc não fecha.</summary>
    Public Sub Embutir()
        _embutido = True
        pnlHeader.Visible = False
    End Sub

    ' =================================================================
    '  CABEÇALHO (faixa escura com logo, padrão das telas novas)
    ' =================================================================
    Private Sub MontarHeader()
        pnlHeader = New Panel With {.Dock = DockStyle.Top, .Height = Esc(64), .BackColor = Color.FromArgb(18, 18, 18)}
        AddHandler pnlHeader.Paint, AddressOf DesenharHeader
        AddHandler pnlHeader.Resize, Sub() pnlHeader.Invalidate()
    End Sub

    Private Sub DesenharHeader(sender As Object, e As PaintEventArgs)
        Dim g = e.Graphics
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
        Dim x = Esc(24)
        If TemaMecano.LogoPrincipal IsNot Nothing Then
            Dim h = Esc(32)
            Dim w = CInt(Math.Round(h * TemaMecano.LogoPrincipal.Width / CDbl(TemaMecano.LogoPrincipal.Height)))
            g.DrawImage(TemaMecano.LogoPrincipal, x, (pnlHeader.Height - h) \ 2, w, h)
            x += w + Esc(20)
        End If
        Using fTitulo = TemaMecano.FonteTitulo(13.0F)
            TextRenderer.DrawText(g, "Apontamentos de horas", fTitulo, New Point(x, Esc(11)), Color.White)
        End Using
        Using fSub = TemaMecano.FonteCorpo(8.5F)
            TextRenderer.DrawText(g, "Seus lançamentos por cliente e atividade  ·  metas do mês  ·  dias pendentes da trava de horas", fSub,
                                  New Point(x + Esc(1), Esc(37)), TemaMecano.CinzaClaro)
        End Using
    End Sub

    ' =================================================================
    '  BARRA DE COMANDO (período, busca, ações)
    ' =================================================================
    Private Sub MontarToolbar()
        pnlToolbar = New Panel With {.Dock = DockStyle.Top, .Height = Esc(62), .BackColor = TemaMecano.Branco}
        AddHandler pnlToolbar.Paint, Sub(s, e)
                                         Using p As New Pen(TemaMecano.BordaCampo)
                                             e.Graphics.DrawLine(p, 0, pnlToolbar.Height - 1, pnlToolbar.Width, pnlToolbar.Height - 1)
                                         End Using
                                     End Sub

        Dim y = Esc(14)
        Dim altura = Esc(34)

        cboPeriodo = New ComboBox With {
            .Left = Esc(24), .Top = y + Esc(3), .Width = Esc(170),
            .DropDownStyle = ComboBoxStyle.DropDownList, .Font = TemaMecano.FonteCorpo(9.75F), .FlatStyle = FlatStyle.System}
        cboPeriodo.Items.AddRange(New Object() {"Hoje", "Esta semana", "Este mês", "Mês anterior", "Últimos 30 dias", "Período…"})
        cboPeriodo.SelectedIndex = CInt(PeriodoLista.EsteMes)
        AddHandler cboPeriodo.SelectedIndexChanged, Sub()
                                                        PosicionarFiltros()
                                                        Carregar()
                                                    End Sub

        dtpDe = New DateTimePicker With {.Top = y + Esc(3), .Width = Esc(112), .Format = DateTimePickerFormat.Short, .Font = TemaMecano.FonteCorpo(9.75F),
                                        .Value = New Date(Date.Today.Year, Date.Today.Month, 1), .Visible = False}
        lblAte = New Label With {.Text = "a", .AutoSize = True, .Top = y + Esc(8), .ForeColor = TemaMecano.CinzaEscuro, .Visible = False}
        dtpAte = New DateTimePicker With {.Top = y + Esc(3), .Width = Esc(112), .Format = DateTimePickerFormat.Short, .Font = TemaMecano.FonteCorpo(9.75F),
                                         .Value = Date.Today, .Visible = False}
        AddHandler dtpDe.ValueChanged, Sub() If cboPeriodo.SelectedIndex = CInt(PeriodoLista.Personalizado) Then Carregar()
        AddHandler dtpAte.ValueChanged, Sub() If cboPeriodo.SelectedIndex = CInt(PeriodoLista.Personalizado) Then Carregar()

        txtBusca = New TextBox With {
            .Top = y + Esc(4), .Width = Esc(300),
            .Font = TemaMecano.FonteCorpo(10.0F), .BorderStyle = BorderStyle.FixedSingle, .ForeColor = TemaMecano.TextoCampo}
        AddHandler txtBusca.HandleCreated, Sub() SendMessage(txtBusca.Handle, EM_SETCUEBANNER, New IntPtr(1), TEXTO_BUSCA)
        AddHandler txtBusca.TextChanged, Sub()
                                             _timerBusca.Stop()
                                             _timerBusca.Start()
                                         End Sub
        AddHandler txtBusca.KeyDown, Sub(s, e)
                                         If e.KeyCode = Keys.Down Then grade.Focus()
                                     End Sub

        btnNovo = BotaoMecano.Primario("+  Novo apontamento", Esc(190), altura)
        btnNovo.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnNovo.Top = y
        AddHandler btnNovo.Click, Sub() Abrir(0, Nothing)

        btnExportar = BotaoMecano.Secundario("Exportar", Esc(100), altura)
        btnExportar.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnExportar.Top = y
        AddHandler btnExportar.Click, Sub() Exportar()

        btnExcluir = BotaoMecano.Secundario("Excluir", Esc(96), altura, perigo:=True)
        btnExcluir.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnExcluir.Top = y
        btnExcluir.Enabled = False
        AddHandler btnExcluir.Click, Sub() Excluir()

        ' gestão (só gestor e diretoria): os outros itens do menu Apontamentos › Time, agora dentro da própria tela
        Dim gestor = ModApontamento.PodeGerenciar()
        btnPeriodo = BotaoMecano.Secundario("Lançar período", Esc(130), altura)
        btnPeriodo.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnPeriodo.Top = y
        btnPeriodo.Visible = gestor
        AddHandler btnPeriodo.Click, Sub() LancarPeriodo()

        btnGerenciamento = BotaoMecano.Secundario("Gerenciamento", Esc(136), altura)
        btnGerenciamento.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnGerenciamento.Top = y
        btnGerenciamento.Visible = gestor
        AddHandler btnGerenciamento.Click, Sub() frmGerenciamentoTime.Abrir()

        pnlToolbar.Controls.AddRange(New Control() {cboPeriodo, dtpDe, lblAte, dtpAte, txtBusca, btnGerenciamento, btnPeriodo, btnExcluir, btnExportar, btnNovo})
        AddHandler pnlToolbar.Resize, Sub() PosicionarAcoes()
        PosicionarFiltros()
        PosicionarAcoes()
    End Sub

    Private Sub PosicionarFiltros()
        Dim personalizado = (cboPeriodo.SelectedIndex = CInt(PeriodoLista.Personalizado))
        dtpDe.Visible = personalizado : lblAte.Visible = personalizado : dtpAte.Visible = personalizado
        Dim x = cboPeriodo.Right + Esc(10)
        If personalizado Then
            dtpDe.Left = x : x = dtpDe.Right + Esc(6)
            lblAte.Left = x : x = lblAte.Right + Esc(6)
            dtpAte.Left = x : x = dtpAte.Right + Esc(12)
        End If
        txtBusca.Left = x
        PosicionarAcoes()
    End Sub

    Private Sub PosicionarAcoes()
        Dim direita = pnlToolbar.Width - Esc(24)
        Dim y = Esc(14)
        ' janela estreita (embutida na tela principal, DPI alto): as ações descem para uma segunda linha
        Dim larguraAcoes = btnNovo.Width + Esc(8) + btnExportar.Width + Esc(8) + btnExcluir.Width
        If btnPeriodo.Visible Then larguraAcoes += Esc(8) + btnPeriodo.Width
        If btnGerenciamento.Visible Then larguraAcoes += Esc(8) + btnGerenciamento.Width
        Dim segundaLinha = (direita - larguraAcoes) < txtBusca.Right + Esc(16)
        Dim topo = If(segundaLinha, y + Esc(44), y)
        Dim altura = If(segundaLinha, Esc(62 + 44), Esc(62))
        If pnlToolbar.Height <> altura Then pnlToolbar.Height = altura
        btnNovo.Top = topo : btnExportar.Top = topo : btnExcluir.Top = topo : btnPeriodo.Top = topo : btnGerenciamento.Top = topo
        btnNovo.Left = direita - btnNovo.Width
        btnExportar.Left = btnNovo.Left - Esc(8) - btnExportar.Width
        btnExcluir.Left = btnExportar.Left - Esc(8) - btnExcluir.Width
        Dim x = btnExcluir.Left
        If btnPeriodo.Visible Then
            btnPeriodo.Left = x - Esc(8) - btnPeriodo.Width
            x = btnPeriodo.Left
        End If
        If btnGerenciamento.Visible Then btnGerenciamento.Left = x - Esc(8) - btnGerenciamento.Width
    End Sub

    ''' <summary>"Anotações por período": férias/afastamento no nome de outra pessoa (só gestor e diretoria).</summary>
    Private Sub LancarPeriodo()
        Using f As New frmApontamentoPeriodo()
            If f.ShowDialog(Me) = DialogResult.OK Then
                Carregar()
                MostrarMensagem(f.MensagemResultado)
            End If
        End Using
    End Sub

    ' =================================================================
    '  RESUMO (cartões: hoje, mês, mês anterior, dias pendentes)
    ' =================================================================
    Private Sub MontarResumo()
        pnlResumo = New PainelResumoHoras(_escala) With {.Dock = DockStyle.Top, .Height = Esc(150)}
        AddHandler pnlResumo.DiaClicado, Sub(dia) Abrir(0, dia)
        AddHandler pnlResumo.BancoClicado, Sub() AbrirBanco()
    End Sub

    ''' <summary>Extrato do banco de horas: os dias em que a pessoa passou da jornada, desde BancoHoras.INICIO_VIGENCIA.</summary>
    Private Sub AbrirBanco()
        Using f As New frmBancoHoras(_banco)
            f.ShowDialog(Me)
        End Using
    End Sub

    ' =================================================================
    '  GRADE
    ' =================================================================
    Private Sub MontarGrade()
        grade = New DataGridView With {.Dock = DockStyle.Fill}
        Try
            GetType(DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic).SetValue(grade, True, Nothing)
        Catch
        End Try
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
            .RowTemplate.Height = Esc(34)
            .StandardTab = True
        End With
        AddHandler grade.CellDoubleClick, Sub(s, e)
                                              If e.RowIndex >= 0 Then Abrir(IdDaLinha(e.RowIndex), Nothing)
                                          End Sub
        AddHandler grade.SelectionChanged, Sub()
                                               btnExcluir.Enabled = (grade.CurrentRow IsNot Nothing)
                                               If grade.CurrentRow IsNot Nothing Then _selecionadoId = IdDaLinha(grade.CurrentRow.Index)
                                           End Sub
        AddHandler grade.KeyDown, Sub(s, e)
                                      If e.KeyCode = Keys.Enter AndAlso grade.CurrentRow IsNot Nothing Then
                                          e.Handled = True : e.SuppressKeyPress = True
                                          Abrir(IdDaLinha(grade.CurrentRow.Index), Nothing)
                                      ElseIf e.KeyCode = Keys.Delete AndAlso grade.CurrentRow IsNot Nothing Then
                                          e.Handled = True
                                          Excluir()
                                      End If
                                  End Sub
        AddHandler grade.CellMouseDown, Sub(s, e)
                                            If e.Button = MouseButtons.Right AndAlso e.RowIndex >= 0 Then
                                                grade.ClearSelection()
                                                grade.Rows(e.RowIndex).Selected = True
                                                grade.CurrentCell = grade.Rows(e.RowIndex).Cells("Empresa")
                                            End If
                                        End Sub
        AddHandler grade.CellFormatting, Sub(s, e)
                                             If e.RowIndex < 0 Then Return
                                             If grade.Columns(e.ColumnIndex).Name = "Tipo" AndAlso CStr(e.Value) = "Interno" Then
                                                 e.CellStyle.ForeColor = TemaMecano.CinzaClaro
                                             End If
                                         End Sub

        menuGrade = New ContextMenuStrip With {.Font = TemaMecano.FonteCorpo(9.5F)}
        menuGrade.Items.Add("Editar…", Nothing, Sub() If grade.CurrentRow IsNot Nothing Then Abrir(IdDaLinha(grade.CurrentRow.Index), Nothing))
        menuGrade.Items.Add("Excluir…", Nothing, Sub() Excluir())
        menuGrade.Items.Add(New ToolStripSeparator())
        menuGrade.Items.Add("Novo apontamento nesta data", Nothing, Sub()
                                                                         If grade.CurrentRow IsNot Nothing Then Abrir(0, DataDaLinha(grade.CurrentRow.Index))
                                                                     End Sub)
        grade.ContextMenuStrip = menuGrade

        lblVazio = New Label With {
            .AutoSize = False, .TextAlign = ContentAlignment.MiddleCenter, .BackColor = TemaMecano.Branco,
            .ForeColor = TemaMecano.CinzaClaro, .Font = TemaMecano.FonteCorpo(10.5F), .Visible = False,
            .Text = "Nenhum apontamento no período."}
        AddHandler grade.Resize, Sub() PosicionarVazio()
        Me.Controls.Add(lblVazio)
    End Sub

    Private Sub PosicionarVazio()
        lblVazio.Bounds = New Rectangle(grade.Left, grade.Top + grade.ColumnHeadersHeight, grade.Width, Math.Max(Esc(60), grade.Height - grade.ColumnHeadersHeight))
    End Sub

    Private Function IdDaLinha(indice As Integer) As Integer
        Return Convert.ToInt32(grade.Rows(indice).Cells("Id").Value)
    End Function

    Private Function DataDaLinha(indice As Integer) As Date
        Return Convert.ToDateTime(grade.Rows(indice).Cells("Data").Value).Date
    End Function

    ' =================================================================
    '  BARRA DE STATUS
    ' =================================================================
    Private Sub MontarBarraStatus()
        barraStatus = New StatusStrip With {.SizingGrip = False, .BackColor = TemaMecano.Branco, .Font = TemaMecano.FonteCorpo(8.75F)}
        lblStatus = New ToolStripStatusLabel With {.ForeColor = TemaMecano.CinzaEscuro, .TextAlign = ContentAlignment.MiddleLeft}
        lblAviso = New ToolStripStatusLabel With {.ForeColor = Color.FromArgb(133, 79, 11), .Spring = True, .TextAlign = ContentAlignment.MiddleLeft}
        lblDica = New ToolStripStatusLabel With {.ForeColor = TemaMecano.CinzaClaro,
            .Text = "Duplo clique ou Enter edita  ·  Ctrl+N novo  ·  Del exclui  ·  Ctrl+F busca  ·  F5 atualiza  ·  Esc fecha"}
        barraStatus.Items.AddRange(New ToolStripItem() {lblStatus, lblAviso, lblDica})
        Dim aviso = ModApontamento.DescricaoFonteDados()
        lblAviso.Text = If(aviso = "", "", "   " & aviso)
    End Sub

    Private Sub MostrarMensagem(texto As String)
        lblStatus.Text = texto
        lblStatus.ForeColor = Color.FromArgb(22, 101, 52)
        _timerMensagem.Stop()
        _timerMensagem.Start()
    End Sub

    Private Sub _timerMensagem_Tick(sender As Object, e As EventArgs) Handles _timerMensagem.Tick
        _timerMensagem.Stop()
        lblStatus.ForeColor = TemaMecano.CinzaEscuro
        lblStatus.Text = _textoContagem
    End Sub

    Private Sub _timerBusca_Tick(sender As Object, e As EventArgs) Handles _timerBusca.Tick
        _timerBusca.Stop()
        AplicarFiltro()
    End Sub

    ' =================================================================
    '  DADOS
    ' =================================================================
    Private Function PeriodoAtual() As Date()
        Dim hoje = Date.Today
        Select Case CType(cboPeriodo.SelectedIndex, PeriodoLista)
            Case PeriodoLista.Hoje
                Return {hoje, hoje}
            Case PeriodoLista.EstaSemana
                Dim desloc = (CInt(hoje.DayOfWeek) + 6) Mod 7   ' segunda = 0
                Dim segunda = hoje.AddDays(-desloc)
                Return {segunda, segunda.AddDays(6)}
            Case PeriodoLista.EsteMes
                Dim ini = New Date(hoje.Year, hoje.Month, 1)
                Return {ini, ini.AddMonths(1).AddDays(-1)}
            Case PeriodoLista.MesAnterior
                Dim ini = New Date(hoje.Year, hoje.Month, 1).AddMonths(-1)
                Return {ini, ini.AddMonths(1).AddDays(-1)}
            Case PeriodoLista.Ultimos30Dias
                Return {hoje.AddDays(-29), hoje}
            Case Else
                Dim de = dtpDe.Value.Date
                Dim ate = dtpAte.Value.Date
                If ate < de Then ate = de
                Return {de, ate}
        End Select
    End Function

    Public Sub Carregar()
        Me.Cursor = Cursors.WaitCursor
        Try
            _resumo = ModApontamento.CalcularResumo()
            _banco = ModApontamento.CalcularBanco()
            pnlResumo.Banco = _banco
            pnlResumo.Resumo = _resumo
            Dim p = PeriodoAtual()
            _itens = ModApontamento.Listar(p(0), p(1))
        Catch ex As Exception
            _itens = New List(Of Apontamento)()
            MessageBox.Show(Me, "Não foi possível ler os apontamentos." & vbCrLf & ex.Message, "Apontamentos",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
        AplicarFiltro()
    End Sub

    Private Sub AplicarFiltro()
        Dim termo = Colaborador.SemAcento(txtBusca.Text.Trim())
        _visiveis = _itens.Where(Function(a)
                                     If termo = "" Then Return True
                                     Return Colaborador.SemAcento(a.Empresa).Contains(termo) OrElse
                                            (a.EmpresaId <> "" AndAlso a.EmpresaId.TrimStart("0"c).StartsWith(termo.TrimStart("0"c)))
                                 End Function).
                    OrderByDescending(Function(a) a.Data).ThenByDescending(Function(a) a.Id).ToList()

        Dim dt As New DataTable()
        dt.Columns.Add("Id", GetType(Integer))
        dt.Columns.Add("Data", GetType(Date))
        dt.Columns.Add("Dia", GetType(String))
        dt.Columns.Add("Empresa", GetType(String))
        dt.Columns.Add("Código", GetType(String))
        dt.Columns.Add("Tipo", GetType(String))
        dt.Columns.Add("Horas", GetType(String))
        For Each a In _visiveis
            dt.Rows.Add(a.Id, a.Data.Date, Apontamento.PtBr.DateTimeFormat.GetAbbreviatedDayName(a.Data.DayOfWeek).TrimEnd("."c),
                        a.Empresa, If(a.EhCliente, a.EmpresaId, ""), a.TipoTexto, a.HorasTexto)
        Next

        grade.SuspendLayout()
        grade.DataSource = dt
        If grade.Columns.Count > 0 Then
            grade.Columns("Id").Visible = False
            grade.Columns("Data").FillWeight = 10
            grade.Columns("Data").DefaultCellStyle.Format = "dd/MM/yyyy"
            grade.Columns("Dia").FillWeight = 5
            grade.Columns("Dia").DefaultCellStyle.ForeColor = TemaMecano.CinzaClaro
            grade.Columns("Empresa").FillWeight = 52
            grade.Columns("Código").FillWeight = 8
            grade.Columns("Tipo").FillWeight = 9
            grade.Columns("Horas").FillWeight = 8
            grade.Columns("Horas").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            grade.Columns("Horas").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
            grade.Columns("Horas").DefaultCellStyle.Font = TemaMecano.FonteCorpo(9.5F, FontStyle.Bold)
            grade.ClearSelection()
            If _selecionadoId > 0 Then
                For Each linha As DataGridViewRow In grade.Rows
                    If IdDaLinha(linha.Index) = _selecionadoId Then
                        linha.Selected = True
                        grade.CurrentCell = linha.Cells("Empresa")
                        Exit For
                    End If
                Next
            End If
        End If
        grade.ResumeLayout()
        btnExcluir.Enabled = (grade.CurrentRow IsNot Nothing AndAlso grade.SelectedRows.Count > 0)

        lblVazio.Text = If(_itens.Count = 0, "Nenhum apontamento no período.", "Nenhum apontamento encontrado com essa busca.")
        lblVazio.Visible = (_visiveis.Count = 0)
        PosicionarVazio()

        Dim total = _visiveis.Sum(Function(a) a.Horas)
        Dim n = _visiveis.Count
        Dim p = PeriodoAtual()
        Dim periodo = If(p(0) = p(1), p(0).ToString("dd/MM"), p(0).ToString("dd/MM") & " a " & p(1).ToString("dd/MM"))
        _textoContagem = If(n = 0, "Nenhum apontamento em " & periodo,
                            String.Format("{0} apontamento{1}  ·  {2} em {3}", n, If(n = 1, "", "s"), Apontamento.FormatarHoras(total), periodo))
        If Not _timerMensagem.Enabled Then lblStatus.Text = _textoContagem
    End Sub

    ' =================================================================
    '  AÇÕES
    ' =================================================================
    Private Sub Abrir(id As Integer, dataInicial As Date?)
        Using f As New frmApontamentoDet(id, dataInicial, _resumo)
            If f.ShowDialog(Me) = DialogResult.OK Then
                _selecionadoId = f.IdGravado
                Carregar()
                If f.MensagemResultado <> "" Then MostrarMensagem(f.MensagemResultado)
            Else
                _selecionadoId = id
            End If
        End Using
    End Sub

    Private Sub Excluir()
        If grade.CurrentRow Is Nothing Then Return
        Dim id = IdDaLinha(grade.CurrentRow.Index)
        Dim a = _visiveis.FirstOrDefault(Function(x) x.Id = id)
        If a Is Nothing Then Return
        If MessageBox.Show(Me, "Excluir o apontamento de " & Apontamento.FormatarData(a.Data) & "?" & vbCrLf & vbCrLf &
                           a.Empresa & "  ·  " & a.HorasTexto, "Excluir apontamento",
                           MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then Return
        Try
            ModApontamento.Excluir(id)
            ModApontamento.AposGravar()
        Catch ex As Exception
            MessageBox.Show(Me, "Não foi possível excluir." & vbCrLf & ex.Message, "Excluir apontamento", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try
        _selecionadoId = 0
        Carregar()
        MostrarMensagem("Apontamento excluído: " & a.HorasTexto & " em " & a.Empresa & " (" & a.Data.ToString("dd/MM") & ").")
    End Sub

    ''' <summary>Exporta a lista filtrada em CSV (abre direto no Excel).</summary>
    Private Sub Exportar()
        If _visiveis.Count = 0 Then
            MessageBox.Show(Me, "Não há apontamentos na lista para exportar.", "Exportar", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Using dlg As New SaveFileDialog With {
                .Filter = "Planilha CSV para Excel (*.csv)|*.csv", .DefaultExt = "csv",
                .FileName = "Apontamentos-" & Date.Today.ToString("yyyy-MM-dd") & ".csv",
                .Title = "Exportar apontamentos"}
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Try
                Dim sb As New StringBuilder()
                sb.AppendLine("Data;Dia;Empresa / atividade;Código;Tipo;Horas;Horas (decimal)")
                For Each a In _visiveis.OrderBy(Function(x) x.Data).ThenBy(Function(x) x.Id)
                    sb.AppendLine(String.Join(";", New String() {
                        a.Data.ToString("dd/MM/yyyy"), Apontamento.PtBr.DateTimeFormat.GetAbbreviatedDayName(a.Data.DayOfWeek).TrimEnd("."c),
                        Csv(a.Empresa), Csv(a.EmpresaId), a.TipoTexto, a.HorasTexto, Apontamento.FormatarDecimal(a.Horas)}))
                Next
                sb.AppendLine(String.Join(";", New String() {"", "", "TOTAL", "", "", Apontamento.FormatarHoras(_visiveis.Sum(Function(x) x.Horas)),
                                                              Apontamento.FormatarDecimal(_visiveis.Sum(Function(x) x.Horas))}))
                File.WriteAllText(dlg.FileName, sb.ToString(), New UTF8Encoding(True))
                MostrarMensagem("Exportado: " & dlg.FileName)
                Try
                    Process.Start(dlg.FileName)
                Catch
                End Try
            Catch ex As Exception
                MessageBox.Show(Me, "Não foi possível exportar." & vbCrLf & ex.Message, "Exportar", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    Private Function Csv(valor As String) As String
        Dim v = If(valor, "")
        If v.IndexOfAny(New Char() {";"c, """"c, ControlChars.Cr, ControlChars.Lf}) >= 0 Then
            Return """" & v.Replace("""", """""") & """"
        End If
        Return v
    End Function

    Private Sub AoTeclar(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Escape Then
            If txtBusca.Text.Trim() <> "" Then
                txtBusca.Text = ""
                txtBusca.Focus()
            ElseIf Not _embutido Then
                Me.Close()
            End If
            e.Handled = True
        ElseIf e.KeyCode = Keys.F5 Then
            ModApontamento.LimparCache()
            Carregar()
            e.Handled = True
        ElseIf e.Control AndAlso e.KeyCode = Keys.F Then
            txtBusca.Focus()
            txtBusca.SelectAll()
            e.Handled = True : e.SuppressKeyPress = True
        ElseIf e.Control AndAlso e.KeyCode = Keys.N Then
            Abrir(0, Nothing)
            e.Handled = True : e.SuppressKeyPress = True
        End If
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            _timerBusca.Dispose()
            _timerMensagem.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

End Class

' ============================================================================
'  PainelResumoHoras - os cinco cartões do alto da tela, desenhados:
'    HOJE · ESTE MÊS (com o gráfico dia a dia; a barra vermelha é hoje) ·
'    MÊS ANTERIOR · DIAS COM MENOS DE 8 H (clique no dia para lançar) ·
'    BANCO DE HORAS (saldo do que passou de 8 h; clique abre o extrato).
'  É o cartão "Horas do mês" da proposta da tela principal, aplicado aqui.
' ============================================================================
Friend Class PainelResumoHoras
    Inherits Panel

    Public Event DiaClicado(dia As Date)
    Public Event BancoClicado()

    Private ReadOnly _escala As Single
    Private _resumo As ResumoHoras = Nothing
    Private _banco As BancoHoras = Nothing
    Private _rBanco As Rectangle = Rectangle.Empty
    Private ReadOnly _chips As New List(Of KeyValuePair(Of Rectangle, Date))()
    Private ReadOnly _verde As Color = Color.FromArgb(22, 101, 52)
    Private ReadOnly _verdeFundo As Color = Color.FromArgb(222, 244, 231)
    Private ReadOnly _vermelhoFundo As Color = Color.FromArgb(252, 232, 232)

    Public Sub New(escala As Single)
        _escala = escala
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint Or ControlStyles.ResizeRedraw Or ControlStyles.OptimizedDoubleBuffer, True)
        Me.BackColor = TemaMecano.FundoTela
    End Sub

    Private Function Esc(v As Integer) As Integer
        Return CInt(Math.Round(v * _escala))
    End Function

    Public Property Resumo As ResumoHoras
        Get
            Return _resumo
        End Get
        Set(value As ResumoHoras)
            _resumo = value
            Invalidate()
        End Set
    End Property

    Public Property Banco As BancoHoras
        Get
            Return _banco
        End Get
        Set(value As BancoHoras)
            _banco = value
            Invalidate()
        End Set
    End Property

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        Me.Cursor = If(_chips.Any(Function(c) c.Key.Contains(e.Location)) OrElse _rBanco.Contains(e.Location), Cursors.Hand, Cursors.Default)
    End Sub

    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        MyBase.OnMouseClick(e)
        If e.Button <> MouseButtons.Left Then Return
        If _rBanco.Contains(e.Location) Then
            RaiseEvent BancoClicado()
            Return
        End If
        For Each c In _chips
            If c.Key.Contains(e.Location) Then
                RaiseEvent DiaClicado(c.Value)
                Return
            End If
        Next
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
        g.Clear(TemaMecano.FundoTela)
        _chips.Clear()
        Dim r = _resumo
        If r Is Nothing Then
            Using f = TemaMecano.FonteCorpo(9.5F)
                TextRenderer.DrawText(g, "Carregando os totais…", f, New Point(Esc(24), Esc(24)), TemaMecano.CinzaClaro)
            End Using
            Return
        End If

        Dim margem = Esc(24), gap = Esc(12), topo = Esc(12)
        Dim altura = Height - Esc(24)
        Dim wHoje = Esc(170), wAnt = Esc(210), wPend = Esc(290), wBanco = Esc(190)
        Dim wMes = Width - 2 * margem - wHoje - wAnt - wPend - wBanco - 4 * gap
        If wMes < Esc(300) Then
            wPend = Math.Max(Esc(200), wPend - (Esc(300) - wMes))
            wMes = Width - 2 * margem - wHoje - wAnt - wPend - wBanco - 4 * gap
        End If
        Dim x = margem
        Dim rHoje As New Rectangle(x, topo, wHoje, altura) : x += wHoje + gap
        Dim rMes As New Rectangle(x, topo, Math.Max(Esc(120), wMes), altura) : x += rMes.Width + gap
        Dim rAnt As New Rectangle(x, topo, wAnt, altura) : x += wAnt + gap
        Dim rPend As New Rectangle(x, topo, wPend, altura) : x += wPend + gap
        Dim rBanco As New Rectangle(x, topo, wBanco, altura)
        _rBanco = rBanco

        ' ---- HOJE
        Cartao(g, rHoje, "HOJE")
        Dim corHoje = If(r.HorasHoje >= ResumoHoras.JORNADA, _verde, TemaMecano.TextoCampo)
        Grande(g, rHoje, Apontamento.FormatarHoras(r.HorasHoje), corHoje)
        Pequeno(g, rHoje, Esc(74), "de " & Apontamento.FormatarHoras(ResumoHoras.JORNADA) & If(r.HorasHoje >= ResumoHoras.JORNADA, "  ·  dia fechado", ""), TemaMecano.CinzaEscuro)
        Barra(g, New Rectangle(rHoje.X + Esc(16), rHoje.Bottom - Esc(22), rHoje.Width - Esc(32), Esc(6)), r.HorasHoje, ResumoHoras.JORNADA, If(r.HorasHoje >= ResumoHoras.JORNADA, _verde, TemaMecano.VermelhoPrimario))

        ' ---- ESTE MÊS (números à esquerda, gráfico dia a dia à direita)
        Cartao(g, rMes, "ESTE MÊS  ·  " & Apontamento.NomeCompetencia(r.Competencia).ToUpperInvariant())
        Dim larguraNumeros = Esc(200)
        Grande(g, New Rectangle(rMes.X, rMes.Y, larguraNumeros, rMes.Height), Apontamento.FormatarHoras(r.MesApontado), TemaMecano.TextoCampo)
        Dim subMes As String
        Dim corSubMes = TemaMecano.CinzaEscuro
        If r.MesMeta <= 0D Then
            subMes = "meta do mês não cadastrada"
        ElseIf r.MesAApontar > 0D Then
            subMes = "de " & Apontamento.FormatarHoras(r.MesMeta) & "  ·  faltam " & Apontamento.FormatarHoras(r.MesAApontar)
        Else
            subMes = "de " & Apontamento.FormatarHoras(r.MesMeta) & "  ·  meta atingida"
            corSubMes = _verde
        End If
        Pequeno(g, New Rectangle(rMes.X, rMes.Y, larguraNumeros, rMes.Height), Esc(74), subMes, corSubMes)
        Barra(g, New Rectangle(rMes.X + Esc(16), rMes.Bottom - Esc(22), larguraNumeros - Esc(24), Esc(6)), r.MesApontado, r.MesMeta, TemaMecano.VermelhoPrimario)
        Grafico(g, New Rectangle(rMes.X + larguraNumeros + Esc(4), rMes.Y + Esc(34), rMes.Width - larguraNumeros - Esc(20), rMes.Height - Esc(50)), r)

        ' ---- MÊS ANTERIOR
        Cartao(g, rAnt, "MÊS ANTERIOR  ·  " & Apontamento.NomeCompetencia(r.CompetenciaAnterior).ToUpperInvariant())
        Dim subAnt As String
        Dim corAnt = TemaMecano.CinzaEscuro
        Dim corGrandeAnt = TemaMecano.TextoCampo
        If Not r.MesAnteriorExigido Then
            subAnt = "de " & Apontamento.FormatarHoras(r.MesAnteriorMeta) & "  ·  sem cobrança (admissão)"
        ElseIf r.MesAnteriorMeta <= 0D Then
            subAnt = "meta do mês não cadastrada"
        ElseIf r.MesAnteriorAApontar > 0D Then
            subAnt = "de " & Apontamento.FormatarHoras(r.MesAnteriorMeta) & "  ·  faltam " & Apontamento.FormatarHoras(r.MesAnteriorAApontar)
            corAnt = TemaMecano.VermelhoPrimario
            corGrandeAnt = TemaMecano.VermelhoPrimario
        Else
            subAnt = "de " & Apontamento.FormatarHoras(r.MesAnteriorMeta) & "  ·  fechado"
            corAnt = _verde
        End If
        Grande(g, rAnt, Apontamento.FormatarHoras(r.MesAnteriorApontado), corGrandeAnt)
        Pequeno(g, rAnt, Esc(74), subAnt, corAnt)
        Barra(g, New Rectangle(rAnt.X + Esc(16), rAnt.Bottom - Esc(22), rAnt.Width - Esc(32), Esc(6)), r.MesAnteriorApontado, r.MesAnteriorMeta,
              If(r.MesAnteriorExigido AndAlso r.MesAnteriorAApontar > 0D, TemaMecano.VermelhoPrimario, _verde))

        ' ---- DIAS PENDENTES
        Cartao(g, rPend, "MENOS DE " & CInt(ResumoHoras.JORNADA) & " H  ·  ÚLTIMOS " & ResumoHoras.DIAS_VERIFICADOS & " DIAS")
        Using fChip = TemaMecano.FonteCorpo(8.25F, FontStyle.Bold)
            If r.DiasPendentes.Count = 0 Then
                FormasMecano.DesenharChip(g, "NENHUM DIA PENDENTE", rPend.X + Esc(16), rPend.Y + Esc(40), Esc(22), fChip, _verdeFundo, _verde)
                Pequeno(g, rPend, Esc(74), "Os menus não travam por dias em aberto.", TemaMecano.CinzaEscuro)
            Else
                ' Chips em linhas; a última linha reserva espaço para o "+N" dos dias que não couberam,
                ' para ele ficar ao lado dos chips e não em cima do aviso do rodapé.
                Dim esquerda = rPend.X + Esc(16)
                Dim limite = rPend.Right - Esc(16)
                Dim fundoChips = rPend.Bottom - Esc(26)          ' acima daqui fica o aviso "N dias travam o menu"
                Dim larguraMais = TextRenderer.MeasureText(g, "+" & r.DiasPendentes.Count, fChip).Width + Esc(11)
                Dim cx = esquerda
                Dim cy = rPend.Y + Esc(38)
                Dim mostrados = 0
                For Each p In r.DiasPendentes
                    Dim texto = Apontamento.FormatarData(p.Data).Substring(0, 9).Replace("/" & p.Data.Year.ToString(), "") & " · " & Apontamento.FormatarHoras(p.Horas)
                    Dim largura = TextRenderer.MeasureText(g, texto, fChip).Width + Esc(11)
                    Dim faltamDepois = r.DiasPendentes.Count - mostrados - 1
                    Dim ultimaLinha = cy + Esc(27) + Esc(22) > fundoChips
                    Dim reserva = If(ultimaLinha AndAlso faltamDepois > 0, larguraMais + Esc(6), 0)
                    If cx + largura + reserva > limite Then
                        If ultimaLinha Then Exit For
                        cx = esquerda
                        cy += Esc(27)
                        ultimaLinha = cy + Esc(27) + Esc(22) > fundoChips
                        reserva = If(ultimaLinha AndAlso faltamDepois > 0, larguraMais + Esc(6), 0)
                        If cx + largura + reserva > limite Then Exit For
                    End If
                    Dim usado = FormasMecano.DesenharChip(g, texto, cx, cy, Esc(22), fChip, _vermelhoFundo, TemaMecano.VermelhoPrimario)
                    _chips.Add(New KeyValuePair(Of Rectangle, Date)(New Rectangle(cx, cy, usado, Esc(22)), p.Data))
                    cx += usado + Esc(6)
                    mostrados += 1
                Next
                Dim restantes = r.DiasPendentes.Count - mostrados
                If restantes > 0 Then FormasMecano.DesenharChip(g, "+" & restantes, cx, cy, Esc(22), fChip, Color.FromArgb(240, 242, 245), TemaMecano.CinzaEscuro)
                Using fInfo = TemaMecano.FonteCorpo(8.5F)
                    Dim aviso = r.DiasPendentes.Count & If(r.DiasPendentes.Count = 1, " dia trava o menu", " dias travam o menu") & "  ·  clique no dia"
                    TextRenderer.DrawText(g, aviso, fInfo, New Rectangle(rPend.X + Esc(16), rPend.Bottom - Esc(28), rPend.Width - Esc(32), Esc(18)),
                                          TemaMecano.VermelhoPrimario, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPrefix)
                End Using
            End If
        End Using

        ' ---- BANCO DE HORAS (o que passou da jornada por dia; clique abre o extrato)
        Cartao(g, rBanco, "BANCO DE HORAS")
        Dim b = _banco
        If b Is Nothing Then
            Pequeno(g, rBanco, Esc(74), "Carregando…", TemaMecano.CinzaClaro)
        Else
            Dim saldo = b.Saldo
            Grande(g, rBanco, If(saldo > 0D, "+", "") & Apontamento.FormatarHoras(saldo), If(saldo > 0D, _verde, TemaMecano.TextoCampo))
            Dim n = b.Dias.Count
            Pequeno(g, rBanco, Esc(74), If(n = 0, "nenhum dia acima de " & Apontamento.FormatarHoras(ResumoHoras.JORNADA),
                                            n & If(n = 1, " dia acima de ", " dias acima de ") & Apontamento.FormatarHoras(ResumoHoras.JORNADA)), TemaMecano.CinzaEscuro)
            Pequeno(g, rBanco, rBanco.Height - Esc(28), "clique para ver o extrato", TemaMecano.VermelhoPrimario)
        End If
    End Sub

    Private Sub Cartao(g As Graphics, r As Rectangle, titulo As String)
        Using gp = FormasMecano.Arredondado(New Rectangle(r.X, r.Y, r.Width - 1, r.Height - 1), Esc(10))
            Using b As New SolidBrush(TemaMecano.Branco)
                g.FillPath(b, gp)
            End Using
            Using p As New Pen(TemaMecano.BordaCampo)
                g.DrawPath(p, gp)
            End Using
        End Using
        Using f = TemaMecano.FonteCorpo(7.75F, FontStyle.Bold)
            TextRenderer.DrawText(g, titulo, f, New Rectangle(r.X + Esc(16), r.Y + Esc(12), r.Width - Esc(32), Esc(16)), TemaMecano.CinzaClaro,
                                  TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPrefix)
        End Using
    End Sub

    ''' <summary>Número grande do cartão ("138h00"). Segoe UI negrito, como os chips e o corpo da tela: a fonte da marca (Neulis Neue) fica só nos títulos, como nas outras telas novas.</summary>
    Private Sub Grande(g As Graphics, r As Rectangle, texto As String, cor As Color)
        Using f = TemaMecano.FonteCorpo(19.0F, FontStyle.Bold)
            TextRenderer.DrawText(g, texto, f, New Rectangle(r.X + Esc(16), r.Y + Esc(32), r.Width - Esc(32), Esc(38)), cor,
                                  TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPrefix)
        End Using
    End Sub

    Private Sub Pequeno(g As Graphics, r As Rectangle, y As Integer, texto As String, cor As Color)
        Using f = TemaMecano.FonteCorpo(8.5F)
            TextRenderer.DrawText(g, texto, f, New Rectangle(r.X + Esc(16), r.Y + y, r.Width - Esc(32), Esc(18)), cor,
                                  TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPrefix)
        End Using
    End Sub

    Private Sub Barra(g As Graphics, r As Rectangle, valor As Decimal, meta As Decimal, cor As Color)
        Using gp = FormasMecano.Arredondado(r, r.Height \ 2)
            Using b As New SolidBrush(Color.FromArgb(236, 238, 241))
                g.FillPath(b, gp)
            End Using
        End Using
        If meta <= 0D OrElse valor <= 0D Then Return
        Dim fracao = CSng(Math.Min(1D, valor / meta))
        Dim w = Math.Max(r.Height, CInt(r.Width * fracao))
        Using gp = FormasMecano.Arredondado(New Rectangle(r.X, r.Y, w, r.Height), r.Height \ 2)
            Using b As New SolidBrush(cor)
                g.FillPath(b, gp)
            End Using
        End Using
    End Sub

    ''' <summary>Gráfico dia a dia do mês: uma barra por dia; fins de semana claros; hoje em vermelho; linha na jornada.</summary>
    Private Sub Grafico(g As Graphics, area As Rectangle, r As ResumoHoras)
        If area.Width < Esc(80) OrElse area.Height < Esc(30) Then Return
        Dim mes = New Date(r.Hoje.Year, r.Hoje.Month, 1)
        Dim dias = Date.DaysInMonth(mes.Year, mes.Month)
        Dim alturaRotulo = Esc(14)
        Dim base = area.Bottom - alturaRotulo
        Dim alturaUtil = base - area.Y
        Dim maximo = Math.Max(10D, If(r.PorDiaDoMes.Count = 0, 10D, r.PorDiaDoMes.Values.Max()))
        Dim passo = area.Width / CSng(dias)
        Dim larguraBarra = Math.Max(2, CInt(Math.Floor(passo * 0.62F)))

        ' linha da jornada (8 h)
        Dim yJornada = base - CInt(alturaUtil * CSng(ResumoHoras.JORNADA / maximo))
        Using p As New Pen(Color.FromArgb(220, 223, 228)) With {.DashStyle = DashStyle.Dash}
            g.DrawLine(p, area.X, yJornada, area.Right, yJornada)
        End Using

        Using fRot = TemaMecano.FonteCorpo(7.0F)
            For dia = 1 To dias
                Dim d = New Date(mes.Year, mes.Month, dia)
                Dim fimDeSemana = (d.DayOfWeek = DayOfWeek.Saturday OrElse d.DayOfWeek = DayOfWeek.Sunday)
                Dim x = area.X + CInt((dia - 1) * passo + (passo - larguraBarra) / 2)
                Dim horas As Decimal = 0D
                r.PorDiaDoMes.TryGetValue(dia, horas)
                Dim h = CInt(alturaUtil * CSng(horas / maximo))
                Dim cor As Color
                If d = r.Hoje Then
                    cor = TemaMecano.VermelhoPrimario
                ElseIf d > r.Hoje Then
                    cor = Color.FromArgb(240, 242, 245)
                ElseIf fimDeSemana Then
                    cor = Color.FromArgb(208, 212, 217)
                ElseIf horas < ResumoHoras.JORNADA Then
                    cor = Color.FromArgb(232, 160, 165)
                Else
                    cor = TemaMecano.CinzaEscuro
                End If
                If d > r.Hoje Then
                    Using b As New SolidBrush(cor)
                        g.FillRectangle(b, x, base - Esc(3), larguraBarra, Esc(3))
                    End Using
                ElseIf h <= 0 Then
                    Using b As New SolidBrush(If(fimDeSemana, Color.FromArgb(232, 234, 237), Color.FromArgb(232, 160, 165)))
                        g.FillRectangle(b, x, base - Esc(3), larguraBarra, Esc(3))
                    End Using
                Else
                    Using b As New SolidBrush(cor)
                        g.FillRectangle(b, x, base - h, larguraBarra, h)
                    End Using
                End If
                If dia = 1 OrElse dia = dias OrElse dia Mod 5 = 0 OrElse d = r.Hoje Then
                    TextRenderer.DrawText(g, dia.ToString(), fRot, New Rectangle(x - Esc(6), base + Esc(1), larguraBarra + Esc(12), alturaRotulo),
                                          If(d = r.Hoje, TemaMecano.VermelhoPrimario, TemaMecano.CinzaClaro), TextFormatFlags.HorizontalCenter Or TextFormatFlags.NoPadding)
                End If
            Next
        End Using
    End Sub
End Class

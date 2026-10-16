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
'    Analítica          (hoje frmTimeConsulta)              linha a linha
'    Sintética          (hoje frmTimeConsultaAcumulado)     por empresa, colaborador e mês
'    Por funcionário    (hoje frmTimeFuncionario, "Consulta Geral")  horas diretas e indiretas
'    Por usuário e data (hoje frmTimeConsultausuario)       total de cada pessoa no intervalo
'
'  Cada aba usa exatamente as procedures e a mesma escolha entre elas que a tela
'  antiga correspondente (ModApontamento.ConsultaAnalitica, ConsultaSintetica,
'  ConsultaPorFuncionario, ConsultaPorUsuarioData), então os números batem.
'  Só leitura; cada aba exporta CSV. Só gestor e diretoria (ModApontamento.PodeGerenciar).
'
'  Destino no projeto: Grupo_Mecano\Cadastro\frmGerenciamentoTime.vb
' ============================================================================
Public Class frmGerenciamentoTime
    Inherits Form

    Private Shared _instancia As frmGerenciamentoTime = Nothing

    ''' <summary>Abre a tela (ou traz para a frente a que já está aberta), na aba pedida: os quatro itens do menu Gerenciamento chamam esta função.</summary>
    Public Shared Sub Abrir(Optional aba As Integer = ABA_ANALITICA)
        If _instancia Is Nothing OrElse _instancia.IsDisposed Then
            _instancia = New frmGerenciamentoTime()
            _instancia.Show()
        Else
            If _instancia.WindowState = FormWindowState.Minimized Then _instancia.WindowState = FormWindowState.Normal
            _instancia.BringToFront()
            _instancia.Activate()
        End If
        _instancia.MostrarAba(aba)
    End Sub

    Public Const ABA_ANALITICA As Integer = 0
    Public Const ABA_SINTETICA As Integer = 1
    Public Const ABA_FUNCIONARIO As Integer = 2
    Public Const ABA_USUARIO_DATA As Integer = 3

    Private _escala As Single = 1.0F

    Private Function Esc(v As Integer) As Integer
        Return CInt(Math.Round(v * _escala))
    End Function

    ' ---------- dados ----------
    Private _clientes As New List(Of ClienteResumo)()
    Private _colaboradores As New List(Of ColaboradorResumo)()
    Private _linhasAnalitica As New List(Of LinhaAnalitica)()
    Private _linhasSintetica As New List(Of LinhaSintetica)()
    Private _linhasFuncionario As New List(Of LinhaFuncionario)()
    Private _linhasUsuarioData As New List(Of LinhaUsuarioData)()
    Private _filtroAnalitica As FiltroAnalitica = Nothing
    Private _filtroSintetica As FiltroSintetica = Nothing
    Private _periodoFuncionario As Date() = Nothing
    Private _periodoUsuarioData As Date() = Nothing
    Private ReadOnly _consultada(3) As Boolean

    ' ---------- controles comuns ----------
    Private pnlHeader As Panel
    Private pnlAbas As Panel
    Private pnlFiltros As Panel
    Private btnConsultar As BotaoMecano
    Private btnExportar As BotaoMecano
    Private grade As DataGridView
    Private lblVazio As Label
    Private barraStatus As StatusStrip
    Private lblStatus As ToolStripStatusLabel
    Private lblDica As ToolStripStatusLabel

    ' ---------- Analítica ----------
    Private pnlAnalitica As Panel
    Private anDe As DateTimePicker
    Private anAte As DateTimePicker
    Private anColab As ComboBox
    Private anEmpresa As ComboBox
    Private anGeral As RadioButton
    Private anCom As RadioButton
    Private anSem As RadioButton

    ' ---------- Sintética ----------
    Private pnlSintetica As Panel
    Private siDe As DateTimePicker
    Private siAte As DateTimePicker
    Private siColab As ComboBox
    Private siEmpresa As ComboBox
    Private siGrupo As ComboBox
    Private siGeral As RadioButton
    Private siCom As RadioButton
    Private siSem As RadioButton
    Private siAcumular As CheckBox

    ' ---------- Por funcionário ----------
    Private pnlFuncionario As Panel
    Private fuDe As DateTimePicker
    Private fuAte As DateTimePicker
    Private fuColab As ComboBox

    ' ---------- Por usuário e data ----------
    Private pnlUsuarioData As Panel
    Private udDe As DateTimePicker
    Private udAte As DateTimePicker
    Private udEsperado As Label

    Private ReadOnly _abas As String() = {"Analítica", "Sintética", "Por funcionário", "Por usuário e data"}
    Private ReadOnly _abasRects As New List(Of Rectangle)()
    Private _abaAtiva As Integer = ABA_ANALITICA

    Private Const TODOS As String = "Todos os colaboradores"
    Private Const TODAS As String = "Todas as empresas"
    Private Const TODAS_ATIVIDADES As String = "Todas as atividades"
    Private Const TODAS_E_ATIVIDADES As String = "Todas as empresas e atividades"
    Private Const TODOS_GRUPOS As String = "Todos os grupos"

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
    '  ABAS (uma por consulta)
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
            Using f = TemaMecano.FonteCorpo(9.75F, If(ativa, FontStyle.Bold, FontStyle.Regular))
                Dim w = TextRenderer.MeasureText(g, _abas(i), f).Width + Esc(24)
                Dim r As New Rectangle(x, 0, w, pnlAbas.Height - 1)
                _abasRects.Add(r)
                Dim cor = If(ativa, TemaMecano.VermelhoPrimario, TemaMecano.CinzaEscuro)
                TextRenderer.DrawText(g, _abas(i), f, r, cor, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPrefix)
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
        If i < 0 OrElse i > 3 Then Return
        _abaAtiva = i
        pnlAnalitica.Visible = (i = ABA_ANALITICA)
        pnlSintetica.Visible = (i = ABA_SINTETICA)
        pnlFuncionario.Visible = (i = ABA_FUNCIONARIO)
        pnlUsuarioData.Visible = (i = ABA_USUARIO_DATA)
        pnlAbas.Invalidate()
        Preencher()
    End Sub

    ''' <summary>Troca de aba por código (captura da demonstração e testes).</summary>
    Public Sub MostrarAba(i As Integer)
        EscolherAba(i)
    End Sub

    Public ReadOnly Property AbaAtiva As Integer
        Get
            Return _abaAtiva
        End Get
    End Property

    ' =================================================================
    '  FILTROS: um painel por aba dentro do mesmo espaço; botões em comum
    ' =================================================================
    Private Sub MontarFiltros()
        pnlFiltros = New Panel With {.Dock = DockStyle.Top, .Height = Esc(118), .BackColor = TemaMecano.FundoTela}

        btnConsultar = BotaoMecano.Primario("Consultar", Esc(130), Esc(36))
        btnConsultar.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnConsultar.Top = Esc(76)
        AddHandler btnConsultar.Click, Sub() Consultar()
        btnExportar = BotaoMecano.Secundario("Exportar", Esc(110), Esc(36))
        btnExportar.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnExportar.Top = Esc(76)
        btnExportar.Enabled = False
        AddHandler btnExportar.Click, Sub() Exportar()

        pnlAnalitica = MontarFiltrosAnalitica()
        pnlSintetica = MontarFiltrosSintetica()
        pnlFuncionario = MontarFiltrosFuncionario()
        pnlUsuarioData = MontarFiltrosUsuarioData()
        pnlSintetica.Visible = False
        pnlFuncionario.Visible = False
        pnlUsuarioData.Visible = False

        pnlFiltros.Controls.AddRange(New Control() {pnlAnalitica, pnlSintetica, pnlFuncionario, pnlUsuarioData, btnConsultar, btnExportar})
        btnConsultar.BringToFront()
        btnExportar.BringToFront()
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

    Private Function NovoPainelFiltro() As Panel
        Return New Panel With {.Dock = DockStyle.Fill, .BackColor = TemaMecano.FundoTela}
    End Function

    Private Function Rotulo(texto As String, x As Integer, y As Integer) As Label
        Return New Label With {.Text = texto, .AutoSize = True, .Left = x, .Top = y, .ForeColor = TemaMecano.CinzaEscuro, .Font = TemaMecano.FonteCorpo(8.75F)}
    End Function

    Private Function NovoCombo(x As Integer, y As Integer, largura As Integer) As ComboBox
        Return New ComboBox With {.Left = x, .Top = y, .Width = largura, .DropDownStyle = ComboBoxStyle.DropDownList,
                                  .Font = TemaMecano.FonteCorpo(9.75F), .FlatStyle = FlatStyle.System, .DropDownHeight = Esc(340)}
    End Function

    Private Function NovaOpcao(texto As String, x As Integer, y As Integer, marcada As Boolean) As RadioButton
        Return New RadioButton With {.Text = texto, .AutoSize = True, .Left = x, .Top = y, .Checked = marcada, .Font = TemaMecano.FonteCorpo(9.5F)}
    End Function

    ''' <summary>Período "de … a …" num painel; devolve o x livre à direita.</summary>
    Private Function MontarPeriodo(pnl As Panel, x As Integer, y As Integer, ByRef de As DateTimePicker, ByRef ate As DateTimePicker) As Integer
        pnl.Controls.Add(Rotulo("Período", x, y))
        de = New DateTimePicker With {.Left = x, .Top = y + Esc(20), .Width = Esc(118), .Format = DateTimePickerFormat.Short, .Font = TemaMecano.FonteCorpo(9.75F),
                                      .Value = New Date(Date.Today.Year, Date.Today.Month, 1)}
        Dim lblA = New Label With {.Text = "a", .AutoSize = True, .Left = de.Right + Esc(6), .Top = y + Esc(24), .ForeColor = TemaMecano.CinzaEscuro}
        ate = New DateTimePicker With {.Left = lblA.Left + Esc(16), .Top = y + Esc(20), .Width = Esc(118), .Format = DateTimePickerFormat.Short, .Font = TemaMecano.FonteCorpo(9.75F), .Value = Date.Today}
        pnl.Controls.AddRange(New Control() {de, lblA, ate})
        Return ate.Right + Esc(20)
    End Function

    ' ---------- Analítica (a tela de hoje: frmTimeConsulta) ----------
    Private Function MontarFiltrosAnalitica() As Panel
        Dim pnl = NovoPainelFiltro()
        Dim y1 = Esc(14), y2 = Esc(66)
        Dim x = MontarPeriodo(pnl, Esc(24), y1, anDe, anAte)

        pnl.Controls.Add(Rotulo("Colaborador", x, y1))
        anColab = NovoCombo(x, y1 + Esc(20), Esc(280))
        x = anColab.Right + Esc(20)

        pnl.Controls.Add(Rotulo("Empresa", x, y1))
        anEmpresa = NovoCombo(x, y1 + Esc(20), Esc(380))

        pnl.Controls.Add(Rotulo("Mostrar", Esc(24), y2))
        anGeral = NovaOpcao("Tudo", Esc(24), y2 + Esc(20), True)
        anCom = NovaOpcao("Só clientes (com código de empresa)", Esc(90), y2 + Esc(20), False)
        anSem = NovaOpcao("Só atividades internas (sem código)", Esc(360), y2 + Esc(20), False)
        pnl.Controls.AddRange(New Control() {anColab, anEmpresa, anGeral, anCom, anSem})
        Return pnl
    End Function

    ' ---------- Sintética (frmTimeConsultaAcumulado) ----------
    Private Function MontarFiltrosSintetica() As Panel
        Dim pnl = NovoPainelFiltro()
        Dim y1 = Esc(14), y2 = Esc(66)
        Dim x = MontarPeriodo(pnl, Esc(24), y1, siDe, siAte)

        pnl.Controls.Add(Rotulo("Colaborador", x, y1))
        siColab = NovoCombo(x, y1 + Esc(20), Esc(250))
        x = siColab.Right + Esc(20)

        pnl.Controls.Add(Rotulo("Empresa ou atividade", x, y1))
        siEmpresa = NovoCombo(x, y1 + Esc(20), Esc(330))
        x = siEmpresa.Right + Esc(20)

        pnl.Controls.Add(Rotulo("Grupo de empresas", x, y1))
        siGrupo = NovoCombo(x, y1 + Esc(20), Esc(200))

        pnl.Controls.Add(Rotulo("Mostrar", Esc(24), y2))
        siGeral = NovaOpcao("Tudo", Esc(24), y2 + Esc(20), True)
        siCom = NovaOpcao("Só clientes", Esc(90), y2 + Esc(20), False)
        siSem = NovaOpcao("Só atividades internas", Esc(200), y2 + Esc(20), False)
        siAcumular = New CheckBox With {.Text = "Somar todos os colaboradores por empresa", .AutoSize = True, .Left = Esc(400), .Top = y2 + Esc(20), .Checked = True, .Font = TemaMecano.FonteCorpo(9.5F)}
        AddHandler siGeral.CheckedChanged, Sub() If siGeral.Checked Then PreencherEmpresasSintetica()
        AddHandler siCom.CheckedChanged, Sub() If siCom.Checked Then PreencherEmpresasSintetica()
        AddHandler siSem.CheckedChanged, Sub() If siSem.Checked Then PreencherEmpresasSintetica()
        pnl.Controls.AddRange(New Control() {siColab, siEmpresa, siGrupo, siGeral, siCom, siSem, siAcumular})
        Return pnl
    End Function

    ' ---------- Por funcionário (frmTimeFuncionario, "Consulta Geral") ----------
    Private Function MontarFiltrosFuncionario() As Panel
        Dim pnl = NovoPainelFiltro()
        Dim y1 = Esc(14), y2 = Esc(66)
        Dim x = MontarPeriodo(pnl, Esc(24), y1, fuDe, fuAte)

        pnl.Controls.Add(Rotulo("Colaborador", x, y1))
        fuColab = NovoCombo(x, y1 + Esc(20), Esc(280))

        Dim lblAjuda = New Label With {.AutoSize = True, .Left = Esc(24), .Top = y2 + Esc(22), .ForeColor = TemaMecano.CinzaClaro, .Font = TemaMecano.FonteCorpo(9.0F),
                                       .Text = "Horas diretas = em clientes (com código de empresa)  ·  indiretas = atividades internas.  Duplo clique numa pessoa abre a Sintética dela."}
        pnl.Controls.AddRange(New Control() {fuColab, lblAjuda})
        Return pnl
    End Function

    ' ---------- Por usuário e data (frmTimeConsultausuario) ----------
    Private Function MontarFiltrosUsuarioData() As Panel
        Dim pnl = NovoPainelFiltro()
        Dim y1 = Esc(14), y2 = Esc(66)
        MontarPeriodo(pnl, Esc(24), y1, udDe, udAte)
        udEsperado = New Label With {.AutoSize = True, .Left = Esc(24), .Top = y2 + Esc(22), .ForeColor = TemaMecano.CinzaClaro, .Font = TemaMecano.FonteCorpo(9.0F)}
        AddHandler udDe.ValueChanged, Sub() AtualizarEsperado()
        AddHandler udAte.ValueChanged, Sub() AtualizarEsperado()
        pnl.Controls.Add(udEsperado)
        AtualizarEsperado()
        Return pnl
    End Function

    Private Sub AtualizarEsperado()
        If udAte.Value.Date < udDe.Value.Date Then
            udEsperado.Text = "A data final é anterior à inicial."
            Return
        End If
        udEsperado.Text = "Esperado no período: " & Apontamento.FormatarHoras(LinhaUsuarioData.HorasEsperadas(udDe.Value.Date, udAte.Value.Date)) &
                          "  (8h00 por dia útil, 4h00 no sábado, nada no domingo).  Em vermelho quem ficou abaixo ou passou disso."
    End Sub

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
        AddHandler grade.CellFormatting, AddressOf FormatarCelula
        AddHandler grade.CellDoubleClick, Sub(s, e) If e.RowIndex >= 0 Then AbrirDetalheDaLinha(e.RowIndex)
        AddHandler grade.KeyDown, Sub(s, e)
                                      If e.KeyCode = Keys.Enter AndAlso grade.CurrentRow IsNot Nothing Then
                                          e.Handled = True
                                          AbrirDetalheDaLinha(grade.CurrentRow.Index)
                                      End If
                                  End Sub

        lblVazio = New Label With {
            .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleCenter, .Visible = True,
            .ForeColor = TemaMecano.CinzaClaro, .Font = TemaMecano.FonteCorpo(10.0F), .BackColor = TemaMecano.Branco,
            .Text = "Escolha o período e os filtros e clique em Consultar."}
    End Sub

    ''' <summary>Por usuário e data: quem está fora do esperado fica em vermelho, como na tela antiga.</summary>
    Private Sub FormatarCelula(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If _abaAtiva <> ABA_USUARIO_DATA OrElse e.RowIndex < 0 Then Return
        If grade.Columns.Contains("Situação") Then
            Dim situacao = Convert.ToString(grade.Rows(e.RowIndex).Cells("Situação").Value)
            If situacao <> "Em dia" Then
                e.CellStyle.ForeColor = TemaMecano.VermelhoPrimario
                If grade.Columns(e.ColumnIndex).Name = "Colaborador" OrElse grade.Columns(e.ColumnIndex).Name = "Horas" Then
                    e.CellStyle.Font = TemaMecano.FonteCorpo(9.5F, FontStyle.Bold)
                End If
            End If
        End If
    End Sub

    ''' <summary>Por funcionário: duplo clique abre a Sintética daquela pessoa no mesmo período (o que o frmTimeFuncionario faz hoje em outra janela).</summary>
    Private Sub AbrirDetalheDaLinha(indice As Integer)
        If _abaAtiva <> ABA_FUNCIONARIO OrElse indice < 0 OrElse indice >= _linhasFuncionario.Count Then Return
        If Not grade.Columns.Contains("Colaborador") Then Return
        Dim nome = Convert.ToString(grade.Rows(indice).Cells("Colaborador").Value)
        Dim l = _linhasFuncionario.FirstOrDefault(Function(x) x.Usuario = nome)
        If l Is Nothing Then Return
        siDe.Value = fuDe.Value
        siAte.Value = fuAte.Value
        siGeral.Checked = True
        siAcumular.Checked = False
        siGrupo.SelectedIndex = 0
        If siEmpresa.Items.Count > 0 Then siEmpresa.SelectedIndex = 0
        SelecionarColaborador(siColab, l.UsuarioId)
        EscolherAba(ABA_SINTETICA)
        Consultar()
    End Sub

    Private Sub SelecionarColaborador(cbo As ComboBox, usuarioId As Integer)
        For i = 0 To cbo.Items.Count - 1
            Dim c = TryCast(cbo.Items(i), ColaboradorResumo)
            If c IsNot Nothing AndAlso c.Id = usuarioId Then
                cbo.SelectedIndex = i
                Return
            End If
        Next
        cbo.SelectedIndex = 0
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

    ''' <summary>Texto da barra de status (para os testes).</summary>
    Public ReadOnly Property TextoStatus As String
        Get
            Return lblStatus.Text
        End Get
    End Property

    ''' <summary>Linhas que a grade mostra agora (para os testes).</summary>
    Public ReadOnly Property QuantidadeLinhas As Integer
        Get
            Return grade.RowCount
        End Get
    End Property

    ' =================================================================
    '  LISTAS (colaboradores, empresas, atividades, grupos)
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

        PreencherColaboradores(anColab)
        PreencherColaboradores(siColab)
        PreencherColaboradores(fuColab)

        anEmpresa.BeginUpdate()
        anEmpresa.Items.Clear()
        anEmpresa.Items.Add(TODAS)
        For Each c In _clientes
            anEmpresa.Items.Add(c)
        Next
        anEmpresa.EndUpdate()
        anEmpresa.SelectedIndex = 0

        siGrupo.BeginUpdate()
        siGrupo.Items.Clear()
        siGrupo.Items.Add(TODOS_GRUPOS)
        For Each grupo In _clientes.Select(Function(c) c.Grupo.Trim()).Where(Function(t) t <> "").Distinct().OrderBy(Function(t) t)
            siGrupo.Items.Add(grupo)
        Next
        siGrupo.EndUpdate()
        siGrupo.SelectedIndex = 0

        PreencherEmpresasSintetica()
    End Sub

    Private Sub PreencherColaboradores(cbo As ComboBox)
        cbo.BeginUpdate()
        cbo.Items.Clear()
        cbo.Items.Add(TODOS)
        For Each c In _colaboradores
            cbo.Items.Add(c)
        Next
        cbo.EndUpdate()
        cbo.SelectedIndex = 0
    End Sub

    ''' <summary>
    ''' A lista "Empresa ou atividade" da Sintética acompanha o modo, como na tela antiga:
    ''' Tudo = clientes e atividades; Só clientes = clientes; Só atividades = o Outros.txt.
    ''' </summary>
    Private Sub PreencherEmpresasSintetica()
        If siEmpresa Is Nothing Then Return
        Dim atividades As List(Of String)
        Try
            atividades = If(siCom.Checked, New List(Of String)(), ModApontamento.AtividadesInternas())
        Catch
            atividades = New List(Of String)()
        End Try
        siEmpresa.BeginUpdate()
        siEmpresa.Items.Clear()
        siEmpresa.Items.Add(If(siSem.Checked, TODAS_ATIVIDADES, If(siCom.Checked, TODAS, TODAS_E_ATIVIDADES)))
        If Not siSem.Checked Then
            For Each c In _clientes
                siEmpresa.Items.Add(c)
            Next
        End If
        For Each a In atividades
            siEmpresa.Items.Add(a)
        Next
        siEmpresa.EndUpdate()
        siEmpresa.SelectedIndex = 0
    End Sub

    ' =================================================================
    '  CONSULTAR (a aba ativa)
    ' =================================================================
    ''' <summary>Consulta a aba ativa (captura da demonstração e testes).</summary>
    Public Sub ConsultarAgora()
        Consultar()
    End Sub

    ''' <summary>Período das quatro abas de uma vez (captura e testes).</summary>
    Public Sub DefinirPeriodo(de As Date, ate As Date)
        For Each dtp In New DateTimePicker() {anDe, siDe, fuDe, udDe}
            dtp.Value = de.Date
        Next
        For Each dtp In New DateTimePicker() {anAte, siAte, fuAte, udAte}
            dtp.Value = ate.Date
        Next
    End Sub

    Private Function PeriodoValido(de As DateTimePicker, ate As DateTimePicker) As Boolean
        If ate.Value.Date < de.Value.Date Then
            MessageBox.Show(Me, "A data final é anterior à inicial.", "Gerenciamento do Time", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ate.Focus()
            Return False
        End If
        Return True
    End Function

    Private Sub Consultar()
        Me.Cursor = Cursors.WaitCursor
        Try
            Select Case _abaAtiva
                Case ABA_ANALITICA
                    If Not PeriodoValido(anDe, anAte) Then Return
                    Dim f = FiltroAnaliticaDaTela()
                    _linhasAnalitica = ModApontamento.ConsultaAnalitica(f)
                    _filtroAnalitica = f
                Case ABA_SINTETICA
                    If Not PeriodoValido(siDe, siAte) Then Return
                    Dim f = FiltroSinteticaDaTela()
                    _linhasSintetica = ModApontamento.ConsultaSintetica(f)
                    _filtroSintetica = f
                Case ABA_FUNCIONARIO
                    If Not PeriodoValido(fuDe, fuAte) Then Return
                    Dim lista = ModApontamento.ConsultaPorFuncionario(fuDe.Value.Date, fuAte.Value.Date)
                    Dim colab = TryCast(fuColab.SelectedItem, ColaboradorResumo)
                    If colab IsNot Nothing Then lista = lista.Where(Function(l) l.UsuarioId = colab.Id).ToList()
                    _linhasFuncionario = lista
                    _periodoFuncionario = {fuDe.Value.Date, fuAte.Value.Date}
                Case ABA_USUARIO_DATA
                    If Not PeriodoValido(udDe, udAte) Then Return
                    _linhasUsuarioData = ModApontamento.ConsultaPorUsuarioData(udDe.Value.Date, udAte.Value.Date)
                    _periodoUsuarioData = {udDe.Value.Date, udAte.Value.Date}
            End Select
            _consultada(_abaAtiva) = True
        Catch ex As Exception
            MessageBox.Show(Me, "Não foi possível consultar." & vbCrLf & ex.Message, "Gerenciamento do Time", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
        Preencher()
    End Sub

    Private Function FiltroAnaliticaDaTela() As FiltroAnalitica
        Dim f As New FiltroAnalitica With {.De = anDe.Value.Date, .Ate = anAte.Value.Date}
        Dim colab = TryCast(anColab.SelectedItem, ColaboradorResumo)
        f.Usuario = If(colab Is Nothing, FiltroAnalitica.TODOS_USUARIOS, colab.Nome)
        Dim emp = TryCast(anEmpresa.SelectedItem, ClienteResumo)
        f.Empresa = If(emp Is Nothing, FiltroAnalitica.TODAS_EMPRESAS, emp.Nome)
        f.EmpresaId = If(emp Is Nothing, "", emp.Tradutor)
        f.Modo = If(anCom.Checked, ModoAnalitica.ComCodigoEmpresa, If(anSem.Checked, ModoAnalitica.SemCodigoEmpresa, ModoAnalitica.Geral))
        Return f
    End Function

    Private Function FiltroSinteticaDaTela() As FiltroSintetica
        Dim f As New FiltroSintetica With {.De = siDe.Value.Date, .Ate = siAte.Value.Date, .AcumularPorEmpresa = siAcumular.Checked}
        Dim colab = TryCast(siColab.SelectedItem, ColaboradorResumo)
        f.Usuario = If(colab Is Nothing, FiltroAnalitica.TODOS_USUARIOS, colab.Nome)
        Dim emp = TryCast(siEmpresa.SelectedItem, ClienteResumo)
        If emp IsNot Nothing Then
            f.Empresa = emp.Nome
            f.EmpresaId = emp.Tradutor
        ElseIf siEmpresa.SelectedIndex > 0 AndAlso TypeOf siEmpresa.SelectedItem Is String Then
            f.Empresa = CStr(siEmpresa.SelectedItem)     ' atividade interna
            f.EmpresaId = ""
        Else
            f.Empresa = FiltroAnalitica.TODAS_EMPRESAS
            f.EmpresaId = ""
        End If
        f.Grupo = If(siGrupo.SelectedIndex <= 0, FiltroSintetica.TODOS_GRUPOS, CStr(siGrupo.SelectedItem))
        f.Modo = If(siCom.Checked, ModoAnalitica.ComCodigoEmpresa, If(siSem.Checked, ModoAnalitica.SemCodigoEmpresa, ModoAnalitica.Geral))
        Return f
    End Function

    ' =================================================================
    '  PREENCHER A GRADE (conforme a aba)
    ' =================================================================
    Private Sub Preencher()
        If Not _consultada(_abaAtiva) Then
            grade.DataSource = Nothing
            lblVazio.Text = "Escolha o período e os filtros e clique em Consultar."
            lblVazio.Visible = True
            btnExportar.Enabled = False
            MostrarMensagem("")
            Return
        End If
        Select Case _abaAtiva
            Case ABA_ANALITICA : PreencherAnalitica()
            Case ABA_SINTETICA : PreencherSintetica()
            Case ABA_FUNCIONARIO : PreencherFuncionario()
            Case ABA_USUARIO_DATA : PreencherUsuarioData()
        End Select
    End Sub

    Private Sub Coluna(nome As String, peso As Single, Optional direita As Boolean = False, Optional negrito As Boolean = False, Optional apagado As Boolean = False)
        If Not grade.Columns.Contains(nome) Then Return
        Dim c = grade.Columns(nome)
        c.FillWeight = peso
        If direita Then
            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            c.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        End If
        If negrito Then c.DefaultCellStyle.Font = TemaMecano.FonteCorpo(9.5F, FontStyle.Bold)
        If apagado Then c.DefaultCellStyle.ForeColor = TemaMecano.CinzaClaro
    End Sub

    Private Sub AplicarTabela(dt As DataTable, vazio As String)
        grade.SuspendLayout()
        grade.DataSource = dt
        grade.ResumeLayout()
        grade.ClearSelection()
        lblVazio.Text = vazio
        lblVazio.Visible = (dt.Rows.Count = 0)
        btnExportar.Enabled = (dt.Rows.Count > 0)
    End Sub

    Private Function TextoPeriodo(de As Date, ate As Date) As String
        Return de.ToString("dd/MM") & " a " & ate.ToString("dd/MM/yyyy")
    End Function

    Private Sub PreencherAnalitica()
        Dim dt As New DataTable()
        dt.Columns.Add("Data", GetType(Date))
        dt.Columns.Add("Dia", GetType(String))
        dt.Columns.Add("Colaborador", GetType(String))
        dt.Columns.Add("Depto", GetType(String))
        dt.Columns.Add("Empresa", GetType(String))
        dt.Columns.Add("Código", GetType(String))
        dt.Columns.Add("Horas", GetType(String))
        For Each l In _linhasAnalitica
            dt.Rows.Add(l.Data, Apontamento.PtBr.DateTimeFormat.GetAbbreviatedDayName(l.Data.DayOfWeek).TrimEnd("."c), l.Usuario, l.Depto, l.Empresa, If(l.EhCliente, l.EmpresaId, ""), l.HorasTexto)
        Next
        AplicarTabela(dt, "Nenhum apontamento com esses filtros.")
        If grade.Columns.Count > 0 Then
            Coluna("Data", 9) : grade.Columns("Data").DefaultCellStyle.Format = "dd/MM/yyyy"
            Coluna("Dia", 5, apagado:=True)
            Coluna("Colaborador", 20)
            Coluna("Depto", 10)
            Coluna("Empresa", 36)
            Coluna("Código", 7)
            Coluna("Horas", 7, direita:=True, negrito:=True)
        End If
        Dim f = _filtroAnalitica
        Dim quem = If(f Is Nothing OrElse f.TodosUsuarios, "toda a equipe", f.Usuario)
        MostrarMensagem(_linhasAnalitica.Count & If(_linhasAnalitica.Count = 1, " apontamento", " apontamentos") & "  ·  " & Apontamento.FormatarHoras(_linhasAnalitica.Sum(Function(l) l.Horas)) & " no total  ·  " &
                        quem & If(f Is Nothing, "", "  ·  " & TextoPeriodo(f.De, f.Ate)))
    End Sub

    ''' <summary>Linhas da Sintética como a tela mostra: acumuladas por empresa quando a opção está marcada.</summary>
    Private Function LinhasSinteticaExibidas() As List(Of LinhaSintetica)
        If _filtroSintetica IsNot Nothing AndAlso _filtroSintetica.AcumularPorEmpresa Then
            Return LinhaSintetica.AcumularPorEmpresa(_linhasSintetica)
        End If
        Return _linhasSintetica
    End Function

    Private Sub PreencherSintetica()
        Dim linhas = LinhasSinteticaExibidas()
        Dim acumulada = (_filtroSintetica IsNot Nothing AndAlso _filtroSintetica.AcumularPorEmpresa)
        Dim dt As New DataTable()
        dt.Columns.Add("Código", GetType(String))
        dt.Columns.Add("Empresa ou atividade", GetType(String))
        dt.Columns.Add("CNPJ", GetType(String))
        dt.Columns.Add("Grupo", GetType(String))
        If Not acumulada Then dt.Columns.Add("Colaborador", GetType(String))
        dt.Columns.Add("Mês", GetType(String))
        dt.Columns.Add("Horas", GetType(String))
        For Each l In linhas
            Dim cnpj = (New ClienteResumo With {.Cnpj = l.Cnpj}).CnpjFormatado
            If acumulada Then
                dt.Rows.Add(l.Codigo, l.Empresa, cnpj, l.Grupo, l.MesTexto, l.HorasTexto)
            Else
                dt.Rows.Add(l.Codigo, l.Empresa, cnpj, l.Grupo, l.Usuario, l.MesTexto, l.HorasTexto)
            End If
        Next
        AplicarTabela(dt, "Nenhum apontamento com esses filtros.")
        If grade.Columns.Count > 0 Then
            Coluna("Código", 7)
            Coluna("Empresa ou atividade", If(acumulada, 44, 32))
            Coluna("CNPJ", 13, apagado:=True)
            Coluna("Grupo", 14)
            Coluna("Colaborador", 18)
            Coluna("Mês", 8)
            Coluna("Horas", 8, direita:=True, negrito:=True)
        End If
        Dim f = _filtroSintetica
        Dim quem = If(f Is Nothing OrElse f.TodosUsuarios, "toda a equipe", f.Usuario)
        Dim empresas = linhas.Select(Function(l) l.ChaveEmpresa).Distinct().Count()
        MostrarMensagem(linhas.Count & If(linhas.Count = 1, " linha", " linhas") & "  ·  " & empresas & If(empresas = 1, " empresa", " empresas") & "  ·  " &
                        Apontamento.FormatarHoras(_linhasSintetica.Sum(Function(l) l.Horas)) & " no total  ·  " & quem &
                        If(f Is Nothing, "", "  ·  " & TextoPeriodo(f.De, f.Ate)))
    End Sub

    Private Sub PreencherFuncionario()
        Dim dt As New DataTable()
        dt.Columns.Add("Colaborador", GetType(String))
        dt.Columns.Add("Depto", GetType(String))
        dt.Columns.Add("Função", GetType(String))
        dt.Columns.Add("Diretas", GetType(String))
        dt.Columns.Add("Indiretas", GetType(String))
        dt.Columns.Add("Total", GetType(String))
        dt.Columns.Add("% diretas", GetType(String))
        dt.Columns.Add("% indiretas", GetType(String))
        For Each l In _linhasFuncionario
            dt.Rows.Add(l.Usuario, l.Depto, l.Funcao, Apontamento.FormatarHoras(l.HorasDiretas), Apontamento.FormatarHoras(l.HorasIndiretas),
                        Apontamento.FormatarHoras(l.Total), l.PercDiretas & "%", l.PercIndiretas & "%")
        Next
        AplicarTabela(dt, "Ninguém tem apontamento nesse período.")
        If grade.Columns.Count > 0 Then
            Coluna("Colaborador", 24)
            Coluna("Depto", 11)
            Coluna("Função", 19)
            Coluna("Diretas", 9, direita:=True)
            Coluna("Indiretas", 9, direita:=True)
            Coluna("Total", 9, direita:=True, negrito:=True)
            Coluna("% diretas", 9, direita:=True, apagado:=True)
            Coluna("% indiretas", 10, direita:=True, apagado:=True)
        End If
        Dim diretas = _linhasFuncionario.Sum(Function(l) l.HorasDiretas)
        Dim indiretas = _linhasFuncionario.Sum(Function(l) l.HorasIndiretas)
        Dim p = _periodoFuncionario
        MostrarMensagem(_linhasFuncionario.Count & If(_linhasFuncionario.Count = 1, " pessoa", " pessoas") & "  ·  diretas " & Apontamento.FormatarHoras(diretas) &
                        "  ·  indiretas " & Apontamento.FormatarHoras(indiretas) & "  ·  total " & Apontamento.FormatarHoras(diretas + indiretas) &
                        If(p Is Nothing, "", "  ·  " & TextoPeriodo(p(0), p(1))))
    End Sub

    Private Sub PreencherUsuarioData()
        Dim dt As New DataTable()
        dt.Columns.Add("Colaborador", GetType(String))
        dt.Columns.Add("Depto", GetType(String))
        dt.Columns.Add("Horas", GetType(String))
        dt.Columns.Add("Esperado", GetType(String))
        dt.Columns.Add("Diferença", GetType(String))
        dt.Columns.Add("Situação", GetType(String))
        For Each l In _linhasUsuarioData
            dt.Rows.Add(l.Usuario, l.Depto, Apontamento.FormatarHoras(l.Horas), Apontamento.FormatarHoras(l.Esperado),
                        If(l.Diferenca >= 0D, "+", "") & Apontamento.FormatarHoras(l.Diferenca), l.Situacao)
        Next
        AplicarTabela(dt, "Nenhum colaborador para mostrar.")
        If grade.Columns.Count > 0 Then
            Coluna("Colaborador", 30)
            Coluna("Depto", 14)
            Coluna("Horas", 12, direita:=True, negrito:=True)
            Coluna("Esperado", 12, direita:=True)
            Coluna("Diferença", 12, direita:=True)
            Coluna("Situação", 12)
        End If
        Dim abaixo = _linhasUsuarioData.Where(Function(l) l.Horas < l.Esperado).Count()
        Dim acima = _linhasUsuarioData.Where(Function(l) l.Horas > l.Esperado + LinhaUsuarioData.TOLERANCIA).Count()
        Dim p = _periodoUsuarioData
        MostrarMensagem(_linhasUsuarioData.Count & If(_linhasUsuarioData.Count = 1, " pessoa", " pessoas") & "  ·  " & abaixo & " abaixo do esperado  ·  " & acima & " acima  ·  " &
                        Apontamento.FormatarHoras(_linhasUsuarioData.Sum(Function(l) l.Horas)) & " no total" &
                        If(p Is Nothing, "", "  ·  " & TextoPeriodo(p(0), p(1))))
    End Sub

    ' =================================================================
    '  EXPORTAR (CSV da aba ativa, com linha de total)
    ' =================================================================
    Private Sub Exportar()
        Dim sb As New StringBuilder()
        Dim nome As String
        Select Case _abaAtiva
            Case ABA_ANALITICA
                If _linhasAnalitica.Count = 0 Then Return
                nome = "time-analitica-" & NomePeriodo(anDe, anAte)
                sb.AppendLine("Data;Dia;Colaborador;Depto;Empresa;Código;Horas;Horas (decimal)")
                For Each l In _linhasAnalitica.OrderBy(Function(x) x.Data).ThenBy(Function(x) x.Usuario)
                    sb.AppendLine(l.Data.ToString("dd/MM/yyyy") & ";" & Apontamento.PtBr.DateTimeFormat.GetAbbreviatedDayName(l.Data.DayOfWeek).TrimEnd("."c) & ";" &
                                  l.Usuario & ";" & l.Depto & ";" & Csv(l.Empresa) & ";" & If(l.EhCliente, l.EmpresaId, "") & ";" &
                                  l.HorasTexto & ";" & Apontamento.FormatarDecimal(l.Horas))
                Next
                sb.AppendLine(";;;;;;Total;" & Apontamento.FormatarDecimal(_linhasAnalitica.Sum(Function(l) l.Horas)))
            Case ABA_SINTETICA
                Dim linhas = LinhasSinteticaExibidas()
                If linhas.Count = 0 Then Return
                nome = "time-sintetica-" & NomePeriodo(siDe, siAte)
                sb.AppendLine("Código;Empresa ou atividade;CNPJ;Grupo;Colaborador;Ano;Mês;Horas;Horas (decimal)")
                For Each l In linhas
                    sb.AppendLine(l.Codigo & ";" & Csv(l.Empresa) & ";" & l.Cnpj & ";" & Csv(l.Grupo) & ";" & l.Usuario & ";" & l.Ano & ";" & l.Mes & ";" &
                                  l.HorasTexto & ";" & Apontamento.FormatarDecimal(l.Horas))
                Next
                sb.AppendLine(";;;;;;;Total;" & Apontamento.FormatarDecimal(_linhasSintetica.Sum(Function(l) l.Horas)))
            Case ABA_FUNCIONARIO
                If _linhasFuncionario.Count = 0 Then Return
                nome = "time-por-funcionario-" & NomePeriodo(fuDe, fuAte)
                sb.AppendLine("Colaborador;Depto;Função;Diretas;Indiretas;Total;% diretas;% indiretas;Diretas (decimal);Indiretas (decimal);Total (decimal)")
                For Each l In _linhasFuncionario
                    sb.AppendLine(l.Usuario & ";" & l.Depto & ";" & Csv(l.Funcao) & ";" & Apontamento.FormatarHoras(l.HorasDiretas) & ";" & Apontamento.FormatarHoras(l.HorasIndiretas) & ";" &
                                  Apontamento.FormatarHoras(l.Total) & ";" & l.PercDiretas & ";" & l.PercIndiretas & ";" &
                                  Apontamento.FormatarDecimal(l.HorasDiretas) & ";" & Apontamento.FormatarDecimal(l.HorasIndiretas) & ";" & Apontamento.FormatarDecimal(l.Total))
                Next
                sb.AppendLine("Total;;;" & Apontamento.FormatarHoras(_linhasFuncionario.Sum(Function(l) l.HorasDiretas)) & ";" & Apontamento.FormatarHoras(_linhasFuncionario.Sum(Function(l) l.HorasIndiretas)) & ";" &
                              Apontamento.FormatarHoras(_linhasFuncionario.Sum(Function(l) l.Total)) & ";;;" &
                              Apontamento.FormatarDecimal(_linhasFuncionario.Sum(Function(l) l.HorasDiretas)) & ";" & Apontamento.FormatarDecimal(_linhasFuncionario.Sum(Function(l) l.HorasIndiretas)) & ";" &
                              Apontamento.FormatarDecimal(_linhasFuncionario.Sum(Function(l) l.Total)))
            Case Else
                If _linhasUsuarioData.Count = 0 Then Return
                nome = "time-por-usuario-" & NomePeriodo(udDe, udAte)
                sb.AppendLine("Colaborador;Depto;Horas;Esperado;Diferença;Situação;Horas (decimal)")
                For Each l In _linhasUsuarioData
                    sb.AppendLine(l.Usuario & ";" & l.Depto & ";" & Apontamento.FormatarHoras(l.Horas) & ";" & Apontamento.FormatarHoras(l.Esperado) & ";" &
                                  If(l.Diferenca >= 0D, "+", "") & Apontamento.FormatarHoras(l.Diferenca) & ";" & l.Situacao & ";" & Apontamento.FormatarDecimal(l.Horas))
                Next
                sb.AppendLine("Total;;" & Apontamento.FormatarHoras(_linhasUsuarioData.Sum(Function(l) l.Horas)) & ";;;;" & Apontamento.FormatarDecimal(_linhasUsuarioData.Sum(Function(l) l.Horas)))
        End Select

        Using dlg As New SaveFileDialog With {.Filter = "CSV (*.csv)|*.csv", .FileName = nome & ".csv"}
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Try
                File.WriteAllText(dlg.FileName, sb.ToString(), New UTF8Encoding(True))
                Process.Start(dlg.FileName)
            Catch ex As Exception
                MessageBox.Show(Me, "Não foi possível gravar o arquivo." & vbCrLf & ex.Message, "Gerenciamento do Time", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    Private Function NomePeriodo(de As DateTimePicker, ate As DateTimePicker) As String
        Return de.Value.ToString("yyyy-MM-dd") & "-a-" & ate.Value.ToString("yyyy-MM-dd")
    End Function

    Private Function Csv(s As String) As String
        Return If(s, "").Replace(";", ",")
    End Function

End Class

Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Linq
Imports System.Windows.Forms

' ============================================================================
'  frmApontamentoDet - lançamento de horas (novo / alterar).
'
'  Substitui o "Cadastro de Tempo" (frmTimeDet) com as mesmas regras:
'    - data não pode ser futura; tempo não pode ser zero;
'    - OU um cliente (TbClientes, grava EMPRESAID = TRADUTOR) OU uma atividade
'      interna do Outros.txt (EMPRESAID vazio) - nunca os dois;
'    - "várias empresas": grava uma linha por empresa. Ou todas com o mesmo
'      tempo (2 h em 4 empresas = 8 h no dia, como o Time de hoje), ou o tempo
'      dividido por igual entre elas (2 h em 4 empresas = 0,50 h cada, 2 h no
'      dia). O rateio é em centésimos de hora, para a soma fechar exata;
'    - "administrativo" só aparece para gestor e diretoria;
'    - teto de 10 h por dia (BancoHoras.LIMITE_DIA): o que já está no dia mais o
'      que vai ser gravado não pode passar disso; o que passar de 8 h é crédito
'      no banco de horas (BancoHoras, em Apontamento.vb), e a tela avisa;
'    - depois de gravar, o frmMain reavalia a trava de horas (AposGravar).
'  O que saiu: o campo "Retrabalho" (oculto no Designer desde 2022 e sem
'  nenhum lançamento desde 2025) e a digitação de CNPJ (era só exibição).
'
'  Tempo é digitado em horas + minutos e gravado como o Time sempre gravou:
'  decimal com duas casas (1 h 30 min = 1,50).
'
'  Aberta pela frmApontamentos com ShowDialog(); devolve DialogResult.OK
'  quando gravou algo (IdGravado, QuantidadeGravada, MensagemResultado).
'
'  Destino no projeto: Grupo_Mecano\Cadastro\frmApontamentoDet.vb
' ============================================================================
Public Class frmApontamentoDet
    Inherits Form

    Private _escala As Single = 1.0F

    Private Function Esc(v As Integer) As Integer
        Return CInt(Math.Round(v * _escala))
    End Function

    Private ReadOnly _id As Integer
    Private _original As Apontamento = Nothing
    Private ReadOnly _resumo As ResumoHoras
    Private ReadOnly _dataInicial As Date
    Private _gravou As Boolean = False
    Private _carregando As Boolean = True
    Private _clientes As New List(Of ClienteResumo)()
    Private _atividades As New List(Of String)()
    Private ReadOnly _disponiveis As New List(Of ClienteResumo)()
    Private ReadOnly _selecionados As New List(Of ClienteResumo)()
    Private _jaNoDia As Decimal = 0D                 ' horas que a pessoa já tem na data escolhida (sem este apontamento)
    Private _jaNoDiaData As Date = Date.MinValue     ' data a que _jaNoDia se refere
    Private _acimaDaJornada As Boolean = False       ' o dia, com o que está digitado, passa de 8h00: subtítulo em vermelho
    Private _colaboradores As New List(Of ColaboradorResumo)()   ' só para gestor e diretoria, em lançamento novo

    Public Property IdGravado As Integer
    Public Property QuantidadeGravada As Integer
    Public Property MensagemResultado As String = ""
    '''<summary>Texto do aviso mostrado na tela quando o dia passou de 8h00 (vazio quando não passou).</summary>
    Public Property AvisoJornada As String = ""

    ' ---------- controles ----------
    Private pnlTopo As Panel
    Private pnlCorpo As Panel
    Private pnlRodape As Panel
    Private secOque As Panel
    Private secQuando As Panel
    Private rbCliente As RadioButton
    Private rbInterno As RadioButton
    Private txtCliente As TextBox            ' digita qualquer parte do nome ou o código...
    Private lstSugestoes As ListBox          ' ...e a lista filtra conforme escreve; Tab, Enter ou clique seleciona
    Private _clienteSelecionado As ClienteResumo = Nothing
    Private _sugestoes As New List(Of ClienteResumo)()
    Private Const MAX_SUGESTOES As Integer = 40
    Private lblClienteInfo As Label
    Private cboAtividade As ComboBox
    Private lblAtividadeAviso As Label
    Private chkVarias As CheckBox
    Private pnlVarias As Panel
    Private txtFiltro As TextBox
    Private lstDisponiveis As CheckedListBox
    Private lstSelecionadas As CheckedListBox
    Private chkTodasDisp As CheckBox
    Private chkTodasSel As CheckBox
    Private lblContagemSel As Label
    Private rbMesmoTempo As RadioButton      ' o tempo informado vale para cada empresa (como o Time de hoje)
    Private rbDividir As RadioButton         ' o tempo informado é dividido por igual entre as empresas
    Private cboGrupo As ComboBox             ' "Meus grupos": marca de uma vez as empresas do grupo escolhido
    Private btnGrupos As BotaoMecano         ' abre a janela de montar os grupos
    Private _grupos As New List(Of GrupoEmpresas)()
    Private dtpData As DateTimePicker
    Private numHoras As NumericUpDown
    Private numMinutos As NumericUpDown
    Private lblEquivale As Label
    Private lblColaborador As Label           ' gestor e diretoria: em nome de quem gravar
    Private cboColaborador As ComboBox
    Private btnSalvar As BotaoMecano
    Private btnSalvarNovo As BotaoMecano
    Private btnCancelar As BotaoMecano
    Private lblRodape As Label

    Private Const LARGURA As Integer = 780
    Private Const SEC_X As Integer = 24
    Private Const SEC_W As Integer = 732
    Private Const ALT_TOPO As Integer = 88
    Private Const ALT_OQUE_BASE As Integer = 164
    Private Const ALT_VARIAS As Integer = 388
    Private Const ALT_QUANDO As Integer = 184
    Private Const ALT_RODAPE As Integer = 64
    Private Const EM_SETCUEBANNER As Integer = &H1501

    <Runtime.InteropServices.DllImport("user32.dll", CharSet:=Runtime.InteropServices.CharSet.Unicode)>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As String) As IntPtr
    End Function

    ''' <param name="id">0 = novo; senão o CODIGOID da TbTime a alterar.</param>
    ''' <param name="dataInicial">data sugerida para um lançamento novo (ex.: dia pendente clicado no resumo).</param>
    ''' <param name="resumo">totais já calculados pela lista, para o subtítulo (opcional).</param>
    Public Sub New(id As Integer, Optional dataInicial As Date? = Nothing, Optional resumo As ResumoHoras = Nothing)
        _id = id
        _resumo = resumo
        _dataInicial = If(dataInicial.HasValue, dataInicial.Value.Date, Date.Today)
        If _dataInicial > Date.Today Then _dataInicial = Date.Today
        Using g = Me.CreateGraphics()
            _escala = g.DpiX / 96.0F
        End Using

        Me.AutoScaleMode = AutoScaleMode.None
        Me.Text = If(_id = 0, "Novo apontamento", "Apontamento")
        Me.BackColor = TemaMecano.FundoTela
        Me.Font = TemaMecano.FonteCorpo(9.75F)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.KeyPreview = True
        Me.DoubleBuffered = True

        MontarTopo()
        MontarRodape()
        MontarCorpo()

        Me.Controls.Add(pnlCorpo)
        Me.Controls.Add(pnlTopo)
        Me.Controls.Add(pnlRodape)
        pnlCorpo.BringToFront()
        AjustarAltura()

        AddHandler Me.Load, Sub() CarregarDados()
        AddHandler Me.KeyDown, AddressOf AoTeclar
        AddHandler Me.FormClosing, Sub(s, e)
                                       If Me.DialogResult <> DialogResult.OK AndAlso _gravou Then Me.DialogResult = DialogResult.OK
                                   End Sub
    End Sub

    ' =================================================================
    '  TOPO: título e contexto do dia
    ' =================================================================
    Private Sub MontarTopo()
        pnlTopo = New Panel With {.Dock = DockStyle.Top, .Height = Esc(ALT_TOPO), .BackColor = TemaMecano.Branco}
        AddHandler pnlTopo.Paint, AddressOf DesenharTopo
    End Sub

    Private Sub DesenharTopo(sender As Object, e As PaintEventArgs)
        Dim g = e.Graphics
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
        Using p As New Pen(TemaMecano.BordaCampo)
            g.DrawLine(p, 0, pnlTopo.Height - 1, pnlTopo.Width, pnlTopo.Height - 1)
        End Using
        Dim titulo As String
        If _id = 0 Then
            titulo = "Novo apontamento"
        ElseIf _original IsNot Nothing Then
            titulo = "Apontamento de " & Apontamento.FormatarData(_original.Data)
        Else
            titulo = "Apontamento"
        End If
        Using fTitulo = TemaMecano.FonteTitulo(15.0F)
            TextRenderer.DrawText(g, titulo, fTitulo, New Rectangle(Esc(24), Esc(18), pnlTopo.Width - Esc(48), Esc(30)),
                                  TemaMecano.TextoCampo, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPrefix)
        End Using
        Dim textoSub As String = Subtitulo()
        Using fSub = TemaMecano.FonteCorpo(9.25F, If(_acimaDaJornada, FontStyle.Bold, FontStyle.Regular))
            TextRenderer.DrawText(g, textoSub, fSub, New Rectangle(Esc(24), Esc(50), pnlTopo.Width - Esc(48), Esc(22)),
                                  If(_acimaDaJornada, TemaMecano.VermelhoPrimario, TemaMecano.CinzaEscuro), TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPrefix)
        End Using
    End Sub

    Private Function Subtitulo() As String
        If _id > 0 AndAlso _original IsNot Nothing Then
            Return "cód. " & _original.Id & "  ·  " & DescricaoDoDia()
        End If
        Return DescricaoDoDia()
    End Function

    ''' <summary>
    ''' "Hoje: 7h00 · faltam 1h00 · com este: 9h00 · acima de 8h00" (em vermelho). Só os números da pessoa; quando passa do teto de 10 h, avisa já aqui.
    ''' </summary>
    Private Function DescricaoDoDia() As String
        Dim dia = If(dtpData Is Nothing, _dataInicial, dtpData.Value.Date)
        Dim jornada = BancoHoras.JornadaDoDia(dia)
        Dim s = If(OutraPessoa(), NomeAlvo() & "  ·  ", "") & If(dia = Date.Today, "Hoje", Apontamento.FormatarData(dia)) & ": " & Apontamento.FormatarHoras(_jaNoDia)
        If jornada > 0D AndAlso _jaNoDia < jornada Then
            s &= "  ·  faltam " & Apontamento.FormatarHoras(jornada - _jaNoDia)
        End If
        Dim novo = TotalASalvar()
        _acimaDaJornada = False
        If novo > 0D Then
            Dim total = _jaNoDia + novo
            s &= "  ·  com este: " & Apontamento.FormatarHoras(total)
            If total > BancoHoras.LIMITE_DIA Then
                s &= "  ·  passa de " & Apontamento.FormatarHoras(BancoHoras.LIMITE_DIA)
                _acimaDaJornada = True
            ElseIf BancoHoras.ATIVO Then
                Dim extra = BancoHoras.ExtraDoLancamento(dia, _jaNoDia, novo)
                If extra > 0D Then s &= "  (" & Apontamento.FormatarHoras(extra) & " no banco)"
            ElseIf total > ResumoHoras.JORNADA Then
                s &= "  ·  acima de " & Apontamento.FormatarHoras(ResumoHoras.JORNADA)
                _acimaDaJornada = True
            End If
        End If
        Return s
    End Function

    ''' <summary>Soma do que vai ser gravado: uma linha, N linhas iguais ou o rateio.</summary>
    Private Function TotalASalvar() As Decimal
        If numHoras Is Nothing Then Return 0D
        Dim min = MinutosInformados()
        If min = 0 Then Return 0D
        Dim horas = Apontamento.DeMinutos(min)
        If chkVarias IsNot Nothing AndAlso chkVarias.Checked AndAlso _selecionados.Count > 0 Then Return PartesPorEmpresa(horas).Sum()
        Return horas
    End Function

    '''<summary>Em nome de quem o lançamento é gravado: a pessoa logada ou, para gestor e diretoria, a escolhida na lista.</summary>
    Private Function UsuarioAlvo() As Integer
        Dim c = If(cboColaborador Is Nothing, Nothing, TryCast(cboColaborador.SelectedItem, ColaboradorResumo))
        Return If(c Is Nothing, ModApontamento.UsuarioLogadoId(), c.Id)
    End Function

    Private Function OutraPessoa() As Boolean
        Return UsuarioAlvo() <> ModApontamento.UsuarioLogadoId()
    End Function

    Private Function NomeAlvo() As String
        Dim c = If(cboColaborador Is Nothing, Nothing, TryCast(cboColaborador.SelectedItem, ColaboradorResumo))
        Return If(c Is Nothing, "", c.Nome)
    End Function

    ''' <summary>Lê (uma vez por data) quanto a pessoa já tem no dia escolhido, sem este apontamento quando é alteração.</summary>
    Private Sub CarregarHorasDoDia(Optional forcar As Boolean = False)
        Dim dia = If(dtpData Is Nothing, _dataInicial, dtpData.Value.Date)
        If dia = _jaNoDiaData AndAlso Not forcar Then Return
        Try
            _jaNoDia = ModApontamento.HorasNoDiaDe(UsuarioAlvo(), dia, _id)
        Catch
            _jaNoDia = 0D
        End Try
        _jaNoDiaData = dia
        If pnlTopo IsNot Nothing Then pnlTopo.Invalidate()
    End Sub

    ' =================================================================
    '  RODAPÉ: Salvar / Salvar e novo / Cancelar
    ' =================================================================
    Private Sub MontarRodape()
        pnlRodape = New Panel With {.Dock = DockStyle.Bottom, .Height = Esc(ALT_RODAPE), .BackColor = TemaMecano.Branco}
        AddHandler pnlRodape.Paint, Sub(s, e)
                                        Using p As New Pen(TemaMecano.BordaCampo)
                                            e.Graphics.DrawLine(p, 0, 0, pnlRodape.Width, 0)
                                        End Using
                                    End Sub

        btnSalvar = BotaoMecano.Primario("Salvar", Esc(130), Esc(36))
        btnSalvar.Top = Esc(14)
        btnSalvar.Left = Esc(LARGURA) - Esc(24) - btnSalvar.Width
        AddHandler btnSalvar.Click, Sub() Salvar(False)

        btnSalvarNovo = BotaoMecano.Secundario("Salvar e novo", Esc(140), Esc(36))
        btnSalvarNovo.Top = Esc(14)
        btnSalvarNovo.Left = btnSalvar.Left - Esc(10) - btnSalvarNovo.Width
        btnSalvarNovo.Visible = (_id = 0)
        AddHandler btnSalvarNovo.Click, Sub() Salvar(True)

        btnCancelar = BotaoMecano.Secundario("Cancelar", Esc(110), Esc(36))
        btnCancelar.Top = Esc(14)
        btnCancelar.Left = If(_id = 0, btnSalvarNovo.Left, btnSalvar.Left) - Esc(10) - btnCancelar.Width
        AddHandler btnCancelar.Click, Sub()
                                          Me.DialogResult = If(_gravou, DialogResult.OK, DialogResult.Cancel)
                                          Me.Close()
                                      End Sub

        ' o texto vai até 12 px antes do Cancelar (com três botões sobra pouco: o que não couber ganha "...");
        ' ele fica atrás dos botões na ordem Z para nunca cobrir a borda deles
        lblRodape = New Label With {
            .AutoSize = False, .Left = Esc(24), .Top = Esc(14), .Width = btnCancelar.Left - Esc(24) - Esc(12), .Height = Esc(36), .AutoEllipsis = True,
            .TextAlign = ContentAlignment.MiddleLeft, .ForeColor = TemaMecano.CinzaClaro, .Font = TemaMecano.FonteCorpo(8.75F),
            .Text = If(_id = 0, "Ctrl+S salva  ·  Ctrl+Enter salva e abre outro  ·  Esc fecha", "Ctrl+S salva  ·  Esc fecha")}

        pnlRodape.Controls.AddRange(New Control() {btnSalvar, btnSalvarNovo, btnCancelar, lblRodape})
    End Sub

    ' =================================================================
    '  CORPO: "o que foi feito" e "quando e quanto tempo"
    ' =================================================================
    Private Sub MontarCorpo()
        pnlCorpo = New Panel With {.Dock = DockStyle.Fill, .AutoScroll = True, .BackColor = TemaMecano.FundoTela}

        ' ---------------- o que foi feito
        secOque = NovaSecao("O QUE FOI FEITO", SEC_X, 20, ALT_OQUE_BASE)

        rbCliente = New RadioButton With {.Text = "Cliente", .AutoSize = True, .Left = Esc(16), .Top = Esc(38), .Checked = True, .Font = TemaMecano.FonteCorpo(9.75F)}
        rbInterno = New RadioButton With {.Text = "Atividade interna", .AutoSize = True, .Left = Esc(120), .Top = Esc(38), .Font = TemaMecano.FonteCorpo(9.75F)}
        AddHandler rbCliente.CheckedChanged, Sub() AplicarTipo()
        AddHandler rbInterno.CheckedChanged, Sub() AplicarTipo()

        txtCliente = New TextBox With {
            .Left = Esc(16), .Top = Esc(70), .Width = Esc(600), .Font = TemaMecano.FonteCorpo(10.0F),
            .BorderStyle = BorderStyle.FixedSingle, .ForeColor = TemaMecano.TextoCampo}
        AddHandler txtCliente.HandleCreated, Sub() SendMessage(txtCliente.Handle, EM_SETCUEBANNER, New IntPtr(1), "Nome ou código")
        AddHandler txtCliente.TextChanged, Sub() AoDigitarCliente()
        AddHandler txtCliente.Enter, Sub()
                                         txtCliente.SelectAll()
                                         If _clienteSelecionado Is Nothing AndAlso txtCliente.Text.Trim() <> "" Then AoDigitarCliente()
                                     End Sub
        AddHandler txtCliente.Leave, Sub()
                                         If Me.ActiveControl Is lstSugestoes Then Return
                                         ResolverClienteDigitado()
                                         EsconderSugestoes()
                                     End Sub
        AddHandler txtCliente.KeyDown, AddressOf TeclaNoCliente

        ' a lista de sugestões fica no corpo (não dentro da seção) para poder descer por cima das outras seções
        lstSugestoes = New ListBox With {
            .Visible = False, .Font = TemaMecano.FonteCorpo(9.75F), .BorderStyle = BorderStyle.FixedSingle,
            .IntegralHeight = False, .ItemHeight = Esc(24), .DrawMode = DrawMode.OwnerDrawFixed,
            .ForeColor = TemaMecano.TextoCampo, .BackColor = TemaMecano.Branco}
        AddHandler lstSugestoes.DrawItem, AddressOf DesenharSugestao
        AddHandler lstSugestoes.MouseClick, Sub(s, e)
                                                Dim i = lstSugestoes.IndexFromPoint(e.Location)
                                                If i >= 0 Then
                                                    lstSugestoes.SelectedIndex = i
                                                    SelecionarSugestao(True)
                                                End If
                                            End Sub
        AddHandler lstSugestoes.KeyDown, Sub(s, e)
                                             If e.KeyCode = Keys.Enter OrElse e.KeyCode = Keys.Tab Then
                                                 e.Handled = True : e.SuppressKeyPress = True
                                                 SelecionarSugestao(True)
                                             ElseIf e.KeyCode = Keys.Escape Then
                                                 e.Handled = True : e.SuppressKeyPress = True
                                                 EsconderSugestoes()
                                                 txtCliente.Focus()
                                             End If
                                         End Sub
        AddHandler lstSugestoes.Leave, Sub()
                                           If Not (Me.ActiveControl Is txtCliente) Then EsconderSugestoes()
                                       End Sub

        lblClienteInfo = New Label With {
            .AutoSize = False, .Left = Esc(16), .Top = Esc(102), .Width = Esc(700), .Height = Esc(20),
            .ForeColor = TemaMecano.CinzaEscuro, .Font = TemaMecano.FonteCorpo(8.75F), .Text = ""}

        cboAtividade = New ComboBox With {
            .Left = Esc(16), .Top = Esc(70), .Width = Esc(600), .DropDownStyle = ComboBoxStyle.DropDownList,
            .Font = TemaMecano.FonteCorpo(10.0F), .FlatStyle = FlatStyle.System, .Visible = False, .DropDownHeight = Esc(320)}
        lblAtividadeAviso = New Label With {
            .AutoSize = False, .Left = Esc(16), .Top = Esc(102), .Width = Esc(700), .Height = Esc(20), .Visible = False,
            .ForeColor = TemaMecano.CinzaEscuro, .Font = TemaMecano.FonteCorpo(8.75F), .Text = ""}

        chkVarias = New CheckBox With {
            .Text = "Lançar para várias empresas", .AutoSize = True,
            .Left = Esc(16), .Top = Esc(130), .Font = TemaMecano.FonteCorpo(9.5F), .Visible = (_id = 0)}
        AddHandler chkVarias.CheckedChanged, Sub() AplicarVarias()

        MontarPainelVarias()
        secOque.Controls.AddRange(New Control() {rbCliente, rbInterno, txtCliente, lblClienteInfo, cboAtividade, lblAtividadeAviso, chkVarias, pnlVarias})

        ' ---------------- quando e quanto tempo
        secQuando = NovaSecao("QUANDO E QUANTO TEMPO", SEC_X, 20 + ALT_OQUE_BASE + 16, ALT_QUANDO)

        Rotulo(secQuando, "Data", 16, 40)
        dtpData = New DateTimePicker With {
            .Left = Esc(16), .Top = Esc(60), .Width = Esc(150), .Format = DateTimePickerFormat.Custom, .CustomFormat = "ddd dd/MM/yyyy",
            .MaxDate = Date.Today, .Value = _dataInicial, .Font = TemaMecano.FonteCorpo(10.0F)}
        AddHandler dtpData.ValueChanged, Sub()
                                             CarregarHorasDoDia()
                                             pnlTopo.Invalidate()
                                         End Sub

        Rotulo(secQuando, "Tempo", 190, 40)
        numHoras = New NumericUpDown With {
            .Left = Esc(190), .Top = Esc(60), .Width = Esc(64), .Minimum = 0, .Maximum = 24, .Increment = 1,
            .Font = TemaMecano.FonteCorpo(10.0F), .TextAlign = HorizontalAlignment.Right}
        Dim lblH = New Label With {.Text = "h", .AutoSize = True, .Left = Esc(258), .Top = Esc(64), .ForeColor = TemaMecano.CinzaEscuro}
        numMinutos = New NumericUpDown With {
            .Left = Esc(280), .Top = Esc(60), .Width = Esc(64), .Minimum = 0, .Maximum = 59, .Increment = 5,
            .Font = TemaMecano.FonteCorpo(10.0F), .TextAlign = HorizontalAlignment.Right}
        Dim lblM = New Label With {.Text = "min", .AutoSize = True, .Left = Esc(348), .Top = Esc(64), .ForeColor = TemaMecano.CinzaEscuro}
        AddHandler numHoras.ValueChanged, Sub() AtualizarEquivale()
        AddHandler numMinutos.ValueChanged, Sub() AtualizarEquivale()
        AddHandler numHoras.Enter, Sub() numHoras.Select(0, 10)
        AddHandler numMinutos.Enter, Sub() numMinutos.Select(0, 10)

        Dim lblAtalhos = New Label With {.Text = "Atalhos", .AutoSize = True, .Left = Esc(16), .Top = Esc(100), .ForeColor = TemaMecano.CinzaEscuro, .Font = TemaMecano.FonteCorpo(8.75F)}
        secQuando.Controls.Add(lblAtalhos)
        Dim x = 16
        For Each atalho In New Integer() {15, 30, 60, 120, 240, 480}
            Dim minutos = atalho
            Dim texto = If(minutos < 60, minutos & " min", (minutos \ 60) & " h")
            Dim b = BotaoMecano.Secundario(texto, Esc(64), Esc(30))
            b.Left = Esc(x) : b.Top = Esc(118)
            b.TabStop = False
            AddHandler b.Click, Sub()
                                    numHoras.Value = minutos \ 60
                                    numMinutos.Value = minutos Mod 60
                                End Sub
            secQuando.Controls.Add(b)
            x += 70
        Next

        lblEquivale = New Label With {
            .AutoSize = False, .Left = Esc(190), .Top = Esc(92), .Width = Esc(520), .Height = Esc(20),
            .ForeColor = TemaMecano.CinzaClaro, .Font = TemaMecano.FonteCorpo(8.75F)}

        ' gestor e diretoria podem lançar em nome de outra pessoa (só em lançamento novo); aparece em CarregarDados
        lblColaborador = New Label With {.Text = "Colaborador", .AutoSize = True, .Left = Esc(420), .Top = Esc(40), .Visible = False,
                                         .ForeColor = TemaMecano.CinzaEscuro, .Font = TemaMecano.FonteCorpo(8.75F)}
        cboColaborador = New ComboBox With {
            .Left = Esc(420), .Top = Esc(60), .Width = Esc(290), .DropDownStyle = ComboBoxStyle.DropDownList, .Visible = False,
            .Font = TemaMecano.FonteCorpo(10.0F), .FlatStyle = FlatStyle.System, .DropDownHeight = Esc(300)}
        AddHandler cboColaborador.SelectedIndexChanged, Sub()
                                                           If _carregando Then Return
                                                           CarregarHorasDoDia(forcar:=True)
                                                           pnlTopo.Invalidate()
                                                       End Sub

        secQuando.Controls.AddRange(New Control() {dtpData, numHoras, lblH, numMinutos, lblM, lblEquivale, lblColaborador, cboColaborador})

        pnlCorpo.Controls.Add(secOque)
        pnlCorpo.Controls.Add(secQuando)
        pnlCorpo.Controls.Add(lstSugestoes)
        lstSugestoes.BringToFront()
        AtualizarEquivale()
    End Sub

    Private Sub MontarPainelVarias()
        pnlVarias = New Panel With {.Left = Esc(16), .Top = Esc(ALT_OQUE_BASE - 4), .Width = Esc(SEC_W - 32), .Height = Esc(ALT_VARIAS - 16), .Visible = False, .BackColor = TemaMecano.Branco}

        Dim lblDisp = New Label With {.Text = "Empresas", .AutoSize = True, .Left = 0, .Top = 0, .ForeColor = TemaMecano.CinzaEscuro, .Font = TemaMecano.FonteCorpo(9.0F, FontStyle.Bold)}
        txtFiltro = New TextBox With {.Left = 0, .Top = Esc(22), .Width = Esc(310), .Font = TemaMecano.FonteCorpo(9.75F), .BorderStyle = BorderStyle.FixedSingle}
        AddHandler txtFiltro.HandleCreated, Sub() SendMessage(txtFiltro.Handle, EM_SETCUEBANNER, New IntPtr(1), "Localizar por nome ou código")
        AddHandler txtFiltro.TextChanged, Sub() PreencherDisponiveis()

        lstDisponiveis = New CheckedListBox With {.Left = 0, .Top = Esc(52), .Width = Esc(310), .Height = Esc(196), .CheckOnClick = True,
                                                  .Font = TemaMecano.FonteCorpo(9.25F), .BorderStyle = BorderStyle.FixedSingle, .IntegralHeight = False}
        chkTodasDisp = New CheckBox With {.Text = "Marcar todas da lista", .AutoSize = True, .Left = 0, .Top = Esc(254), .Font = TemaMecano.FonteCorpo(8.75F)}
        AddHandler chkTodasDisp.CheckedChanged, Sub() MarcarTodos(lstDisponiveis, chkTodasDisp.Checked)

        Dim bx = Esc(326)
        Dim btnUm = BotaoMecano.Secundario(">", Esc(44), Esc(30)) : btnUm.Left = bx : btnUm.Top = Esc(64)
        Dim btnTodos = BotaoMecano.Secundario(">>", Esc(44), Esc(30)) : btnTodos.Left = bx : btnTodos.Top = Esc(102)
        Dim btnVoltaUm = BotaoMecano.Secundario("<", Esc(44), Esc(30)) : btnVoltaUm.Left = bx : btnVoltaUm.Top = Esc(150)
        Dim btnVoltaTodos = BotaoMecano.Secundario("<<", Esc(44), Esc(30)) : btnVoltaTodos.Left = bx : btnVoltaTodos.Top = Esc(188)
        AddHandler btnUm.Click, Sub() Mover(lstDisponiveis, _disponiveis, _selecionados, False)
        AddHandler btnTodos.Click, Sub() Mover(lstDisponiveis, _disponiveis, _selecionados, True)
        AddHandler btnVoltaUm.Click, Sub() Mover(lstSelecionadas, _selecionados, _disponiveis, False)
        AddHandler btnVoltaTodos.Click, Sub() Mover(lstSelecionadas, _selecionados, _disponiveis, True)

        Dim sx = Esc(386)
        Dim lblSel = New Label With {.Text = "Selecionadas para gravar", .AutoSize = True, .Left = sx, .Top = 0, .ForeColor = TemaMecano.CinzaEscuro, .Font = TemaMecano.FonteCorpo(9.0F, FontStyle.Bold)}
        lblContagemSel = New Label With {.AutoSize = False, .Left = sx, .Top = Esc(22), .Width = Esc(310), .Height = Esc(24), .TextAlign = ContentAlignment.MiddleLeft,
                                         .ForeColor = TemaMecano.CinzaClaro, .Font = TemaMecano.FonteCorpo(8.75F), .Text = "Nenhuma empresa selecionada."}
        lstSelecionadas = New CheckedListBox With {.Left = sx, .Top = Esc(52), .Width = Esc(310), .Height = Esc(196), .CheckOnClick = True,
                                                   .Font = TemaMecano.FonteCorpo(9.25F), .BorderStyle = BorderStyle.FixedSingle, .IntegralHeight = False}
        chkTodasSel = New CheckBox With {.Text = "Marcar todas", .AutoSize = True, .Left = sx, .Top = Esc(254), .Font = TemaMecano.FonteCorpo(8.75F)}
        AddHandler chkTodasSel.CheckedChanged, Sub() MarcarTodos(lstSelecionadas, chkTodasSel.Checked)

        ' como o tempo informado abaixo se aplica às empresas da direita
        ' os textos das duas opções mostram os números reais assim que a pessoa informa o tempo (AtualizarTextosModo)
        Dim lblModo = New Label With {.Text = "Como gravar o tempo?", .AutoSize = True, .Left = 0, .Top = Esc(288), .ForeColor = TemaMecano.CinzaEscuro, .Font = TemaMecano.FonteCorpo(9.0F, FontStyle.Bold)}
        rbMesmoTempo = New RadioButton With {.Text = "Cada empresa fica com o tempo inteiro", .AutoSize = True,
                                             .Left = Esc(190), .Top = Esc(286), .Checked = True, .Font = TemaMecano.FonteCorpo(9.25F)}
        rbDividir = New RadioButton With {.Text = "Dividir o tempo entre as empresas", .AutoSize = True,
                                          .Left = Esc(190), .Top = Esc(308), .Font = TemaMecano.FonteCorpo(9.25F)}
        AddHandler rbMesmoTempo.CheckedChanged, Sub() AtualizarEquivale()
        AddHandler rbDividir.CheckedChanged, Sub() AtualizarEquivale()

        pnlVarias.Controls.AddRange(New Control() {lblDisp, txtFiltro, lstDisponiveis, chkTodasDisp, btnUm, btnTodos, btnVoltaUm, btnVoltaTodos, lblSel, lblContagemSel, lstSelecionadas, chkTodasSel, lblModo, rbMesmoTempo, rbDividir})

        ' "Meus grupos": uma linha acima das listas; tudo o que já existe desce para abrir espaço
        Dim dy = Esc(36)
        For Each c As Control In pnlVarias.Controls
            c.Top += dy
        Next
        Dim lblGrupo = New Label With {.Text = "Grupo", .AutoSize = True, .Left = 0, .Top = Esc(6), .ForeColor = TemaMecano.CinzaEscuro, .Font = TemaMecano.FonteCorpo(9.0F, FontStyle.Bold)}
        cboGrupo = New ComboBox With {.Left = Esc(56), .Top = Esc(2), .Width = Esc(314), .DropDownStyle = ComboBoxStyle.DropDownList,
                                      .Font = TemaMecano.FonteCorpo(9.5F), .FlatStyle = FlatStyle.System, .DropDownHeight = Esc(260)}
        AddHandler cboGrupo.SelectedIndexChanged, Sub() If Not _carregando Then AplicarGrupo()
        btnGrupos = BotaoMecano.Secundario("Meus grupos…", Esc(150), Esc(28))
        btnGrupos.Left = Esc(386) : btnGrupos.Top = Esc(1)
        btnGrupos.TabStop = False
        AddHandler btnGrupos.Click, Sub() AbrirGrupos()
        pnlVarias.Controls.AddRange(New Control() {lblGrupo, cboGrupo, btnGrupos})
    End Sub

    '''<summary>Preenche a lista de grupos da pessoa (ou esconde a linha quando a tabela ainda não existe).</summary>
    Private Sub PreencherGrupos(selecionar As String)
        Dim disponivel = ModApontamento.GruposDisponiveis()
        cboGrupo.Visible = disponivel
        btnGrupos.Visible = disponivel
        For Each c As Control In pnlVarias.Controls
            If TypeOf c Is Label AndAlso DirectCast(c, Label).Text = "Grupo" Then c.Visible = disponivel
        Next
        If Not disponivel Then Return
        Try
            _grupos = ModApontamento.Grupos()
        Catch
            _grupos = New List(Of GrupoEmpresas)()
        End Try
        Dim antes = _carregando
        _carregando = True
        cboGrupo.BeginUpdate()
        cboGrupo.Items.Clear()
        cboGrupo.Items.Add(If(_grupos.Count = 0, "Nenhum grupo ainda: clique em Meus grupos…", "Escolha um grupo para marcar as empresas"))
        For Each g In _grupos
            cboGrupo.Items.Add(g)
        Next
        cboGrupo.EndUpdate()
        cboGrupo.SelectedIndex = 0
        If selecionar <> "" Then
            For i = 1 To cboGrupo.Items.Count - 1
                If String.Equals(DirectCast(cboGrupo.Items(i), GrupoEmpresas).Nome, selecionar, StringComparison.OrdinalIgnoreCase) Then cboGrupo.SelectedIndex = i
            Next
        End If
        _carregando = antes
    End Sub

    '''<summary>Passa as empresas do grupo escolhido para a lista da direita (as que já estavam lá ficam).</summary>
    Private Sub AplicarGrupo()
        Dim g = TryCast(cboGrupo.SelectedItem, GrupoEmpresas)
        If g Is Nothing Then Return
        Dim empresas = g.Empresas(_clientes.Where(Function(c) c.Ativo))
        For Each c In empresas
            _disponiveis.Remove(c)
            If Not _selecionados.Contains(c) Then _selecionados.Add(c)
        Next
        PreencherDisponiveis()
        PreencherSelecionadas()
        If empresas.Count = 0 Then
            lblContagemSel.Text = "Esse grupo não tem empresa ativa."
        End If
    End Sub

    '''<summary>Abre "Meus grupos"; se há empresas selecionadas, já entra como um grupo novo com elas.</summary>
    Private Sub AbrirGrupos()
        Dim iniciais = If(_selecionados.Count > 0, _selecionados.Select(Function(c) c.Tradutor).ToList(), Nothing)
        Using f As New frmGruposEmpresas(iniciais)
            f.ShowDialog(Me)
            If f.Alterou Then PreencherGrupos(f.UltimoGravado)
        End Using
    End Sub

    '''<summary>Só para a captura de tela e testes: escolhe um grupo pelo nome.</summary>
    Public Sub EscolherGrupo(nome As String)
        For i = 1 To cboGrupo.Items.Count - 1
            If String.Equals(DirectCast(cboGrupo.Items(i), GrupoEmpresas).Nome, nome, StringComparison.OrdinalIgnoreCase) Then
                cboGrupo.SelectedIndex = i
                Return
            End If
        Next
    End Sub

    ''' <summary>True quando a caixa "várias empresas" está marcada e o tempo deve ser dividido entre elas.</summary>
    Private ReadOnly Property ModoDividir As Boolean
        Get
            Return chkVarias.Checked AndAlso rbDividir IsNot Nothing AndAlso rbDividir.Checked
        End Get
    End Property

    ''' <summary>
    ''' Horas que cada empresa selecionada vai receber, na ordem da lista da direita.
    ''' No modo "mesmo tempo" todas recebem o informado; no rateio, o informado dividido
    ''' em centésimos (lista vazia quando não dá nem 0,01 h por empresa).
    ''' </summary>
    Private Function PartesPorEmpresa(horas As Decimal) As List(Of Decimal)
        Dim n = _selecionados.Count
        If ModoDividir Then Return Apontamento.Ratear(horas, n)
        Return Enumerable.Repeat(horas, n).ToList()
    End Function

    ''' <summary>"0h30 para cada uma" ou, quando as partes diferem por um centésimo, "cerca de 0h40 para cada uma".</summary>
    Private Shared Function DescreverPartes(partes As List(Of Decimal)) As String
        If partes.Count = 0 Then Return ""
        Dim menor = partes.Min()
        Dim maior = partes.Max()
        If menor = maior Then Return Apontamento.FormatarHoras(menor) & " para cada uma"
        Return "cerca de " & Apontamento.FormatarHoras(maior) & " para cada uma"
    End Function

    ''' <summary>
    ''' Põe os números da pessoa nas duas opções, para ela ver o efeito antes de escolher:
    ''' "Cada empresa fica com 2h00  (4 empresas = 8h00 no total)" e
    ''' "Dividir as 2h00 entre as 4 empresas  (0h30 para cada uma)".
    ''' Sem tempo ou sem empresas, ficam as frases genéricas.
    ''' </summary>
    Private Sub AtualizarTextosModo()
        If rbMesmoTempo Is Nothing OrElse numHoras Is Nothing Then Return
        Dim n = _selecionados.Count
        Dim min = MinutosInformados()
        If n = 0 OrElse min = 0 Then
            rbMesmoTempo.Text = "Cada empresa fica com o tempo inteiro"
            rbDividir.Text = "Dividir o tempo entre as empresas"
            Return
        End If
        Dim horas = Apontamento.DeMinutos(min)
        Dim plural = If(n = 1, " empresa", " empresas")
        rbMesmoTempo.Text = "Cada empresa fica com " & Apontamento.FormatarHoras(horas) & "  (total " & Apontamento.FormatarHoras(horas * n) & ")"
        Dim partes = Apontamento.Ratear(horas, n)
        rbDividir.Text = "Dividir " & Apontamento.FormatarHoras(horas) & " entre " & n & plural &
                         "  (" & If(partes.Count = 0, "tempo pequeno demais para dividir", DescreverPartes(partes)) & ")"
    End Sub

    Private Function NovaSecao(titulo As String, x As Integer, y As Integer, altura As Integer) As Panel
        ' fundo branco: os rótulos e caixas de marcar herdam a cor do painel (senão ficam com um retângulo cinza atrás)
        Dim sec As New Panel With {.Left = Esc(x), .Top = Esc(y), .Width = Esc(SEC_W), .Height = Esc(altura), .BackColor = TemaMecano.Branco}
        AddHandler sec.Paint, Sub(s, e)
                                  Dim g = e.Graphics
                                  g.SmoothingMode = SmoothingMode.AntiAlias
                                  Dim r As New Rectangle(0, 0, sec.Width - 1, sec.Height - 1)
                                  Using gp = FormasMecano.Arredondado(r, Esc(10))
                                      Using b As New SolidBrush(TemaMecano.Branco)
                                          g.FillPath(b, gp)
                                      End Using
                                      Using p As New Pen(TemaMecano.BordaCampo)
                                          g.DrawPath(p, gp)
                                      End Using
                                  End Using
                                  Using f = TemaMecano.FonteCorpo(9.0F, FontStyle.Bold)
                                      TextRenderer.DrawText(g, titulo, f, New Point(Esc(16), Esc(14)), TemaMecano.CinzaEscuro, TextFormatFlags.NoPadding)
                                  End Using
                              End Sub
        Return sec
    End Function

    Private Sub Rotulo(pai As Control, texto As String, x As Integer, y As Integer)
        pai.Controls.Add(New Label With {.Text = texto, .AutoSize = True, .Left = Esc(x), .Top = Esc(y), .ForeColor = TemaMecano.CinzaEscuro, .Font = TemaMecano.FonteCorpo(9.0F)})
    End Sub

    ' =================================================================
    '  DADOS
    ' =================================================================
    Private Sub CarregarDados()
        _carregando = True
        Me.Cursor = Cursors.WaitCursor
        Try
            _clientes = ModApontamento.Clientes()
            _atividades = ModApontamento.AtividadesInternas()
            If _id = 0 AndAlso ModApontamento.PodeGerenciar() Then _colaboradores = ModApontamento.Colaboradores()
            CarregarHorasDoDia()
        Catch ex As Exception
            MessageBox.Show(Me, "Não foi possível ler a lista de clientes." & vbCrLf & ex.Message, "Apontamento", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try

        If _colaboradores.Count > 0 Then
            Dim eu = ModApontamento.UsuarioLogadoId()
            cboColaborador.Items.Clear()
            Dim idx = -1
            For Each c In _colaboradores.Where(Function(x) Not x.Desligado)
                cboColaborador.Items.Add(c)
                If c.Id = eu Then idx = cboColaborador.Items.Count - 1
            Next
            If idx < 0 Then
                cboColaborador.Items.Insert(0, New ColaboradorResumo With {.Id = eu, .Nome = ModApontamento.UsuarioLogadoNome()})
                idx = 0
            End If
            cboColaborador.SelectedIndex = idx
            lblColaborador.Visible = True
            cboColaborador.Visible = True
        End If

        cboAtividade.Items.Clear()
        For Each t In _atividades
            cboAtividade.Items.Add(t)
        Next
        If _atividades.Count = 0 Then
            lblAtividadeAviso.Text = "Outros.txt não encontrado na rede: só é possível apontar para clientes."
            rbInterno.Enabled = False
        End If

        If _id > 0 Then
            Try
                _original = ModApontamento.Obter(_id)
            Catch ex As Exception
                MessageBox.Show(Me, "Não foi possível ler o apontamento." & vbCrLf & ex.Message, "Apontamento", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            If _original Is Nothing Then
                MessageBox.Show(Me, "Esse apontamento não existe mais.", "Apontamento", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.DialogResult = DialogResult.Cancel
                Me.Close()
                Return
            End If
            Me.Text = "Apontamento — " & _original.Data.ToString("dd/MM/yyyy")
            dtpData.Value = If(_original.Data > Date.Today, Date.Today, _original.Data)
            Dim min = Apontamento.Minutos(_original.Horas)
            numHoras.Value = Math.Min(24, min \ 60)
            numMinutos.Value = min Mod 60
            If _original.EhCliente Then
                rbCliente.Checked = True
                SelecionarCliente(_original.EmpresaId, _original.Empresa)
            Else
                rbInterno.Checked = True
                If Not cboAtividade.Items.Contains(_original.Empresa) Then cboAtividade.Items.Insert(0, _original.Empresa)   ' texto antigo, fora da lista atual
                cboAtividade.SelectedItem = _original.Empresa
            End If
        End If

        _carregando = False
        AplicarTipo()
        AtualizarEquivale()
        pnlTopo.Invalidate()
        If rbCliente.Checked Then txtCliente.Focus() Else cboAtividade.Focus()
    End Sub

    Private Sub SelecionarCliente(tradutor As String, nome As String)
        Dim cod = If(tradutor, "").Trim().TrimStart("0"c)
        Dim c = _clientes.FirstOrDefault(Function(x) x.Tradutor.TrimStart("0"c) = cod)
        If c Is Nothing Then
            ' cliente que não está mais na tabela: mantém o texto gravado para não perder o lançamento
            c = New ClienteResumo With {.Tradutor = If(tradutor, "").Trim(), .Nome = If(nome, "").Trim(), .Ativo = False}
            _clientes.Insert(0, c)
        End If
        DefinirCliente(c)
    End Sub

    ' ---------------------------------------------------------------- busca do cliente conforme digita
    Private Sub DefinirCliente(c As ClienteResumo)
        _clienteSelecionado = c
        Dim antes = _carregando
        _carregando = True
        txtCliente.Text = If(c Is Nothing, "", c.Rotulo)
        txtCliente.SelectionStart = txtCliente.Text.Length
        _carregando = antes
        EsconderSugestoes()
        MostrarInfoCliente()
    End Sub

    ''' <summary>Só para a captura de tela da demonstração: digita no campo do cliente e mostra as sugestões.</summary>
    Public Sub DigitarCliente(texto As String)
        rbCliente.Checked = True
        txtCliente.Focus()
        txtCliente.Text = texto
        txtCliente.SelectionStart = texto.Length
    End Sub

    ''' <summary>Filtra os clientes por TODAS as palavras digitadas, em qualquer posição do nome ou do código.</summary>
    Private Sub AoDigitarCliente()
        If _carregando Then Return
        If _clienteSelecionado IsNot Nothing AndAlso txtCliente.Text = _clienteSelecionado.Rotulo Then Return
        _clienteSelecionado = Nothing
        Dim termos = Colaborador.SemAcento(txtCliente.Text.Trim()).Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
        If termos.Length = 0 Then
            EsconderSugestoes()
            MostrarInfoCliente()
            Return
        End If
        Dim encontrados = _clientes.Where(Function(c)
                                              Dim alvo = Colaborador.SemAcento(c.Rotulo)
                                              Return termos.All(Function(t) alvo.Contains(t))
                                          End Function).ToList()
        _sugestoes = encontrados.Take(MAX_SUGESTOES).ToList()
        lstSugestoes.BeginUpdate()
        lstSugestoes.Items.Clear()
        For Each c In _sugestoes
            lstSugestoes.Items.Add(c)
        Next
        lstSugestoes.EndUpdate()
        If _sugestoes.Count = 0 Then
            EsconderSugestoes()
            lblClienteInfo.Text = "Nenhuma empresa com """ & txtCliente.Text.Trim() & """."
            lblClienteInfo.ForeColor = TemaMecano.VermelhoPrimario
            Return
        End If
        lstSugestoes.SelectedIndex = 0
        Dim linhas = Math.Min(_sugestoes.Count, 8)
        lstSugestoes.Bounds = New Rectangle(secOque.Left + txtCliente.Left, secOque.Top + txtCliente.Bottom + Esc(2), txtCliente.Width, linhas * lstSugestoes.ItemHeight + Esc(4))
        lstSugestoes.Visible = True
        lstSugestoes.BringToFront()
        lblClienteInfo.Text = If(encontrados.Count > MAX_SUGESTOES, "Mais de " & MAX_SUGESTOES & " empresas",
                                 encontrados.Count & If(encontrados.Count = 1, " empresa", " empresas"))
        lblClienteInfo.ForeColor = TemaMecano.CinzaEscuro
    End Sub

    Private Sub EsconderSugestoes()
        If lstSugestoes IsNot Nothing Then lstSugestoes.Visible = False
    End Sub

    ''' <summary>Pega o item destacado da lista. Devolve True se selecionou.</summary>
    Private Function SelecionarSugestao(irParaOTempo As Boolean) As Boolean
        If Not lstSugestoes.Visible OrElse lstSugestoes.SelectedIndex < 0 Then Return False
        DefinirCliente(DirectCast(lstSugestoes.SelectedItem, ClienteResumo))
        If irParaOTempo Then numHoras.Focus() Else txtCliente.Focus()
        Return True
    End Function

    Private Sub TeclaNoCliente(sender As Object, e As KeyEventArgs)
        If Not lstSugestoes.Visible Then
            If e.KeyCode = Keys.Down AndAlso txtCliente.Text.Trim() <> "" AndAlso _clienteSelecionado Is Nothing Then
                AoDigitarCliente() : e.Handled = True
            End If
            Return
        End If
        Select Case e.KeyCode
            Case Keys.Down
                If lstSugestoes.SelectedIndex < lstSugestoes.Items.Count - 1 Then lstSugestoes.SelectedIndex += 1
                e.Handled = True
            Case Keys.Up
                If lstSugestoes.SelectedIndex > 0 Then lstSugestoes.SelectedIndex -= 1
                e.Handled = True
            Case Keys.Enter
                e.Handled = True : e.SuppressKeyPress = True
                SelecionarSugestao(True)
            Case Keys.Escape
                e.Handled = True : e.SuppressKeyPress = True
                EsconderSugestoes()
        End Select
    End Sub

    ''' <summary>Tab com a lista aberta: seleciona o destacado e segue para o próximo campo (o Tab normal não chega ao KeyDown).</summary>
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        If keyData = Keys.Tab AndAlso txtCliente IsNot Nothing AndAlso txtCliente.Focused AndAlso lstSugestoes.Visible Then
            If SelecionarSugestao(True) Then Return True
        End If
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Private Sub DesenharSugestao(sender As Object, e As DrawItemEventArgs)
        e.DrawBackground()
        If e.Index < 0 Then Return
        Dim c = DirectCast(lstSugestoes.Items(e.Index), ClienteResumo)
        Dim selecionado = (e.State And DrawItemState.Selected) = DrawItemState.Selected
        Using b As New SolidBrush(If(selecionado, Color.FromArgb(252, 232, 232), TemaMecano.Branco))
            e.Graphics.FillRectangle(b, e.Bounds)
        End Using
        Dim x = e.Bounds.X + Esc(8)
        Using fCod = TemaMecano.FonteCorpo(9.0F, FontStyle.Bold)
            TextRenderer.DrawText(e.Graphics, c.Tradutor.PadLeft(4, "0"c), fCod, New Rectangle(x, e.Bounds.Y, Esc(48), e.Bounds.Height),
                                  TemaMecano.CinzaEscuro, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPrefix)
        End Using
        Using fNome = TemaMecano.FonteCorpo(9.75F)
            TextRenderer.DrawText(e.Graphics, c.Nome & If(c.Ativo, "", "   (não cliente)"), fNome,
                                  New Rectangle(x + Esc(52), e.Bounds.Y, e.Bounds.Width - Esc(64), e.Bounds.Height),
                                  If(c.Ativo, TemaMecano.TextoCampo, TemaMecano.CinzaClaro),
                                  TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPrefix)
        End Using
    End Sub

    Private Sub AplicarTipo()
        Dim cliente = rbCliente.Checked
        txtCliente.Visible = cliente
        If Not cliente Then EsconderSugestoes()
        lblClienteInfo.Visible = cliente
        cboAtividade.Visible = Not cliente
        lblAtividadeAviso.Visible = Not cliente
        chkVarias.Enabled = cliente
        If Not cliente AndAlso chkVarias.Checked Then chkVarias.Checked = False
        If Not _carregando Then
            If cliente Then txtCliente.Focus() Else cboAtividade.Focus()
        End If
    End Sub

    Private Sub AplicarVarias()
        Dim varias = chkVarias.Checked
        pnlVarias.Visible = varias
        txtCliente.Enabled = Not varias
        If varias Then EsconderSugestoes()
        lblClienteInfo.Visible = Not varias AndAlso rbCliente.Checked
        secOque.Height = Esc(If(varias, ALT_OQUE_BASE + ALT_VARIAS, ALT_OQUE_BASE))
        secQuando.Top = secOque.Bottom + Esc(16)
        secOque.Invalidate()
        AjustarAltura()
        If varias Then
            _disponiveis.Clear()
            _disponiveis.AddRange(_clientes.Where(Function(c) c.Ativo))
            _selecionados.Clear()
            rbMesmoTempo.Checked = True      ' abre sempre no comportamento do Time de hoje
            txtFiltro.Text = ""
            PreencherDisponiveis()
            PreencherSelecionadas()
            PreencherGrupos("")
            If cboGrupo.Visible AndAlso _grupos.Count > 0 Then cboGrupo.Focus() Else txtFiltro.Focus()
        End If
    End Sub

    ''' <summary>Só para a captura de tela da demonstração.</summary>
    Public Sub DefinirVarias(ativo As Boolean)
        If chkVarias.Visible Then chkVarias.Checked = ativo
    End Sub

    ''' <summary>Só para a captura de tela da demonstração: passa as N primeiras empresas para a direita, liga o rateio e informa o tempo.</summary>
    Public Sub DemonstrarRateio(quantidade As Integer, minutos As Integer)
        DefinirVarias(True)
        For Each c In _disponiveis.OrderBy(Function(x) x.Nome).Take(quantidade).ToList()
            _disponiveis.Remove(c)
            _selecionados.Add(c)
        Next
        PreencherDisponiveis()
        PreencherSelecionadas()
        rbDividir.Checked = True
        numHoras.Value = minutos \ 60
        numMinutos.Value = minutos Mod 60
        AtualizarEquivale()
    End Sub

    Private Sub AjustarAltura()
        Dim alturaCorpo = Esc(20) + secOque.Height + Esc(16) + secQuando.Height + Esc(20)
        Dim desejada = Esc(ALT_TOPO) + alturaCorpo + Esc(ALT_RODAPE)
        Dim maxima = Screen.FromControl(Me).WorkingArea.Height - Esc(60)
        Me.ClientSize = New Size(Esc(LARGURA), Math.Min(desejada, maxima))
        If Me.Visible Then Me.CenterToParent()
    End Sub

    Private Sub PreencherDisponiveis()
        Dim termo = Colaborador.SemAcento(txtFiltro.Text.Trim())
        lstDisponiveis.BeginUpdate()
        lstDisponiveis.Items.Clear()
        For Each c In _disponiveis.OrderBy(Function(x) x.Nome)
            If termo = "" OrElse Colaborador.SemAcento(c.Rotulo).Contains(termo) Then lstDisponiveis.Items.Add(c)
        Next
        lstDisponiveis.EndUpdate()
        chkTodasDisp.Checked = False
    End Sub

    Private Sub PreencherSelecionadas()
        lstSelecionadas.BeginUpdate()
        lstSelecionadas.Items.Clear()
        For Each c In _selecionados.OrderBy(Function(x) x.Nome)
            lstSelecionadas.Items.Add(c)
        Next
        lstSelecionadas.EndUpdate()
        chkTodasSel.Checked = False
        ' a quantidade de empresas muda o "cada" dos dois textos (contagem e equivalência)
        If lblEquivale IsNot Nothing Then AtualizarEquivale() Else AtualizarContagemSel()
    End Sub

    ''' <summary>Texto acima da lista da direita: quantas empresas e quanto cada uma recebe.</summary>
    Private Sub AtualizarContagemSel()
        If lblContagemSel Is Nothing Then Return
        Dim n = _selecionados.Count
        Dim min = MinutosInformados()
        Dim texto As String
        If n = 0 Then
            texto = "Nenhuma empresa selecionada."
        ElseIf min = 0 Then
            texto = n & If(n = 1, " empresa selecionada.", " empresas selecionadas.")
        Else
            Dim horas = Apontamento.DeMinutos(min)
            Dim partes = PartesPorEmpresa(horas)
            If partes.Count = 0 Then
                texto = n & " empresas: " & Apontamento.FormatarHoras(horas) & " é pouco para dividir entre todas."
            Else
                texto = n & If(n = 1, " empresa: ", " empresas: ") & DescreverPartes(partes) & ". Total: " & Apontamento.FormatarHoras(partes.Sum()) & "."
            End If
        End If
        lblContagemSel.Text = texto
        lblContagemSel.ForeColor = If(n = 0, TemaMecano.CinzaClaro, TemaMecano.CinzaEscuro)
        AtualizarTextosModo()
    End Sub

    Private Sub MarcarTodos(lista As CheckedListBox, marcar As Boolean)
        For i = 0 To lista.Items.Count - 1
            lista.SetItemChecked(i, marcar)
        Next
    End Sub

    Private Sub Mover(origemLista As CheckedListBox, origem As List(Of ClienteResumo), destino As List(Of ClienteResumo), todosVisiveis As Boolean)
        Dim itens As New List(Of ClienteResumo)()
        If todosVisiveis Then
            itens.AddRange(origemLista.Items.Cast(Of ClienteResumo)())
        Else
            itens.AddRange(origemLista.CheckedItems.Cast(Of ClienteResumo)())
        End If
        If itens.Count = 0 Then Return
        For Each c In itens
            origem.Remove(c)
            If Not destino.Contains(c) Then destino.Add(c)
        Next
        PreencherDisponiveis()
        PreencherSelecionadas()
    End Sub

    Private Sub MostrarInfoCliente()
        Dim c = _clienteSelecionado
        If c Is Nothing Then
            lblClienteInfo.Text = ""
            lblClienteInfo.ForeColor = TemaMecano.CinzaEscuro
            Return
        End If
        Dim partes As New List(Of String)()
        partes.Add("cód. " & c.Tradutor)
        If c.CnpjFormatado <> "" Then partes.Add("CNPJ " & c.CnpjFormatado)
        If c.Grupo <> "" Then partes.Add("grupo " & c.CodGrupo & " · " & c.Grupo)
        If Not c.Ativo Then partes.Add("NÃO CLIENTE")
        lblClienteInfo.Text = String.Join("   ·   ", partes)
        lblClienteInfo.ForeColor = If(c.Ativo, TemaMecano.CinzaEscuro, TemaMecano.VermelhoPrimario)
    End Sub

    ''' <summary>Saiu do campo sem escolher na lista: aceita se o texto bate exatamente com um cliente (nome ou código) ou se só sobrou um na busca.</summary>
    Private Sub ResolverClienteDigitado()
        If _clienteSelecionado IsNot Nothing OrElse txtCliente.Text.Trim() = "" Then Return
        Dim texto = Colaborador.SemAcento(txtCliente.Text.Trim())
        Dim exato = _clientes.FirstOrDefault(Function(c) Colaborador.SemAcento(c.Rotulo) = texto OrElse Colaborador.SemAcento(c.Nome) = texto OrElse
                                                        (texto <> "" AndAlso c.Tradutor.TrimStart("0"c) = texto.TrimStart("0"c)))
        If exato Is Nothing AndAlso _sugestoes.Count = 1 Then exato = _sugestoes(0)
        If exato IsNot Nothing Then DefinirCliente(exato)
    End Sub

    Private Function MinutosInformados() As Integer
        Return CInt(numHoras.Value) * 60 + CInt(numMinutos.Value)
    End Function

    Private Sub AtualizarEquivale()
        Dim min = MinutosInformados()
        If min = 0 Then
            lblEquivale.Text = ""
        ElseIf chkVarias.Checked AndAlso _selecionados.Count > 1 Then
            Dim horas = Apontamento.DeMinutos(min)
            Dim partes = PartesPorEmpresa(horas)
            If ModoDividir Then
                lblEquivale.Text = "= " & If(partes.Count = 0, "tempo pequeno demais para dividir", DescreverPartes(partes))
            Else
                lblEquivale.Text = "= " & Apontamento.FormatarHoras(horas) & " para cada empresa  ·  total " & Apontamento.FormatarHoras(partes.Sum())
            End If
        Else
            lblEquivale.Text = "= " & Apontamento.FormatarHoras(Apontamento.DeMinutos(min))
        End If
        AtualizarContagemSel()
        If pnlTopo IsNot Nothing Then pnlTopo.Invalidate()   ' o subtítulo mostra o efeito do tempo digitado no dia
    End Sub

    ' =================================================================
    '  VALIDAR E GRAVAR
    ' =================================================================
    Private Function Validar() As Boolean
        If dtpData.Value.Date > Date.Today Then
            Aviso("A data não pode ser maior que hoje.", dtpData) : Return False
        End If
        If MinutosInformados() = 0 Then
            Aviso("Tempo zerado não pode ser gravado. Informe as horas e os minutos.", numHoras) : Return False
        End If
        If chkVarias.Checked Then
            If _selecionados.Count = 0 Then
                Aviso("Passe ao menos uma empresa para a lista da direita.", txtFiltro) : Return False
            End If
            If ModoDividir AndAlso PartesPorEmpresa(Apontamento.DeMinutos(MinutosInformados())).Count = 0 Then
                Aviso("O tempo é pouco para dividir entre " & _selecionados.Count & " empresas: algumas ficariam com zero." & vbCrLf &
                      "Aumente o tempo ou tire empresas da lista.", numHoras) : Return False
            End If
        ElseIf rbCliente.Checked Then
            If _clienteSelecionado Is Nothing AndAlso lstSugestoes.Visible Then SelecionarSugestao(False)
            ResolverClienteDigitado()
            If _clienteSelecionado Is Nothing Then
                Aviso("Escolha um cliente da lista (ou marque 'Atividade interna').", txtCliente) : Return False
            End If
        Else
            If cboAtividade.SelectedIndex < 0 Then
                Aviso("Escolha a atividade interna.", cboAtividade) : Return False
            End If
        End If
        Return ValidarLimiteDoDia()
    End Function

    ''' <summary>Trava de 10 h por dia: o que já está no dia (lido agora, do banco) mais o que vai ser gravado.</summary>
    Private Function ValidarLimiteDoDia() As Boolean
        CarregarHorasDoDia(forcar:=True)
        Dim dia = dtpData.Value.Date
        Dim novo = TotalASalvar()
        Dim cabe = BancoHoras.CabeNoDia(_jaNoDia)
        If novo <= cabe Then Return True
        Dim quando = If(dia = Date.Today, "hoje", "em " & dia.ToString("dd/MM"))
        Dim limite = Apontamento.FormatarHoras(BancoHoras.LIMITE_DIA)
        Dim quem = If(OutraPessoa(), NomeAlvo() & " já tem ", "Você já tem ")
        If cabe <= 0D Then
            Aviso(quem & Apontamento.FormatarHoras(_jaNoDia) & " " & quando & ". O limite é " & limite & " por dia, então não dá para lançar mais nesse dia.", numHoras)
        Else
            Aviso(quem & Apontamento.FormatarHoras(_jaNoDia) & " " & quando & ". O limite é " & limite & " por dia, então cabem mais " & Apontamento.FormatarHoras(cabe) & "." & vbCrLf &
                  If(chkVarias.Checked, "Diminua o tempo ou tire empresas da lista.", "Diminua o tempo."), numHoras)
        End If
        Return False
    End Function

    Private Sub Aviso(texto As String, campo As Control)
        MessageBox.Show(Me, texto, "Apontamento", MessageBoxButtons.OK, MessageBoxIcon.Information)
        If campo IsNot Nothing AndAlso campo.CanFocus Then campo.Focus()
    End Sub

    Private Sub Salvar(eNovo As Boolean)
        If Not Validar() Then Return
        Dim horas = Apontamento.DeMinutos(MinutosInformados())
        Dim dataLanc = dtpData.Value.Date
        Dim totalGravado = horas
        Me.Cursor = Cursors.WaitCursor
        Try
            If chkVarias.Checked Then
                ' uma linha por empresa, na ordem da lista da direita; no rateio a sobra dos
                ' centésimos fica com as primeiras (Apontamento.Ratear)
                Dim empresas = _selecionados.OrderBy(Function(x) x.Nome).ToList()
                Dim partes = PartesPorEmpresa(horas)
                Dim n = 0
                Dim ultimo = 0
                For i = 0 To empresas.Count - 1
                    Dim c = empresas(i)
                    ultimo = ModApontamento.Inserir(New Apontamento With {.Data = dataLanc, .Empresa = c.Nome, .EmpresaId = c.Tradutor, .Horas = partes(i), .UsuarioId = UsuarioAlvo()})
                    n += 1
                Next
                totalGravado = partes.Sum()
                IdGravado = ultimo
                QuantidadeGravada += n
                Dim cabecalho = n & If(n = 1, " apontamento gravado em ", " apontamentos gravados em ") & dataLanc.ToString("dd/MM") & ": "
                If n = 1 Then
                    MensagemResultado = cabecalho & Apontamento.FormatarHoras(totalGravado) & " em " & empresas(0).Nome & "."
                ElseIf ModoDividir Then
                    MensagemResultado = cabecalho & Apontamento.FormatarHoras(horas) & " dividido, " & DescreverPartes(partes) & "."
                Else
                    MensagemResultado = cabecalho & Apontamento.FormatarHoras(horas) & " para cada uma, " & Apontamento.FormatarHoras(totalGravado) & " no total."
                End If
            Else
                Dim a = If(_original Is Nothing, New Apontamento(), _original.Clonar())
                a.Data = dataLanc
                a.Horas = horas
                a.UsuarioId = If(a.UsuarioId > 0, a.UsuarioId, UsuarioAlvo())
                If rbCliente.Checked Then
                    Dim c = _clienteSelecionado
                    a.Empresa = c.Nome
                    a.EmpresaId = c.Tradutor
                Else
                    a.Empresa = cboAtividade.SelectedItem.ToString()
                    a.EmpresaId = ""
                End If
                If a.Id = 0 Then
                    IdGravado = ModApontamento.Inserir(a)
                Else
                    ModApontamento.Alterar(a)
                    IdGravado = a.Id
                End If
                QuantidadeGravada += 1
                MensagemResultado = If(_original Is Nothing, "Apontamento gravado: ", "Apontamento alterado: ") &
                                    Apontamento.FormatarHoras(horas) & " em " & a.Empresa & " (" & dataLanc.ToString("dd/MM") & ")."
            End If
            If OutraPessoa() Then MensagemResultado &= "  Em nome de " & NomeAlvo() & "."
            Dim totalDia = _jaNoDia + totalGravado
            If BancoHoras.ATIVO Then
                ' o que passou da jornada do dia é crédito no banco de horas (só informação: a conta é refeita a partir da TbTime)
                Dim extra = BancoHoras.ExtraDoLancamento(dataLanc, _jaNoDia, totalGravado)
                If extra > 0D Then MensagemResultado &= "  " & Apontamento.FormatarHoras(extra) & " foi para o banco de horas."
            ElseIf totalDia > ResumoHoras.JORNADA Then
                ' aviso de 8 h (decisão de 19/10/2026): a pessoa é avisada na tela, mas grava; o teto de 10 h é a trava
                AvisoJornada = If(dataLanc = Date.Today, "Hoje", "O dia " & dataLanc.ToString("dd/MM")) & " ficou com " & Apontamento.FormatarHoras(totalDia) &
                               ", acima de " & Apontamento.FormatarHoras(ResumoHoras.JORNADA) & "."
                MensagemResultado &= "  " & AvisoJornada
            End If
            ModApontamento.AposGravar()
        Catch ex As Exception
            MessageBox.Show(Me, "Não foi possível gravar." & vbCrLf & ex.Message, "Apontamento", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        Finally
            Me.Cursor = Cursors.Default
        End Try

        _gravou = True
        If AvisoJornada <> "" Then
            MessageBox.Show(Me, AvisoJornada, "Apontamento", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        If eNovo Then
            PrepararProximo(totalGravado)
        Else
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End If
    End Sub

    ''' <summary>"Salvar e novo": mantém a data e o tipo, limpa o resto e atualiza o contexto do dia.</summary>
    Private Sub PrepararProximo(totalGravado As Decimal)
        ' totalGravado já é a soma das linhas gravadas (uma, N iguais ou o rateio)
        If _resumo IsNot Nothing AndAlso dtpData.Value.Date = _resumo.Hoje AndAlso Not OutraPessoa() Then
            _resumo.HorasHoje += totalGravado
        End If
        CarregarHorasDoDia(forcar:=True)     ' o dia agora tem o que acabou de ser gravado
        If chkVarias.Checked Then chkVarias.Checked = False
        DefinirCliente(Nothing)
        cboAtividade.SelectedIndex = -1
        numHoras.Value = 0
        numMinutos.Value = 0
        lblRodape.Text = MensagemResultado
        lblRodape.ForeColor = If(AvisoJornada <> "", TemaMecano.VermelhoPrimario, Color.FromArgb(22, 101, 52))
        AvisoJornada = ""
        pnlTopo.Invalidate()
        If rbCliente.Checked Then txtCliente.Focus() Else cboAtividade.Focus()
    End Sub

    Private Sub AoTeclar(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Escape Then
            Me.DialogResult = If(_gravou, DialogResult.OK, DialogResult.Cancel)
            Me.Close()
            e.Handled = True
        ElseIf e.Control AndAlso e.KeyCode = Keys.S Then
            e.Handled = True : e.SuppressKeyPress = True
            Salvar(False)
        ElseIf e.Control AndAlso e.KeyCode = Keys.Enter Then
            e.Handled = True : e.SuppressKeyPress = True
            Salvar(_id = 0)
        End If
    End Sub

End Class

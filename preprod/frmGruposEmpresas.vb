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
'  frmGruposEmpresas - "Meus grupos de empresas": cada pessoa monta os grupos
'  que marca sempre juntos ("minha carteira", "fechamento do mês"…) para, no
'  lançamento com várias empresas, escolher o grupo em vez de marcar uma a uma.
'
'  Um grupo é um nome e uma lista de códigos de cliente (TRADUTOR), da pessoa
'  logada. Gravado na TbTimeGrupoEmpresa (script 2026-10-20-Grupos-de-Empresas.sql)
'  pelo ModApontamento.SalvarGrupo / ExcluirGrupo.
'
'  Destino no projeto: Grupo_Mecano\Cadastro\frmGruposEmpresas.vb
' ============================================================================
Public Class frmGruposEmpresas
    Inherits Form

    Private Const LARGURA As Integer = 820
    Private Const ALT_TOPO As Integer = 72
    Private Const ALT_RODAPE As Integer = 64
    Private Const EM_SETCUEBANNER As Integer = &H1501

    <Runtime.InteropServices.DllImport("user32.dll", CharSet:=Runtime.InteropServices.CharSet.Unicode)>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As String) As IntPtr
    End Function

    Private _escala As Single = 1.0F
    Private Function Esc(v As Integer) As Integer
        Return CInt(Math.Round(v * _escala))
    End Function

    ' ---------- dados ----------
    Private _clientes As New List(Of ClienteResumo)()
    Private _grupos As New List(Of GrupoEmpresas)()
    Private _atual As GrupoEmpresas = Nothing          ' o grupo em edição (cópia)
    Private _nomeAnterior As String = ""               ' nome gravado antes de uma renomeação ("" = grupo novo)
    Private ReadOnly _iniciais As List(Of String)      ' códigos para começar um grupo novo (vindos do lançamento)
    Private ReadOnly _disponiveis As New List(Of ClienteResumo)()
    Private ReadOnly _noGrupo As New List(Of ClienteResumo)()
    Private _carregando As Boolean = True

    ''' <summary>True quando algum grupo foi gravado ou excluído (quem abriu recarrega a lista).</summary>
    Public Property Alterou As Boolean = False
    ''' <summary>Nome do último grupo gravado (para quem abriu já deixar selecionado).</summary>
    Public Property UltimoGravado As String = ""

    ' ---------- controles ----------
    Private pnlTopo As Panel
    Private pnlCorpo As Panel
    Private pnlRodape As Panel
    Private lstGrupos As ListBox
    Private btnNovo As BotaoMecano
    Private btnExcluir As BotaoMecano
    Private txtNome As TextBox
    Private txtFiltro As TextBox
    Private lstDisponiveis As CheckedListBox
    Private lstNoGrupo As CheckedListBox
    Private lblContagem As Label
    Private btnSalvar As BotaoMecano
    Private btnFechar As BotaoMecano
    Private lblRodape As Label

    ''' <param name="codigosIniciais">códigos (TRADUTOR) para começar um grupo novo já preenchido; Nothing abre no primeiro grupo da pessoa.</param>
    Public Sub New(Optional codigosIniciais As IEnumerable(Of String) = Nothing)
        _iniciais = If(codigosIniciais Is Nothing, New List(Of String)(), codigosIniciais.ToList())
        Using g = Me.CreateGraphics()
            _escala = g.DpiX / 96.0F
        End Using

        Me.AutoScaleMode = AutoScaleMode.None
        Me.Text = "Meus grupos de empresas"
        Me.BackColor = TemaMecano.FundoTela
        Me.Font = TemaMecano.FonteCorpo(9.75F)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.KeyPreview = True
        Me.DoubleBuffered = True
        Me.ClientSize = New Size(Esc(LARGURA), Esc(ALT_TOPO + 490 + ALT_RODAPE))

        MontarTopo()
        MontarRodape()
        MontarCorpo()
        Me.Controls.Add(pnlCorpo)
        Me.Controls.Add(pnlTopo)
        Me.Controls.Add(pnlRodape)
        pnlCorpo.BringToFront()

        AddHandler Me.Load, Sub() CarregarDados()
        AddHandler Me.KeyDown, Sub(s, e)
                                   If e.KeyCode = Keys.Escape Then
                                       e.Handled = True
                                       Me.Close()
                                   ElseIf e.Control AndAlso e.KeyCode = Keys.S Then
                                       e.Handled = True
                                       Salvar()
                                   End If
                               End Sub
        AddHandler Me.FormClosing, AddressOf AoFechar
    End Sub

    ' =================================================================
    '  TOPO
    ' =================================================================
    Private Sub MontarTopo()
        pnlTopo = New Panel With {.Dock = DockStyle.Top, .Height = Esc(ALT_TOPO), .BackColor = TemaMecano.Branco}
        AddHandler pnlTopo.Paint, Sub(s, e)
                                      Dim g = e.Graphics
                                      g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
                                      Using p As New Pen(TemaMecano.BordaCampo)
                                          g.DrawLine(p, 0, pnlTopo.Height - 1, pnlTopo.Width, pnlTopo.Height - 1)
                                      End Using
                                      Using f = TemaMecano.FonteTitulo(15.0F)
                                          TextRenderer.DrawText(g, "Meus grupos de empresas", f, New Rectangle(Esc(24), 0, pnlTopo.Width - Esc(48), pnlTopo.Height - 1), TemaMecano.TextoCampo,
                                                                TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPrefix)
                                      End Using
                                  End Sub
    End Sub

    ' =================================================================
    '  CORPO: grupos à esquerda; nome e empresas à direita
    ' =================================================================
    Private Sub MontarCorpo()
        pnlCorpo = New Panel With {.Dock = DockStyle.Fill, .BackColor = TemaMecano.FundoTela}

        ' ---- esquerda: a lista de grupos
        Dim secGrupos = NovaSecao("GRUPOS", 24, 20, 230, 450)
        lstGrupos = New ListBox With {.Left = Esc(16), .Top = Esc(40), .Width = Esc(198), .Height = Esc(296), .Font = TemaMecano.FonteCorpo(9.75F),
                                      .BorderStyle = BorderStyle.FixedSingle, .IntegralHeight = False}
        AddHandler lstGrupos.SelectedIndexChanged, Sub() If Not _carregando Then EscolherGrupo()
        btnNovo = BotaoMecano.Secundario("Novo grupo", Esc(198), Esc(32))
        btnNovo.Left = Esc(16) : btnNovo.Top = Esc(352)
        AddHandler btnNovo.Click, Sub() If ConfirmarDescartar() Then Novo(Nothing)
        btnExcluir = BotaoMecano.Secundario("Excluir", Esc(198), Esc(32), perigo:=True)
        btnExcluir.Left = Esc(16) : btnExcluir.Top = Esc(392)
        AddHandler btnExcluir.Click, Sub() Excluir()
        secGrupos.Controls.AddRange(New Control() {lstGrupos, btnNovo, btnExcluir})

        ' ---- direita: nome e empresas do grupo
        Dim secGrupo = NovaSecao("O GRUPO", 270, 20, 526, 450)
        Rotulo(secGrupo, "Nome do grupo", 16, 40)
        txtNome = New TextBox With {.Left = Esc(16), .Top = Esc(60), .Width = Esc(494), .Font = TemaMecano.FonteCorpo(10.0F), .BorderStyle = BorderStyle.FixedSingle,
                                    .MaxLength = GrupoEmpresas.TAMANHO_NOME, .ForeColor = TemaMecano.TextoCampo}
        AddHandler txtNome.HandleCreated, Sub() SendMessage(txtNome.Handle, EM_SETCUEBANNER, New IntPtr(1), "Ex.: Minha carteira, Fechamento do mês")

        Dim lblDisp = New Label With {.Text = "Empresas", .AutoSize = True, .Left = Esc(16), .Top = Esc(104), .ForeColor = TemaMecano.CinzaEscuro, .Font = TemaMecano.FonteCorpo(9.0F, FontStyle.Bold)}
        txtFiltro = New TextBox With {.Left = Esc(16), .Top = Esc(126), .Width = Esc(210), .Font = TemaMecano.FonteCorpo(9.5F), .BorderStyle = BorderStyle.FixedSingle}
        AddHandler txtFiltro.HandleCreated, Sub() SendMessage(txtFiltro.Handle, EM_SETCUEBANNER, New IntPtr(1), "Localizar por nome ou código")
        AddHandler txtFiltro.TextChanged, Sub() PreencherDisponiveis()
        lstDisponiveis = New CheckedListBox With {.Left = Esc(16), .Top = Esc(154), .Width = Esc(210), .Height = Esc(236), .CheckOnClick = True,
                                                  .Font = TemaMecano.FonteCorpo(9.0F), .BorderStyle = BorderStyle.FixedSingle, .IntegralHeight = False}

        Dim bx = Esc(238)
        Dim btnUm = BotaoMecano.Secundario(">", Esc(40), Esc(30)) : btnUm.Left = bx : btnUm.Top = Esc(190)
        Dim btnTodos = BotaoMecano.Secundario(">>", Esc(40), Esc(30)) : btnTodos.Left = bx : btnTodos.Top = Esc(228)
        Dim btnVoltaUm = BotaoMecano.Secundario("<", Esc(40), Esc(30)) : btnVoltaUm.Left = bx : btnVoltaUm.Top = Esc(276)
        Dim btnVoltaTodos = BotaoMecano.Secundario("<<", Esc(40), Esc(30)) : btnVoltaTodos.Left = bx : btnVoltaTodos.Top = Esc(314)
        AddHandler btnUm.Click, Sub() Mover(lstDisponiveis, _disponiveis, _noGrupo, False)
        AddHandler btnTodos.Click, Sub() Mover(lstDisponiveis, _disponiveis, _noGrupo, True)
        AddHandler btnVoltaUm.Click, Sub() Mover(lstNoGrupo, _noGrupo, _disponiveis, False)
        AddHandler btnVoltaTodos.Click, Sub() Mover(lstNoGrupo, _noGrupo, _disponiveis, True)

        Dim sx = Esc(290)
        Dim lblSel = New Label With {.Text = "No grupo", .AutoSize = True, .Left = sx, .Top = Esc(104), .ForeColor = TemaMecano.CinzaEscuro, .Font = TemaMecano.FonteCorpo(9.0F, FontStyle.Bold)}
        lblContagem = New Label With {.AutoSize = False, .Left = sx, .Top = Esc(126), .Width = Esc(220), .Height = Esc(24), .TextAlign = ContentAlignment.MiddleLeft,
                                      .ForeColor = TemaMecano.CinzaClaro, .Font = TemaMecano.FonteCorpo(8.75F)}
        lstNoGrupo = New CheckedListBox With {.Left = sx, .Top = Esc(154), .Width = Esc(220), .Height = Esc(236), .CheckOnClick = True,
                                              .Font = TemaMecano.FonteCorpo(9.0F), .BorderStyle = BorderStyle.FixedSingle, .IntegralHeight = False}

        secGrupo.Controls.AddRange(New Control() {txtNome, lblDisp, txtFiltro, lstDisponiveis, btnUm, btnTodos, btnVoltaUm, btnVoltaTodos, lblSel, lblContagem, lstNoGrupo})

        pnlCorpo.Controls.Add(secGrupos)
        pnlCorpo.Controls.Add(secGrupo)
    End Sub

    Private Function NovaSecao(titulo As String, x As Integer, y As Integer, largura As Integer, altura As Integer) As Panel
        Dim p As New Panel With {.Left = Esc(x), .Top = Esc(y), .Width = Esc(largura), .Height = Esc(altura), .BackColor = TemaMecano.Branco}
        AddHandler p.Paint, Sub(s, e)
                                Dim g = e.Graphics
                                g.SmoothingMode = SmoothingMode.AntiAlias
                                g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
                                Using gp = FormasMecano.Arredondado(New Rectangle(0, 0, p.Width - 1, p.Height - 1), Esc(10))
                                    Using pen As New Pen(TemaMecano.BordaCampo)
                                        g.DrawPath(pen, gp)
                                    End Using
                                End Using
                                Using f = TemaMecano.FonteCorpo(9.0F, FontStyle.Bold)
                                    TextRenderer.DrawText(g, titulo, f, New Point(Esc(16), Esc(12)), TemaMecano.CinzaEscuro, TextFormatFlags.NoPadding)
                                End Using
                            End Sub
        Return p
    End Function

    Private Sub Rotulo(pai As Control, texto As String, x As Integer, y As Integer)
        pai.Controls.Add(New Label With {.Text = texto, .AutoSize = True, .Left = Esc(x), .Top = Esc(y), .ForeColor = TemaMecano.CinzaEscuro, .Font = TemaMecano.FonteCorpo(8.75F)})
    End Sub

    ' =================================================================
    '  RODAPÉ
    ' =================================================================
    Private Sub MontarRodape()
        pnlRodape = New Panel With {.Dock = DockStyle.Bottom, .Height = Esc(ALT_RODAPE), .BackColor = TemaMecano.Branco}
        AddHandler pnlRodape.Paint, Sub(s, e)
                                        Using p As New Pen(TemaMecano.BordaCampo)
                                            e.Graphics.DrawLine(p, 0, 0, pnlRodape.Width, 0)
                                        End Using
                                    End Sub
        btnSalvar = BotaoMecano.Primario("Salvar grupo", Esc(140), Esc(36))
        btnSalvar.Top = Esc(14)
        btnSalvar.Left = Esc(LARGURA) - Esc(24) - btnSalvar.Width
        AddHandler btnSalvar.Click, Sub() Salvar()
        btnFechar = BotaoMecano.Secundario("Fechar", Esc(110), Esc(36))
        btnFechar.Top = Esc(14)
        btnFechar.Left = btnSalvar.Left - Esc(8) - btnFechar.Width
        AddHandler btnFechar.Click, Sub() Me.Close()
        lblRodape = New Label With {.AutoSize = False, .Left = Esc(24), .Top = Esc(14), .Width = btnFechar.Left - Esc(36), .Height = Esc(36), .TextAlign = ContentAlignment.MiddleLeft,
                                    .ForeColor = TemaMecano.CinzaClaro, .Font = TemaMecano.FonteCorpo(9.0F), .Text = "Ctrl+S salva  ·  Esc fecha"}
        pnlRodape.Controls.AddRange(New Control() {lblRodape, btnFechar, btnSalvar})
    End Sub

    ' =================================================================
    '  DADOS
    ' =================================================================
    Private Sub CarregarDados()
        _carregando = True
        Me.Cursor = Cursors.WaitCursor
        Try
            _clientes = ModApontamento.Clientes()
            _grupos = ModApontamento.Grupos()
        Catch ex As Exception
            MessageBox.Show(Me, "Não foi possível ler os grupos." & vbCrLf & ex.Message, "Meus grupos", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
        PreencherGrupos(Nothing)
        If _iniciais.Count > 0 OrElse _grupos.Count = 0 Then
            Novo(_iniciais)
        Else
            lstGrupos.SelectedIndex = 0
            EscolherGrupo()
        End If
        _carregando = False
        txtNome.Focus()
    End Sub

    Private Sub PreencherGrupos(selecionar As String)
        _carregando = True
        lstGrupos.BeginUpdate()
        lstGrupos.Items.Clear()
        For Each g In _grupos.OrderBy(Function(x) x.Nome)
            lstGrupos.Items.Add(g)
        Next
        lstGrupos.EndUpdate()
        If selecionar <> "" Then
            For i = 0 To lstGrupos.Items.Count - 1
                If String.Equals(DirectCast(lstGrupos.Items(i), GrupoEmpresas).Nome, selecionar, StringComparison.OrdinalIgnoreCase) Then lstGrupos.SelectedIndex = i
            Next
        End If
        _carregando = False
    End Sub

    ''' <summary>Começa um grupo novo, vazio ou já com os códigos recebidos.</summary>
    Private Sub Novo(codigos As IEnumerable(Of String))
        _carregando = True
        lstGrupos.ClearSelected()
        _atual = New GrupoEmpresas With {.UsuarioId = ModApontamento.UsuarioLogadoId()}
        If codigos IsNot Nothing Then _atual.Codigos.AddRange(codigos)
        _nomeAnterior = ""
        txtNome.Text = ""
        MostrarGrupo()
        btnExcluir.Enabled = False
        _carregando = False
        txtNome.Focus()
    End Sub

    Private Sub EscolherGrupo()
        Dim g = TryCast(lstGrupos.SelectedItem, GrupoEmpresas)
        If g Is Nothing Then Return
        If _atual IsNot Nothing AndAlso Not String.Equals(_atual.Nome, g.Nome, StringComparison.OrdinalIgnoreCase) AndAlso Not ConfirmarDescartar() Then
            PreencherGrupos(_nomeAnterior)
            Return
        End If
        _atual = g.Clonar()
        _nomeAnterior = g.Nome
        txtNome.Text = g.Nome
        MostrarGrupo()
        btnExcluir.Enabled = True
    End Sub

    ''' <summary>Leva o grupo em edição para as duas listas.</summary>
    Private Sub MostrarGrupo()
        _noGrupo.Clear()
        _noGrupo.AddRange(_atual.Empresas(_clientes))
        _disponiveis.Clear()
        _disponiveis.AddRange(_clientes.Where(Function(c) c.Ativo AndAlso Not _noGrupo.Contains(c)))
        txtFiltro.Text = ""
        PreencherDisponiveis()
        PreencherNoGrupo()
    End Sub

    Private Sub PreencherDisponiveis()
        Dim termo = Colaborador.SemAcento(txtFiltro.Text.Trim())
        lstDisponiveis.BeginUpdate()
        lstDisponiveis.Items.Clear()
        For Each c In _disponiveis.OrderBy(Function(x) x.Nome)
            If termo = "" OrElse Colaborador.SemAcento(c.Rotulo).Contains(termo) Then lstDisponiveis.Items.Add(c)
        Next
        lstDisponiveis.EndUpdate()
    End Sub

    Private Sub PreencherNoGrupo()
        lstNoGrupo.BeginUpdate()
        lstNoGrupo.Items.Clear()
        For Each c In _noGrupo.OrderBy(Function(x) x.Nome)
            lstNoGrupo.Items.Add(c)
        Next
        lstNoGrupo.EndUpdate()
        Dim n = _noGrupo.Count
        lblContagem.Text = If(n = 0, "Nenhuma empresa no grupo.", n & If(n = 1, " empresa.", " empresas."))
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
        PreencherNoGrupo()
    End Sub

    ''' <summary>O que está na tela, como grupo.</summary>
    Private Function GrupoDaTela() As GrupoEmpresas
        Dim g As New GrupoEmpresas With {.UsuarioId = ModApontamento.UsuarioLogadoId(), .Nome = txtNome.Text.Trim()}
        g.Codigos.AddRange(_noGrupo.Select(Function(c) c.Tradutor))
        Return g
    End Function

    Private Function Sujo() As Boolean
        If _atual Is Nothing Then Return False
        Dim g = GrupoDaTela()
        If g.Nome <> _atual.Nome Then Return True
        Return Not g.Codigos.OrderBy(Function(x) x).SequenceEqual(_atual.Codigos.OrderBy(Function(x) x))
    End Function

    ''' <summary>Antes de trocar de grupo ou fechar: se há alteração não gravada, pergunta. True = pode seguir.</summary>
    Private Function ConfirmarDescartar() As Boolean
        If Not Sujo() Then Return True
        Dim nome = If(txtNome.Text.Trim() = "", "este grupo", """" & txtNome.Text.Trim() & """")
        Dim r = MessageBox.Show(Me, "Gravar as alterações em " & nome & "?", "Meus grupos", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If r = DialogResult.Cancel Then Return False
        If r = DialogResult.Yes Then Return Salvar()
        Return True
    End Function

    Private Sub AoFechar(sender As Object, e As FormClosingEventArgs)
        If Not ConfirmarDescartar() Then e.Cancel = True
    End Sub

    ' =================================================================
    '  GRAVAR / EXCLUIR
    ' =================================================================
    Private Function Salvar() As Boolean
        Dim g = GrupoDaTela()
        If g.Nome = "" Then
            MessageBox.Show(Me, "Dê um nome ao grupo.", "Meus grupos", MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtNome.Focus() : Return False
        End If
        If g.Codigos.Count = 0 Then
            MessageBox.Show(Me, "Passe pelo menos uma empresa para o grupo.", "Meus grupos", MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtFiltro.Focus() : Return False
        End If
        Dim repetido = _grupos.Any(Function(x) String.Equals(x.Nome, g.Nome, StringComparison.OrdinalIgnoreCase) AndAlso
                                                Not String.Equals(x.Nome, _nomeAnterior, StringComparison.OrdinalIgnoreCase))
        If repetido Then
            MessageBox.Show(Me, "Você já tem um grupo chamado """ & g.Nome & """.", "Meus grupos", MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtNome.Focus() : Return False
        End If
        Me.Cursor = Cursors.WaitCursor
        Try
            ModApontamento.SalvarGrupo(g, _nomeAnterior)
            _grupos = ModApontamento.Grupos()
        Catch ex As Exception
            MessageBox.Show(Me, "Não foi possível gravar o grupo." & vbCrLf & ex.Message, "Meus grupos", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Finally
            Me.Cursor = Cursors.Default
        End Try
        Alterou = True
        UltimoGravado = g.Nome
        _atual = g.Clonar()
        _nomeAnterior = g.Nome
        PreencherGrupos(g.Nome)
        btnExcluir.Enabled = True
        lblRodape.Text = "Grupo """ & g.Nome & """ gravado com " & g.Codigos.Count & If(g.Codigos.Count = 1, " empresa.", " empresas.")
        lblRodape.ForeColor = Color.FromArgb(22, 101, 52)
        Return True
    End Function

    Private Sub Excluir()
        If _nomeAnterior = "" Then Return
        If MessageBox.Show(Me, "Excluir o grupo """ & _nomeAnterior & """?", "Meus grupos", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return
        Me.Cursor = Cursors.WaitCursor
        Try
            ModApontamento.ExcluirGrupo(_nomeAnterior)
            _grupos = ModApontamento.Grupos()
        Catch ex As Exception
            MessageBox.Show(Me, "Não foi possível excluir o grupo." & vbCrLf & ex.Message, "Meus grupos", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        Finally
            Me.Cursor = Cursors.Default
        End Try
        Alterou = True
        Dim nome = _nomeAnterior
        PreencherGrupos(Nothing)
        Novo(Nothing)
        lblRodape.Text = "Grupo """ & nome & """ excluído."
        lblRodape.ForeColor = TemaMecano.CinzaEscuro
    End Sub

    ' ---------- só para captura e testes ----------
    ''' <summary>Monta um grupo novo na tela (nome e primeiras N empresas) sem gravar.</summary>
    Public Sub DemonstrarGrupo(nome As String, quantidade As Integer)
        Novo(Nothing)
        txtNome.Text = nome
        For Each c In _disponiveis.OrderBy(Function(x) x.Nome).Take(quantidade).ToList()
            _disponiveis.Remove(c)
            _noGrupo.Add(c)
        Next
        PreencherDisponiveis()
        PreencherNoGrupo()
    End Sub

    ''' <summary>Grava o que está na tela (testes).</summary>
    Public Function SalvarAgora() As Boolean
        Return Salvar()
    End Function

End Class

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
'  frmApontamentoPeriodo - "Anotações por período" (menu Apontamentos › Time).
'
'  Gestor ou diretoria escolhe um colaborador, um intervalo de datas e o motivo
'  (Férias / Afastamento ou Licença maternidade) e a tela grava um apontamento
'  por dia em nome dessa pessoa, com a regra do frmTimePeriodo de hoje:
'  domingo não grava; sábado grava "Sabado" com 4 h; dia útil grava o motivo
'  com 8 h. EMPRESAID vazio (atividade interna), via Time_Insert.
'
'  O que muda: a tela mostra ANTES quantos dias e horas vai gravar, pula os
'  dias em que a pessoa já tem apontamento (hoje duplicava) e só abre para
'  gestor e diretoria (hoje qualquer um gravava férias no nome de outro).
'
'  Destino no projeto: Grupo_Mecano\Cadastro\frmApontamentoPeriodo.vb
' ============================================================================
Public Class frmApontamentoPeriodo
    Inherits Form

    Private _escala As Single = 1.0F

    Private Function Esc(v As Integer) As Integer
        Return CInt(Math.Round(v * _escala))
    End Function

    Private _colaboradores As New List(Of ColaboradorResumo)()
    Private _dias As New List(Of DiaPeriodo)()
    Private _carregando As Boolean = True

    Public Property QuantidadeGravada As Integer
    Public Property MensagemResultado As String = ""

    ' ---------- controles ----------
    Private pnlTopo As Panel
    Private pnlCorpo As Panel
    Private pnlRodape As Panel
    Private secQuem As Panel
    Private secQuando As Panel
    Private cboColaborador As ComboBox
    Private chkDesligados As CheckBox
    Private ReadOnly rbMotivos As New List(Of RadioButton)()
    Private dtpDe As DateTimePicker
    Private dtpAte As DateTimePicker
    Private lblResumo As Label
    Private lblPulados As Label
    Private btnGravar As BotaoMecano
    Private btnCancelar As BotaoMecano
    Private lblRodape As Label

    Private Const LARGURA As Integer = 660
    Private Const SEC_X As Integer = 24
    Private Const SEC_W As Integer = 612
    Private Const ALT_TOPO As Integer = 88
    Private Const ALT_QUEM As Integer = 156
    Private Const ALT_QUANDO As Integer = 200
    Private Const ALT_RODAPE As Integer = 64

    Public Sub New()
        Using g = Me.CreateGraphics()
            _escala = g.DpiX / 96.0F
        End Using

        Me.AutoScaleMode = AutoScaleMode.None
        Me.Text = "Anotações por período"
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
        Me.ClientSize = New Size(Esc(LARGURA), Esc(ALT_TOPO) + Esc(20) + Esc(ALT_QUEM) + Esc(16) + Esc(ALT_QUANDO) + Esc(20) + Esc(ALT_RODAPE))

        AddHandler Me.Load, Sub() CarregarDados()
        AddHandler Me.KeyDown, Sub(s, e)
                                   If e.KeyCode = Keys.Escape Then
                                       Me.DialogResult = If(QuantidadeGravada > 0, DialogResult.OK, DialogResult.Cancel)
                                       Me.Close()
                                   ElseIf e.Control AndAlso (e.KeyCode = Keys.S OrElse e.KeyCode = Keys.Enter) Then
                                       e.Handled = True : e.SuppressKeyPress = True
                                       Gravar()
                                   End If
                               End Sub
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
                                      Dim flags = TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPrefix
                                      Using f = TemaMecano.FonteTitulo(15.0F)
                                          TextRenderer.DrawText(g, "Anotações por período", f, New Rectangle(Esc(24), Esc(18), pnlTopo.Width - Esc(48), Esc(30)), TemaMecano.TextoCampo, flags)
                                      End Using
                                      Using f = TemaMecano.FonteCorpo(9.25F)
                                          TextRenderer.DrawText(g, "Grava um dia de cada vez no nome da pessoa: " &
                                                                Apontamento.FormatarHoras(ResumoHoras.JORNADA) & " por dia útil, " &
                                                                Apontamento.FormatarHoras(BancoHoras.JORNADA_SABADO) & " no sábado, nada no domingo.", f,
                                                                New Rectangle(Esc(24), Esc(50), pnlTopo.Width - Esc(48), Esc(22)), TemaMecano.CinzaEscuro, flags)
                                      End Using
                                  End Sub
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

        btnGravar = BotaoMecano.Primario("Gravar", Esc(130), Esc(36))
        btnGravar.Top = Esc(14)
        btnGravar.Left = Esc(LARGURA) - Esc(24) - btnGravar.Width
        AddHandler btnGravar.Click, Sub() Gravar()

        btnCancelar = BotaoMecano.Secundario("Cancelar", Esc(110), Esc(36))
        btnCancelar.Top = Esc(14)
        btnCancelar.Left = btnGravar.Left - Esc(10) - btnCancelar.Width
        AddHandler btnCancelar.Click, Sub()
                                          Me.DialogResult = If(QuantidadeGravada > 0, DialogResult.OK, DialogResult.Cancel)
                                          Me.Close()
                                      End Sub

        lblRodape = New Label With {
            .AutoSize = False, .Left = Esc(24), .Top = Esc(14), .Width = btnCancelar.Left - Esc(24) - Esc(12), .Height = Esc(36), .AutoEllipsis = True,
            .TextAlign = ContentAlignment.MiddleLeft, .ForeColor = TemaMecano.CinzaClaro, .Font = TemaMecano.FonteCorpo(8.75F),
            .Text = "Ctrl+Enter grava  ·  Esc fecha"}

        pnlRodape.Controls.AddRange(New Control() {btnGravar, btnCancelar, lblRodape})
    End Sub

    ' =================================================================
    '  CORPO
    ' =================================================================
    Private Sub MontarCorpo()
        pnlCorpo = New Panel With {.Dock = DockStyle.Fill, .BackColor = TemaMecano.FundoTela}

        ' ---------------- para quem e por quê
        secQuem = NovaSecao("PARA QUEM E POR QUÊ", SEC_X, 20, ALT_QUEM)
        Rotulo(secQuem, "Colaborador", 16, 40)
        cboColaborador = New ComboBox With {
            .Left = Esc(16), .Top = Esc(60), .Width = Esc(400), .DropDownStyle = ComboBoxStyle.DropDownList,
            .Font = TemaMecano.FonteCorpo(10.0F), .FlatStyle = FlatStyle.System, .DropDownHeight = Esc(320)}
        AddHandler cboColaborador.SelectedIndexChanged, Sub() Recalcular()
        chkDesligados = New CheckBox With {.Text = "mostrar desligados", .AutoSize = True, .Left = Esc(430), .Top = Esc(63), .Font = TemaMecano.FonteCorpo(8.75F), .ForeColor = TemaMecano.CinzaEscuro}
        AddHandler chkDesligados.CheckedChanged, Sub() PreencherColaboradores()

        Rotulo(secQuem, "Motivo (é o texto que fica no apontamento)", 16, 98)
        Dim x = 16
        For Each motivo In PeriodoAfastamento.Motivos
            Dim rb = New RadioButton With {.Text = motivo, .AutoSize = True, .Left = Esc(x), .Top = Esc(118), .Font = TemaMecano.FonteCorpo(9.75F), .Checked = (rbMotivos.Count = 0)}
            AddHandler rb.CheckedChanged, Sub() Recalcular()
            rbMotivos.Add(rb)
            secQuem.Controls.Add(rb)
            x += 220
        Next
        secQuem.Controls.AddRange(New Control() {cboColaborador, chkDesligados})

        ' ---------------- período
        secQuando = NovaSecao("PERÍODO", SEC_X, 20 + ALT_QUEM + 16, ALT_QUANDO)
        Rotulo(secQuando, "De", 16, 40)
        dtpDe = New DateTimePicker With {
            .Left = Esc(16), .Top = Esc(60), .Width = Esc(170), .Format = DateTimePickerFormat.Custom, .CustomFormat = "ddd dd/MM/yyyy",
            .Value = Date.Today, .Font = TemaMecano.FonteCorpo(10.0F)}
        Rotulo(secQuando, "Até", 210, 40)
        dtpAte = New DateTimePicker With {
            .Left = Esc(210), .Top = Esc(60), .Width = Esc(170), .Format = DateTimePickerFormat.Custom, .CustomFormat = "ddd dd/MM/yyyy",
            .Value = Date.Today, .Font = TemaMecano.FonteCorpo(10.0F)}
        AddHandler dtpDe.ValueChanged, Sub()
                                           If dtpAte.Value.Date < dtpDe.Value.Date Then dtpAte.Value = dtpDe.Value.Date
                                           Recalcular()
                                       End Sub
        AddHandler dtpAte.ValueChanged, Sub() Recalcular()

        lblResumo = New Label With {
            .AutoSize = False, .Left = Esc(16), .Top = Esc(100), .Width = Esc(SEC_W - 32), .Height = Esc(24),
            .ForeColor = TemaMecano.TextoCampo, .Font = TemaMecano.FonteCorpo(10.0F, FontStyle.Bold), .Text = ""}
        lblPulados = New Label With {
            .AutoSize = False, .Left = Esc(16), .Top = Esc(128), .Width = Esc(SEC_W - 32), .Height = Esc(56),
            .ForeColor = TemaMecano.CinzaEscuro, .Font = TemaMecano.FonteCorpo(8.75F), .Text = ""}
        secQuando.Controls.AddRange(New Control() {dtpDe, dtpAte, lblResumo, lblPulados})

        pnlCorpo.Controls.Add(secQuem)
        pnlCorpo.Controls.Add(secQuando)
    End Sub

    Private Function NovaSecao(titulo As String, x As Integer, y As Integer, altura As Integer) As Panel
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
        If Not ModApontamento.PodeGerenciar() Then
            MessageBox.Show(Me, "Só gestor e diretoria lançam período para outra pessoa.", "Anotações por período", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.DialogResult = DialogResult.Cancel
            Me.Close()
            Return
        End If
        _carregando = True
        Me.Cursor = Cursors.WaitCursor
        Try
            _colaboradores = ModApontamento.Colaboradores()
        Catch ex As Exception
            MessageBox.Show(Me, "Não foi possível ler a lista de colaboradores." & vbCrLf & ex.Message, "Anotações por período", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
        PreencherColaboradores()
        _carregando = False
        Recalcular()
        cboColaborador.Focus()
    End Sub

    Private Sub PreencherColaboradores()
        Dim atual = ColaboradorEscolhido()
        cboColaborador.BeginUpdate()
        cboColaborador.Items.Clear()
        For Each c In _colaboradores
            If c.Desligado AndAlso Not chkDesligados.Checked Then Continue For
            cboColaborador.Items.Add(c)
        Next
        cboColaborador.EndUpdate()
        If atual IsNot Nothing AndAlso cboColaborador.Items.Contains(atual) Then cboColaborador.SelectedItem = atual
    End Sub

    Private Function ColaboradorEscolhido() As ColaboradorResumo
        Return TryCast(cboColaborador.SelectedItem, ColaboradorResumo)
    End Function

    Private Function MotivoEscolhido() As String
        Dim rb = rbMotivos.FirstOrDefault(Function(r) r.Checked)
        Return If(rb Is Nothing, PeriodoAfastamento.Motivos(0), rb.Text)
    End Function

    ''' <summary>Monta a lista de dias e mostra, antes de gravar, o que vai acontecer.</summary>
    Private Sub Recalcular()
        If _carregando OrElse lblResumo Is Nothing Then Return
        _dias.Clear()
        Dim c = ColaboradorEscolhido()
        Dim de = dtpDe.Value.Date
        Dim ate = dtpAte.Value.Date
        If c Is Nothing Then
            lblResumo.Text = "Escolha o colaborador."
            lblPulados.Text = ""
            Return
        End If
        If ate < de Then
            lblResumo.Text = "A data final é anterior à inicial."
            lblPulados.Text = ""
            Return
        End If

        ' o que a pessoa já tem em cada dia do intervalo (um acesso só)
        Dim porDia As New Dictionary(Of Date, Decimal)()
        Try
            For Each a In ModApontamento.ListarDe(c.Id, de, ate)
                Dim acum As Decimal = 0D
                porDia.TryGetValue(a.Data.Date, acum)
                porDia(a.Data.Date) = acum + a.Horas
            Next
        Catch
        End Try
        _dias = PeriodoAfastamento.Montar(de, ate, MotivoEscolhido(), Function(d)
                                                                           Dim v As Decimal = 0D
                                                                           porDia.TryGetValue(d, v)
                                                                           Return v
                                                                       End Function)
        lblResumo.Text = PeriodoAfastamento.Descrever(_dias) & "  ·  domingos não entram"

        Dim pulados = _dias.Where(Function(d) d.Pular).ToList()
        If pulados.Count = 0 Then
            lblPulados.Text = "Nenhum desses dias tem apontamento ainda."
        Else
            Dim datas = String.Join(", ", pulados.Take(8).Select(Function(d) d.Data.ToString("dd/MM")))
            If pulados.Count > 8 Then datas &= " …"
            lblPulados.Text = pulados.Count & If(pulados.Count = 1, " dia já tem apontamento e não será gravado: ", " dias já têm apontamento e não serão gravados: ") & datas & "."
        End If
    End Sub

    ''' <summary>Só para a captura de tela da demonstração: escolhe a segunda pessoa da lista e um intervalo de três semanas a partir da próxima segunda.</summary>
    Public Sub DemonstrarPeriodo()
        If cboColaborador.Items.Count > 1 Then cboColaborador.SelectedIndex = 1 Else If cboColaborador.Items.Count > 0 Then cboColaborador.SelectedIndex = 0
        Dim seg = Date.Today.AddDays(1)
        While seg.DayOfWeek <> DayOfWeek.Monday
            seg = seg.AddDays(1)
        End While
        dtpDe.Value = seg
        dtpAte.Value = seg.AddDays(18)
        Recalcular()
    End Sub

    ' =================================================================
    '  GRAVAR
    ' =================================================================
    Private Sub Gravar()
        Dim c = ColaboradorEscolhido()
        If c Is Nothing Then
            MessageBox.Show(Me, "Escolha o colaborador.", "Anotações por período", MessageBoxButtons.OK, MessageBoxIcon.Information)
            cboColaborador.Focus() : Return
        End If
        If dtpAte.Value.Date < dtpDe.Value.Date Then
            MessageBox.Show(Me, "A data final é anterior à inicial.", "Anotações por período", MessageBoxButtons.OK, MessageBoxIcon.Information)
            dtpAte.Focus() : Return
        End If
        Recalcular()
        Dim gravar = _dias.Where(Function(d) Not d.Pular).ToList()
        If gravar.Count = 0 Then
            MessageBox.Show(Me, "Não há nenhum dia para gravar nesse intervalo: ou são domingos, ou a pessoa já tem apontamento em todos.", "Anotações por período", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim pergunta = "Gravar " & gravar.Count & If(gravar.Count = 1, " lançamento", " lançamentos") & " de """ & MotivoEscolhido() & """ para " & c.Nome &
                       ", de " & dtpDe.Value.ToString("dd/MM/yyyy") & " a " & dtpAte.Value.ToString("dd/MM/yyyy") & "?" & vbCrLf & vbCrLf &
                       PeriodoAfastamento.Descrever(_dias)
        If MessageBox.Show(Me, pergunta, "Anotações por período", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then Return

        Me.Cursor = Cursors.WaitCursor
        Dim n = 0
        Try
            For Each d In gravar
                ModApontamento.Inserir(New Apontamento With {.Data = d.Data, .Empresa = d.Texto, .EmpresaId = "", .Horas = d.Horas, .UsuarioId = c.Id})
                n += 1
            Next
            ModApontamento.AposGravar()
        Catch ex As Exception
            MessageBox.Show(Me, "Não foi possível gravar tudo: " & n & " de " & gravar.Count & " dias gravados." & vbCrLf & ex.Message, "Anotações por período", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
        If n = 0 Then Return

        QuantidadeGravada += n
        MensagemResultado = n & If(n = 1, " dia gravado", " dias gravados") & " para " & c.Nome & " (" & MotivoEscolhido() & ", " &
                            Apontamento.FormatarHoras(gravar.Take(n).Sum(Function(d) d.Horas)) & ")."
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

End Class

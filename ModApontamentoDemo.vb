Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms

' ============================================================================
'  ModApontamento (VERSÃO DE DEMONSTRAÇÃO) - mesma interface do
'  ModApontamento.vb real, mas lê e grava em arquivos texto ao lado do
'  executável, sem banco nem rede:
'    clientes.txt      empresas FICTÍCIAS (código, nome, CNPJ, grupo)
'    outros.txt        cópia do Outros.txt da rede (atividades internas)
'    horasmes.txt      metas mensais reais da TbHorasMes (2025 e 2026)
'    apontamentos.txt  lançamentos da demo; se não existir, é GERADO na
'                      primeira abertura, em volta da data de hoje, com dois
'                      dias pendentes de propósito. Apague para recomeçar.
'
'  NÃO vai para o projeto Grupo_Mecano: lá entra o ModApontamento.vb (SQL).
' ============================================================================
Public Module ModApontamento

    Public ReadOnly ModoDemo As Boolean = True

    ''' <summary>Pasta dos dados: ao lado do exe ou, se o exe estiver em outra pasta (demo da tela principal), a TO-DO_Entrega-Apontamentos mais próxima acima.</summary>
    Public Function PastaDados() As String
        Dim p = Application.StartupPath
        If File.Exists(Path.Combine(p, "clientes.txt")) Then Return p
        Dim d = New DirectoryInfo(p)
        While d IsNot Nothing
            Dim cand = Path.Combine(d.FullName, "TO-DO_Entrega-Apontamentos")
            If File.Exists(Path.Combine(cand, "clientes.txt")) Then Return cand
            d = d.Parent
        End While
        Return p
    End Function

    Private ReadOnly _arqApontamentos As String = Path.Combine(PastaDados(), "apontamentos.txt")
    Private ReadOnly _arqClientes As String = Path.Combine(PastaDados(), "clientes.txt")
    Private ReadOnly _arqOutros As String = Path.Combine(PastaDados(), "outros.txt")
    Private ReadOnly _arqHorasMes As String = Path.Combine(PastaDados(), "horasmes.txt")

    Private _cache As List(Of Apontamento) = Nothing
    Private _clientes As List(Of ClienteResumo) = Nothing
    Private _metas As Dictionary(Of String, Decimal) = Nothing

    ' ------------------------------------------------------------ sessão
    Public Function UsuarioLogadoId() As Integer
        Return Usersys
    End Function

    Public Function UsuarioLogadoNome() As String
        Return If(UserName, "")
    End Function

    Public Function Admissao() As Date
        Return UserAdmissao.Date
    End Function

    Public Function VeAtividadesAdministrativas() As Boolean
        Dim funcao = If(UserFuncao, "").ToUpperInvariant().Trim()
        Return funcao.StartsWith("GESTOR") OrElse funcao.StartsWith("DIRETORIA")
    End Function

    Public Function DescricaoFonteDados() As String
        Return "DEMONSTRAÇÃO · dados fictícios em apontamentos.txt, sem banco"
    End Function

    ' ------------------------------------------------------------ apontamentos
    Public Function Listar(de As Date, ate As Date) As List(Of Apontamento)
        Carregar()
        Return _cache.Where(Function(a) a.Data.Date >= de.Date AndAlso a.Data.Date <= ate.Date).Select(Function(a) a.Clonar()).ToList()
    End Function

    Public Function ListarDesde(de As Date) As List(Of Apontamento)
        Return Listar(de, Date.Today)
    End Function

    ''' <summary>Banco de horas da pessoa, refeito a partir dos apontamentos desde BancoHoras.INICIO_VIGENCIA (mesma conta da versão SQL).</summary>
    Public Function CalcularBanco() As BancoHoras
        Return BancoHoras.Calcular(ListarDesde(BancoHoras.INICIO_VIGENCIA), BancoHoras.INICIO_VIGENCIA, Date.Today)
    End Function

    ''' <summary>Horas que a pessoa já tem numa data, sem contar o apontamento que está sendo alterado (trava de 10 h por dia).</summary>
    Public Function HorasNoDia(data As Date, excetoId As Integer) As Decimal
        Return Listar(data, data).Where(Function(a) a.Id <> excetoId).Sum(Function(a) a.Horas)
    End Function

    Public Function Obter(id As Integer) As Apontamento
        Carregar()
        Dim a = _cache.FirstOrDefault(Function(x) x.Id = id)
        Return If(a Is Nothing, Nothing, a.Clonar())
    End Function

    Public Function Inserir(a As Apontamento) As Integer
        Carregar()
        a.Id = If(_cache.Count = 0, 1, _cache.Max(Function(x) x.Id) + 1)
        If a.UsuarioId = 0 Then a.UsuarioId = Usersys
        _cache.Add(a.Clonar())
        Salvar()
        Return a.Id
    End Function

    Public Sub Alterar(a As Apontamento)
        Carregar()
        Dim i = _cache.FindIndex(Function(x) x.Id = a.Id)
        If i < 0 Then Throw New InvalidOperationException("Apontamento " & a.Id & " não existe mais.")
        _cache(i) = a.Clonar()
        Salvar()
    End Sub

    Public Sub Excluir(id As Integer)
        Carregar()
        _cache.RemoveAll(Function(x) x.Id = id)
        Salvar()
    End Sub

    ''' <summary>Na demo não há menu para liberar.</summary>
    Public Sub AposGravar()
    End Sub

    ' ------------------------------------------------------------ clientes e atividades
    Public Function Clientes() As List(Of ClienteResumo)
        If _clientes IsNot Nothing Then Return _clientes
        Dim lista As New List(Of ClienteResumo)()
        If File.Exists(_arqClientes) Then
            For Each linha In File.ReadAllLines(_arqClientes, Encoding.UTF8).Skip(1)
                Dim p = linha.Split("|"c)
                If p.Length < 6 OrElse p(0).Trim() = "" Then Continue For
                Dim grupo As Integer
                Integer.TryParse(p(3), grupo)
                lista.Add(New ClienteResumo With {.Tradutor = p(0).Trim(), .Nome = p(1).Trim(), .Cnpj = p(2).Trim(),
                                                 .CodGrupo = grupo, .Grupo = p(4).Trim(), .Ativo = (p(5).Trim() = "1")})
            Next
        End If
        _clientes = lista.OrderByDescending(Function(c) c.Ativo).ThenBy(Function(c) c.Nome).ToList()
        Return _clientes
    End Function

    Public Sub LimparCache()
        _clientes = Nothing
        _metas = Nothing
        _cache = Nothing
    End Sub

    Public Function AtividadesInternas() As List(Of String)
        Dim linhas As New List(Of String)()
        If File.Exists(_arqOutros) Then
            linhas = File.ReadAllLines(_arqOutros, Encoding.UTF8).Select(Function(l) l.Trim()).Where(Function(l) l <> "").Distinct().ToList()
        End If
        If Not VeAtividadesAdministrativas() Then
            linhas = linhas.Where(Function(l) Not l.ToLowerInvariant().Contains("administrativo")).ToList()
        End If
        Return linhas
    End Function

    Public Function ArquivoAtividadesDisponivel() As Boolean
        Return File.Exists(_arqOutros)
    End Function

    ' ------------------------------------------------------------ metas
    Public Function MetaMes(competencia As String) As Decimal
        If _metas Is Nothing Then
            _metas = New Dictionary(Of String, Decimal)()
            If File.Exists(_arqHorasMes) Then
                For Each linha In File.ReadAllLines(_arqHorasMes, Encoding.UTF8).Skip(1)
                    Dim p = linha.Split("|"c)
                    Dim v As Decimal
                    If p.Length >= 2 AndAlso Decimal.TryParse(p(1), NumberStyles.Any, CultureInfo.InvariantCulture, v) Then _metas(p(0).Trim()) = v
                Next
            End If
        End If
        Dim meta As Decimal = 0D
        If Not _metas.TryGetValue(competencia, meta) Then
            ' competência fora da tabela (demo aberta em outro ano): dias úteis x 8 h
            Dim d As Date
            If Date.TryParseExact("01/" & competencia, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, d) Then
                meta = DiasUteis(d) * ResumoHoras.JORNADA
            End If
        End If
        Return meta
    End Function

    Public Function CalcularResumo() As ResumoHoras
        Dim hoje = Date.Today
        Return ResumoHoras.Calcular(ListarDesde(ResumoHoras.InicioNecessario(hoje)), AddressOf MetaMes, Admissao(), hoje)
    End Function

    ' ------------------------------------------------------------ arquivo
    Private Sub Carregar()
        If _cache IsNot Nothing Then Return
        _cache = New List(Of Apontamento)()
        If Not File.Exists(_arqApontamentos) Then
            GerarDemonstracao()
            Salvar()
            Return
        End If
        For Each linha In File.ReadAllLines(_arqApontamentos, Encoding.UTF8).Skip(1)
            Dim p = linha.Split("|"c)
            If p.Length < 5 Then Continue For
            Dim id As Integer, d As Date, h As Decimal
            If Not Integer.TryParse(p(0), id) Then Continue For
            If Not Date.TryParseExact(p(1), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, d) Then Continue For
            Decimal.TryParse(p(4), NumberStyles.Any, CultureInfo.InvariantCulture, h)
            _cache.Add(New Apontamento With {.Id = id, .Data = d, .Empresa = p(2), .EmpresaId = p(3), .Horas = h, .UsuarioId = Usersys})
        Next
    End Sub

    Private Sub Salvar()
        Dim sb As New StringBuilder()
        sb.AppendLine("ID|DATA|EMPRESA|EMPRESAID|HORAS")
        For Each a In _cache.OrderBy(Function(x) x.Data).ThenBy(Function(x) x.Id)
            sb.AppendLine(String.Join("|", New String() {
                a.Id.ToString(), a.Data.ToString("yyyy-MM-dd"), a.Empresa.Replace("|", " "), a.EmpresaId,
                a.Horas.ToString("0.00", CultureInfo.InvariantCulture)}))
        Next
        File.WriteAllText(_arqApontamentos, sb.ToString(), New UTF8Encoding(False))
    End Sub

    Private Function DiasUteis(mes As Date) As Integer
        Dim n = 0
        Dim d = New Date(mes.Year, mes.Month, 1)
        While d.Month = mes.Month
            If d.DayOfWeek <> DayOfWeek.Saturday AndAlso d.DayOfWeek <> DayOfWeek.Sunday Then n += 1
            d = d.AddDays(1)
        End While
        Return n
    End Function

    ''' <summary>
    ''' Lançamentos de demonstração em volta de HOJE: mês anterior fechado na meta
    ''' (com sábados de 4 h, como o frmTimePeriodo grava), mês atual até ontem com
    ''' dois dias pendentes (um com 6h30, um sem nada), dois dias acima de 8 h (9h30 e
    ''' 10h00, para o banco de horas) e hoje pela metade.
    ''' </summary>
    Private Sub GerarDemonstracao()
        Dim rnd As New Random(2026)
        Dim ativos As List(Of ClienteResumo) = Clientes().Where(Function(c) c.Ativo).ToList()
        Dim atividades = AtividadesInternas()
        If ativos.Count = 0 Then Return
        Dim internasComuns = {"Estudos Tecnicos / treinamento / orientacao / pesquisas", "Arquivo / organizacao", "Problemas com Sistemas / computador / falta de energia"}.
            Where(Function(t) atividades.Contains(t)).ToList()
        Dim preferidos = ativos.Take(14).ToList()   ' carteira "da pessoa": os mesmos clientes voltam todo dia

        Dim hoje = Date.Today
        Dim inicio = New Date(hoje.Year, hoje.Month, 1).AddMonths(-2)
        Dim pendenteParcial = DiaUtilAntes(hoje, 4)     ' 6h30
        Dim pendenteVazio = DiaUtilAntes(hoje, 9)       ' sem lançamento
        Dim extraMeio = DiaUtilAntes(hoje, 2)           ' 9h30: 1h30 para o banco de horas
        Dim extraTeto = DiaUtilAntes(hoje, 6)           ' 10h00, o teto do dia: 2h00 para o banco
        Dim proximoId = 1

        Dim d = inicio
        While d <= hoje
            If d.DayOfWeek <> DayOfWeek.Saturday AndAlso d.DayOfWeek <> DayOfWeek.Sunday Then
                Dim alvoMin As Integer = 480
                If d = hoje Then alvoMin = 210
                If d = pendenteParcial Then alvoMin = 390
                If d = pendenteVazio Then alvoMin = 0
                If d = extraMeio Then alvoMin = 570
                If d = extraTeto Then alvoMin = 600
                Dim restante = alvoMin
                Dim linhas = 0
                While restante > 0
                    Dim minutos As Integer
                    If restante <= 60 Then
                        minutos = restante
                    Else
                        Dim opcoes = {30, 60, 60, 90, 120, 120, 180, 240}
                        minutos = Math.Min(restante, opcoes(rnd.Next(opcoes.Length)))
                    End If
                    Dim a As New Apontamento With {.Id = proximoId, .Data = d, .Horas = Apontamento.DeMinutos(minutos), .UsuarioId = Usersys}
                    If internasComuns.Count > 0 AndAlso rnd.Next(100) < 12 Then
                        a.Empresa = internasComuns(rnd.Next(internasComuns.Count))
                        a.EmpresaId = ""
                    Else
                        Dim c = If(rnd.Next(100) < 75, preferidos(rnd.Next(preferidos.Count)), ativos(rnd.Next(ativos.Count)))
                        a.Empresa = c.Nome
                        a.EmpresaId = c.Tradutor
                    End If
                    _cache.Add(a)
                    proximoId += 1
                    restante -= minutos
                    linhas += 1
                    If linhas >= 6 AndAlso restante > 0 Then
                        _cache.Last().Horas += Apontamento.DeMinutos(restante)
                        restante = 0
                    End If
                End While
            End If
            d = d.AddDays(1)
        End While

        ' Mês anterior: completa a meta com sábados de 4 h ("Sabado", como o Time grava) e, se sobrar, no último dia útil.
        Dim mesAnterior = New Date(hoje.Year, hoje.Month, 1).AddMonths(-1)
        Dim meta = MetaMes(Apontamento.Competencia(mesAnterior))
        Dim apontado = _cache.Where(Function(a) a.Data.Year = mesAnterior.Year AndAlso a.Data.Month = mesAnterior.Month).Sum(Function(a) a.Horas)
        Dim falta = meta - apontado
        Dim sab = mesAnterior
        While falta >= 4D AndAlso sab.Month = mesAnterior.Month
            If sab.DayOfWeek = DayOfWeek.Saturday Then
                _cache.Add(New Apontamento With {.Id = proximoId, .Data = sab, .Empresa = "Sabado", .EmpresaId = "", .Horas = 4D, .UsuarioId = Usersys})
                proximoId += 1
                falta -= 4D
            End If
            sab = sab.AddDays(1)
        End While
        If falta > 0D Then
            Dim ultimo = _cache.Where(Function(a) a.Data.Year = mesAnterior.Year AndAlso a.Data.Month = mesAnterior.Month AndAlso a.EhCliente).
                OrderByDescending(Function(a) a.Data).FirstOrDefault()
            If ultimo IsNot Nothing Then ultimo.Horas += falta
        End If
    End Sub

    Private Function DiaUtilAntes(hoje As Date, diasUteis As Integer) As Date
        Dim d = hoje
        Dim n = 0
        While n < diasUteis
            d = d.AddDays(-1)
            If d.DayOfWeek <> DayOfWeek.Saturday AndAlso d.DayOfWeek <> DayOfWeek.Sunday Then n += 1
        End While
        Return d
    End Function

End Module

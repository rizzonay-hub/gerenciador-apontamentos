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
        Return ListarDe(Usersys, de, ate)
    End Function

    Public Function ListarDe(usuarioId As Integer, de As Date, ate As Date) As List(Of Apontamento)
        Carregar()
        Return _cache.Where(Function(a) a.UsuarioId = usuarioId AndAlso a.Data.Date >= de.Date AndAlso a.Data.Date <= ate.Date).Select(Function(a) a.Clonar()).ToList()
    End Function

    ' ------------------------------------------------------------ gestão (gestor e diretoria)
    Public Function PodeGerenciar() As Boolean
        Return VeAtividadesAdministrativas()
    End Function

    Private _colaboradores As List(Of ColaboradorResumo) = Nothing
    Private ReadOnly _arqColaboradores As String = Path.Combine(PastaDados(), "colaboradores.txt")

    ''' <summary>Colaboradores FICTÍCIOS do colaboradores.txt desta pasta (id|nome|depto|funcao|desligado).</summary>
    Public Function Colaboradores() As List(Of ColaboradorResumo)
        If _colaboradores IsNot Nothing Then Return _colaboradores
        Dim lista As New List(Of ColaboradorResumo)()
        If File.Exists(_arqColaboradores) Then
            For Each linha In File.ReadAllLines(_arqColaboradores, Encoding.UTF8)
                If linha.Trim() = "" OrElse linha.StartsWith("#") Then Continue For
                Dim p = linha.Split("|"c)
                If p.Length < 3 Then Continue For
                Dim id As Integer
                If Not Integer.TryParse(p(0), id) Then Continue For
                lista.Add(New ColaboradorResumo With {.Id = id, .Nome = p(1).Trim(), .Depto = p(2).Trim(), .Funcao = If(p.Length >= 4, p(3).Trim(), ""), .Desligado = (p.Length >= 5 AndAlso p(4).Trim() = "1")})
            Next
        End If
        If Not lista.Any(Function(c) c.Id = Usersys) Then lista.Insert(0, New ColaboradorResumo With {.Id = Usersys, .Nome = UsuarioLogadoNome(), .Depto = "FISCAL"})
        _colaboradores = lista.OrderBy(Function(c) c.Desligado).ThenBy(Function(c) c.Nome).ToList()
        Return _colaboradores
    End Function

    ''' <summary>Mesma seleção que as procedures Time_GetByEmpresa* fazem, sobre o txt.</summary>
    Public Function ConsultaAnalitica(f As FiltroAnalitica) As List(Of LinhaAnalitica)
        Carregar()
        Dim pessoas = Colaboradores()
        Dim usuarioId As Integer = -1
        If Not f.TodosUsuarios Then
            Dim c = pessoas.FirstOrDefault(Function(x) x.Nome = f.Usuario)
            usuarioId = If(c Is Nothing, 0, c.Id)
        End If
        Dim itens = _cache.Where(Function(a) a.Data.Date >= f.De.Date AndAlso a.Data.Date <= f.Ate.Date AndAlso (usuarioId = -1 OrElse a.UsuarioId = usuarioId))
        Select Case f.Modo
            Case ModoAnalitica.ComCodigoEmpresa
                itens = itens.Where(Function(a) a.EhCliente AndAlso (f.TodasEmpresas OrElse a.EmpresaId.TrimStart("0"c) = f.EmpresaId.TrimStart("0"c)))
            Case ModoAnalitica.SemCodigoEmpresa
                itens = itens.Where(Function(a) Not a.EhCliente AndAlso (f.TodasEmpresas OrElse a.Empresa = f.Empresa))
            Case Else
                itens = itens.Where(Function(a) f.TodasEmpresas OrElse a.Empresa = f.Empresa)
        End Select
        Return itens.Select(Function(a)
                                Dim c = pessoas.FirstOrDefault(Function(x) x.Id = a.UsuarioId)
                                Return New LinhaAnalitica With {
                                    .Id = a.Id, .Data = a.Data, .Empresa = a.Empresa, .EmpresaId = a.EmpresaId, .Horas = a.Horas,
                                    .UsuarioId = a.UsuarioId, .Usuario = If(c Is Nothing, "usuário " & a.UsuarioId, c.Nome), .Depto = If(c Is Nothing, "", c.Depto)}
                            End Function).
            OrderByDescending(Function(l) l.Data).ThenBy(Function(l) l.Usuario).ThenBy(Function(l) l.Empresa).ToList()
    End Function

    ''' <summary>Mesma seleção das procedures Time_GetByEmpresaAcumulado*, sobre o txt: soma por empresa, colaborador e mês.</summary>
    Public Function ConsultaSintetica(f As FiltroSintetica) As List(Of LinhaSintetica)
        Carregar()
        Dim pessoas = Colaboradores()
        Dim clientes = Me_Clientes()
        Dim usuarioId As Integer = -1
        If Not f.TodosUsuarios Then
            Dim c = pessoas.FirstOrDefault(Function(x) x.Nome = f.Usuario)
            usuarioId = If(c Is Nothing, 0, c.Id)
        End If
        Dim itens = _cache.Where(Function(a) a.Data.Date >= f.De.Date AndAlso a.Data.Date <= f.Ate.Date AndAlso (usuarioId = -1 OrElse a.UsuarioId = usuarioId))
        If Not f.TodosGrupos Then
            ' grupo é atributo do cliente: com um grupo escolhido só entram linhas com código, do grupo (Time_GetByEmpresaAcumuladoGeral)
            Dim doGrupo = New HashSet(Of String)(clientes.Where(Function(c) c.Grupo = f.Grupo).Select(Function(c) c.Tradutor.TrimStart("0"c)))
            itens = itens.Where(Function(a) a.EhCliente AndAlso doGrupo.Contains(a.EmpresaId.TrimStart("0"c)))
            If Not f.TodasEmpresas Then itens = itens.Where(Function(a) a.Empresa = f.Empresa)
        Else
            Select Case f.Modo
                Case ModoAnalitica.ComCodigoEmpresa
                    itens = itens.Where(Function(a) a.EhCliente AndAlso (f.TodasEmpresas OrElse a.EmpresaId.TrimStart("0"c) = f.EmpresaId.TrimStart("0"c)))
                Case ModoAnalitica.SemCodigoEmpresa
                    itens = itens.Where(Function(a) Not a.EhCliente AndAlso (f.TodasEmpresas OrElse a.Empresa = f.Empresa))
                Case Else
                    itens = itens.Where(Function(a) f.TodasEmpresas OrElse a.Empresa = f.Empresa)
            End Select
        End If
        Return itens.GroupBy(Function(a) a.EmpresaId.TrimStart("0"c) & "|" & a.Empresa.ToUpperInvariant() & "|" & a.UsuarioId & "|" & a.Data.Year & "|" & a.Data.Month).
            Select(Function(gr)
                       Dim p = gr.First()
                       Dim c = pessoas.FirstOrDefault(Function(x) x.Id = p.UsuarioId)
                       Dim cli = If(p.EhCliente, clientes.FirstOrDefault(Function(x) x.Tradutor.TrimStart("0"c) = p.EmpresaId.TrimStart("0"c)), Nothing)
                       Return New LinhaSintetica With {
                           .Codigo = p.EmpresaId, .Cnpj = If(cli Is Nothing, "", cli.Cnpj), .Empresa = p.Empresa, .Grupo = If(cli Is Nothing, "", cli.Grupo),
                           .UsuarioId = p.UsuarioId, .Usuario = If(c Is Nothing, "usuário " & p.UsuarioId, c.Nome),
                           .Ano = p.Data.Year, .Mes = p.Data.Month, .Horas = gr.Sum(Function(a) a.Horas)}
                   End Function).
            OrderBy(Function(l) l.Ano).ThenBy(Function(l) l.Mes).ThenBy(Function(l) l.Empresa).ThenBy(Function(l) l.Usuario).ToList()
    End Function

    Private Function Me_Clientes() As List(Of ClienteResumo)
        Return Clientes()
    End Function

    ''' <summary>Horas diretas (clientes) e indiretas (atividades internas) por pessoa, como Time_GetByDataTudoHorasCustoDireto/Indireto.</summary>
    Public Function ConsultaPorFuncionario(de As Date, ate As Date) As List(Of LinhaFuncionario)
        Carregar()
        Dim pessoas = Colaboradores()
        Return _cache.Where(Function(a) a.Data.Date >= de.Date AndAlso a.Data.Date <= ate.Date).
            GroupBy(Function(a) a.UsuarioId).
            Select(Function(gr)
                       Dim c = pessoas.FirstOrDefault(Function(x) x.Id = gr.Key)
                       Return New LinhaFuncionario With {
                           .UsuarioId = gr.Key, .Usuario = If(c Is Nothing, "usuário " & gr.Key, c.Nome),
                           .Depto = If(c Is Nothing, "", c.Depto), .Funcao = If(c Is Nothing, "", c.Funcao),
                           .HorasDiretas = Math.Round(gr.Where(Function(a) a.EhCliente).Sum(Function(a) a.Horas), 2),
                           .HorasIndiretas = Math.Round(gr.Where(Function(a) Not a.EhCliente).Sum(Function(a) a.Horas), 2)}
                   End Function).
            Where(Function(l) l.Total > 0D).OrderBy(Function(l) l.Usuario).ToList()
    End Function

    ''' <summary>Total de horas por pessoa no intervalo (Time_GetByEmpresaUsuario), com zero para quem não lançou.</summary>
    Public Function ConsultaPorUsuarioData(de As Date, ate As Date) As List(Of LinhaUsuarioData)
        Carregar()
        Dim esperado = LinhaUsuarioData.HorasEsperadas(de, ate)
        Dim pessoas = Colaboradores()
        Dim somas = _cache.Where(Function(a) a.Data.Date >= de.Date AndAlso a.Data.Date <= ate.Date).
            GroupBy(Function(a) a.UsuarioId).ToDictionary(Function(gr) gr.Key, Function(gr) gr.Sum(Function(a) a.Horas))
        Dim lista As New List(Of LinhaUsuarioData)()
        For Each c In pessoas
            Dim h As Decimal = 0D
            Dim tem = somas.TryGetValue(c.Id, h)
            If tem OrElse Not c.Desligado Then
                lista.Add(New LinhaUsuarioData With {.UsuarioId = c.Id, .Usuario = c.Nome, .Depto = c.Depto, .Horas = h, .Esperado = esperado})
            End If
        Next
        Return lista.OrderBy(Function(l) l.Usuario).ToList()
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
        Return HorasNoDiaDe(Usersys, data, excetoId)
    End Function

    '''<summary>O mesmo para outra pessoa: o gestor lançando em nome dela vê a trava de 10 h do dia dela.</summary>
    Public Function HorasNoDiaDe(usuarioId As Integer, data As Date, excetoId As Integer) As Decimal
        Return ListarDe(usuarioId, data, data).Where(Function(a) a.Id <> excetoId).Sum(Function(a) a.Horas)
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
            Dim uid = Usersys                       ' arquivos antigos (sem a coluna) são da pessoa logada
            If p.Length >= 6 Then Integer.TryParse(p(5), uid)
            _cache.Add(New Apontamento With {.Id = id, .Data = d, .Empresa = p(2), .EmpresaId = p(3), .Horas = h, .UsuarioId = uid})
        Next
    End Sub

    Private Sub Salvar()
        Dim sb As New StringBuilder()
        sb.AppendLine("ID|DATA|EMPRESA|EMPRESAID|HORAS|USUARIOID")
        For Each a In _cache.OrderBy(Function(x) x.Data).ThenBy(Function(x) x.Id)
            sb.AppendLine(String.Join("|", New String() {
                a.Id.ToString(), a.Data.ToString("yyyy-MM-dd"), a.Empresa.Replace("|", " "), a.EmpresaId,
                a.Horas.ToString("0.00", CultureInfo.InvariantCulture), a.UsuarioId.ToString()}))
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
                GerarDia(Usersys, d, alvoMin, rnd, preferidos, ativos, internasComuns, proximoId)
            End If
            d = d.AddDays(1)
        End While

        ' Mês anterior: completa a meta com sábados de 4 h ("Sabado", como o Time grava) e, se sobrar, no último dia útil.
        Dim mesAnterior = New Date(hoje.Year, hoje.Month, 1).AddMonths(-1)
        Dim meta = MetaMes(Apontamento.Competencia(mesAnterior))
        Dim apontado = _cache.Where(Function(a) a.UsuarioId = Usersys AndAlso a.Data.Year = mesAnterior.Year AndAlso a.Data.Month = mesAnterior.Month).Sum(Function(a) a.Horas)
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
            Dim ultimo = _cache.Where(Function(a) a.UsuarioId = Usersys AndAlso a.Data.Year = mesAnterior.Year AndAlso a.Data.Month = mesAnterior.Month AndAlso a.EhCliente).
                OrderByDescending(Function(a) a.Data).FirstOrDefault()
            If ultimo IsNot Nothing Then ultimo.Horas += falta
        End If

        ' Mais duas pessoas (colaboradores.txt, fictícias), mês anterior e atual, 8 h por dia útil:
        ' dão conteúdo ao Gerenciamento (consulta de toda a equipe) sem mexer nos cartões da pessoa logada.
        For Each outro In Colaboradores().Where(Function(c) c.Id <> Usersys AndAlso Not c.Desligado).Take(2).ToList()
            Dim carteira = ativos.Skip(14 + 10 * (outro.Id Mod 7)).Take(12).ToList()
            If carteira.Count = 0 Then carteira = preferidos
            Dim d2 = mesAnterior
            While d2 <= hoje
                If d2.DayOfWeek <> DayOfWeek.Saturday AndAlso d2.DayOfWeek <> DayOfWeek.Sunday Then
                    GerarDia(outro.Id, d2, If(d2 = hoje, 240, 480), rnd, carteira, ativos, internasComuns, proximoId)
                End If
                d2 = d2.AddDays(1)
            End While
        Next
    End Sub

    ''' <summary>Um dia de uma pessoa: linhas de 30 min a 4 h até fechar o alvo, 12 % em atividades internas, o resto em clientes (75 % da carteira).</summary>
    Private Sub GerarDia(usuarioId As Integer, d As Date, alvoMin As Integer, rnd As Random, preferidos As List(Of ClienteResumo),
                         ativos As List(Of ClienteResumo), internasComuns As List(Of String), ByRef proximoId As Integer)
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
            Dim a As New Apontamento With {.Id = proximoId, .Data = d, .Horas = Apontamento.DeMinutos(minutos), .UsuarioId = usuarioId}
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

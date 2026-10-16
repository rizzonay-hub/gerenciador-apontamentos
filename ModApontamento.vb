Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Linq

' ============================================================================
'  ModApontamento - acesso a dados dos APONTAMENTOS DE HORAS (o "Time").
'
'  Usa as MESMAS procedures que o frmTime/frmTimeDet usam hoje (Time_GetByData,
'  Time_Get, Time_Insert, Time_Update, Time_Delete, HorasMes_GetList) e a
'  TbClientes para a lista de empresas, então nada muda no banco: a tela nova
'  grava exatamente o que a antiga gravava (EMPRESA = nome do cliente ou texto
'  da atividade; EMPRESAID = TRADUTOR ou vazio; TEMPO em horas com 2 casas).
'
'  As atividades internas continuam vindo do arquivo Outros.txt da rede, o
'  mesmo que o frmTimeDet lê (I:\Sistemas\Nibo\Clientes\Outros.txt).
'
'  Destino no projeto: Grupo_Mecano\App_Code\ModApontamento.vb
'  (não incluir ModApontamentoDemo.vb no projeto: é só da demonstração)
' ============================================================================
Public Module ModApontamento

    Public ReadOnly ModoDemo As Boolean = False

    ' Arquivo das atividades internas, como no frmTimeDet. O I: é o "Dados" do
    ' MEC-SRV-00; se a letra não estiver mapeada, tenta o caminho de rede.
    Private ReadOnly ARQUIVOS_OUTROS As String() = {
        "I:\Sistemas\Nibo\Clientes\Outros.txt",
        "\\mec-srv-00\Dados\Sistemas\Nibo\Clientes\Outros.txt"}

    Private _clientes As List(Of ClienteResumo) = Nothing
    Private _metas As Dictionary(Of String, Decimal) = Nothing

    Private ReadOnly Property ConnString As String
        Get
            Return My.Settings.Myconnection
        End Get
    End Property

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

    ''' <summary>Gestor ou diretoria vê as atividades "administrativo" do Outros.txt (mesma regra do frmTimeDet).</summary>
    Public Function VeAtividadesAdministrativas() As Boolean
        Dim funcao = If(UserFuncao, "").ToUpperInvariant().Trim()
        Return funcao.StartsWith("GESTOR") OrElse funcao.StartsWith("DIRETORIA")
    End Function

    ''' <summary>Aviso da barra de status (vazio na versão SQL).</summary>
    Public Function DescricaoFonteDados() As String
        Return ""
    End Function

    ' ------------------------------------------------------------ apontamentos
    ''' <summary>Apontamentos do usuário logado entre duas datas (Time_GetByData).</summary>
    Public Function Listar(de As Date, ate As Date) As List(Of Apontamento)
        Return ListarDe(Usersys, de, ate)
    End Function

    ''' <summary>Apontamentos de um colaborador entre duas datas (Time_GetByData). Usado pela gestão: período e consultas.</summary>
    Public Function ListarDe(usuarioId As Integer, de As Date, ate As Date) As List(Of Apontamento)
        Dim lista As New List(Of Apontamento)()
        Using cn As New SqlConnection(ConnString)
            cn.Open()
            Using cmd As New SqlCommand("dbo.Time_GetByData", cn) With {.CommandType = CommandType.StoredProcedure, .CommandTimeout = 60}
                cmd.Parameters.Add("@data1", SqlDbType.Char, 10).Value = de.ToString("dd/MM/yyyy")
                cmd.Parameters.Add("@data2", SqlDbType.Char, 10).Value = ate.ToString("dd/MM/yyyy")
                cmd.Parameters.Add("@usuarioid", SqlDbType.Int).Value = usuarioId
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        lista.Add(Ler(rd))
                    End While
                End Using
            End Using
        End Using
        Return lista
    End Function

    ''' <summary>De uma data até hoje (para o resumo e a trava).</summary>
    Public Function ListarDesde(de As Date) As List(Of Apontamento)
        Return Listar(de, Date.Today)
    End Function

    ''' <summary>Banco de horas da pessoa, refeito a partir dos apontamentos desde BancoHoras.INICIO_VIGENCIA.</summary>
    Public Function CalcularBanco() As BancoHoras
        Return BancoHoras.Calcular(ListarDesde(BancoHoras.INICIO_VIGENCIA), BancoHoras.INICIO_VIGENCIA, Date.Today)
    End Function

    ''' <summary>Horas que a pessoa já tem numa data, sem contar o apontamento que está sendo alterado (trava de 10 h por dia).</summary>
    Public Function HorasNoDia(data As Date, excetoId As Integer) As Decimal
        Return Listar(data, data).Where(Function(a) a.Id <> excetoId).Sum(Function(a) a.Horas)
    End Function

    Public Function Obter(id As Integer) As Apontamento
        Using cn As New SqlConnection(ConnString)
            cn.Open()
            Using cmd As New SqlCommand("dbo.Time_Get", cn) With {.CommandType = CommandType.StoredProcedure}
                cmd.Parameters.Add("@CODIGOID", SqlDbType.Int).Value = id
                Using rd = cmd.ExecuteReader()
                    If rd.Read() Then Return Ler(rd)
                End Using
            End Using
        End Using
        Return Nothing
    End Function

    ''' <summary>Grava um apontamento novo e devolve o CODIGOID (Time_Insert).</summary>
    Public Function Inserir(a As Apontamento) As Integer
        Using cn As New SqlConnection(ConnString)
            cn.Open()
            Using cmd As New SqlCommand("dbo.Time_Insert", cn) With {.CommandType = CommandType.StoredProcedure}
                cmd.Parameters.Add("@EMPRESA", SqlDbType.NVarChar, 100).Value = Cortar(a.Empresa, 100)
                cmd.Parameters.Add("@DATA", SqlDbType.SmallDateTime).Value = a.Data.Date
                cmd.Parameters.Add("@TEMPO", SqlDbType.Decimal).Value = a.Horas
                cmd.Parameters.Add("@USUARIOID", SqlDbType.Int).Value = If(a.UsuarioId > 0, a.UsuarioId, Usersys)
                cmd.Parameters.Add("@EMPRESAID", SqlDbType.NVarChar, 7).Value = Cortar(a.EmpresaId, 7)
                Dim pId = cmd.Parameters.Add("@CODIGOID", SqlDbType.Int)
                pId.Direction = ParameterDirection.Output
                cmd.ExecuteNonQuery()
                a.Id = If(pId.Value Is Nothing OrElse pId.Value Is DBNull.Value, 0, Convert.ToInt32(pId.Value))
                Return a.Id
            End Using
        End Using
    End Function

    Public Sub Alterar(a As Apontamento)
        Using cn As New SqlConnection(ConnString)
            cn.Open()
            Using cmd As New SqlCommand("dbo.Time_Update", cn) With {.CommandType = CommandType.StoredProcedure}
                cmd.Parameters.Add("@CODIGOID", SqlDbType.Int).Value = a.Id
                cmd.Parameters.Add("@EMPRESA", SqlDbType.NVarChar, 100).Value = Cortar(a.Empresa, 100)
                cmd.Parameters.Add("@DATA", SqlDbType.SmallDateTime).Value = a.Data.Date
                cmd.Parameters.Add("@TEMPO", SqlDbType.Decimal).Value = a.Horas
                cmd.Parameters.Add("@EMPRESAID", SqlDbType.NVarChar, 7).Value = Cortar(a.EmpresaId, 7)
                cmd.Parameters.Add("@USUARIOID", SqlDbType.Int).Value = If(a.UsuarioId > 0, a.UsuarioId, Usersys)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub Excluir(id As Integer)
        Using cn As New SqlConnection(ConnString)
            cn.Open()
            Using cmd As New SqlCommand("dbo.Time_Delete", cn) With {.CommandType = CommandType.StoredProcedure}
                cmd.Parameters.Add("@CODIGOID", SqlDbType.Int).Value = id
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Depois de gravar ou excluir: o frmMain reavalia a trava de horas e libera
    ''' (ou trava) os menus, como o frmTimeDet faz hoje chamando Opl2.menuacesso().
    ''' </summary>
    Public Sub AposGravar()
        Try
            Opl2.menuacesso()
        Catch
        End Try
    End Sub

    Private Function Ler(rd As IDataRecord) As Apontamento
        Return New Apontamento With {
            .Id = Convert.ToInt32(rd("CODIGOID")),
            .Empresa = Texto(rd("EMPRESA")),
            .Data = If(rd("DATA") Is DBNull.Value, Date.Today, Convert.ToDateTime(rd("DATA")).Date),
            .Horas = If(rd("TEMPO") Is DBNull.Value, 0D, Convert.ToDecimal(rd("TEMPO"))),
            .UsuarioId = If(rd("USUARIOID") Is DBNull.Value, 0, Convert.ToInt32(rd("USUARIOID"))),
            .EmpresaId = Texto(rd("EMPRESAID"))}
    End Function

    ' ------------------------------------------------------------ clientes e atividades
    ''' <summary>Todos os clientes da TbClientes, em cache (ativos primeiro, depois os "não cliente").</summary>
    Public Function Clientes() As List(Of ClienteResumo)
        If _clientes IsNot Nothing Then Return _clientes
        Dim lista As New List(Of ClienteResumo)()
        Using cn As New SqlConnection(ConnString)
            cn.Open()
            Using cmd As New SqlCommand("SELECT RTRIM(ISNULL(TRADUTOR,'')) AS TRADUTOR, RTRIM(ISNULL(NOME,'')) AS NOME, RTRIM(ISNULL(CPF,'')) AS CPF, " &
                                        "ISNULL(CODIGOGRU,0) AS CODIGOGRU, RTRIM(ISNULL(DESCRIGRU,'')) AS DESCRIGRU, RTRIM(ISNULL(SITUACAO,'')) AS SITUACAO " &
                                        "FROM dbo.TbClientes WHERE RTRIM(ISNULL(TRADUTOR,'')) <> '' AND RTRIM(ISNULL(NOME,'')) <> '' ORDER BY NOME", cn) With {.CommandTimeout = 60}
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        Dim situacao = Texto(rd("SITUACAO")).ToUpperInvariant()
                        lista.Add(New ClienteResumo With {
                            .Tradutor = Texto(rd("TRADUTOR")), .Nome = Texto(rd("NOME")), .Cnpj = Texto(rd("CPF")),
                            .CodGrupo = Convert.ToInt32(rd("CODIGOGRU")), .Grupo = Texto(rd("DESCRIGRU")),
                            .Ativo = Not (situacao.StartsWith("NÃO CLIENTE") OrElse situacao.StartsWith("NAO CLIENTE"))})
                    End While
                End Using
            End Using
        End Using
        _clientes = lista.OrderByDescending(Function(c) c.Ativo).ThenBy(Function(c) c.Nome).ToList()
        Return _clientes
    End Function

    ''' <summary>Descarta o cache (F5 na tela).</summary>
    Public Sub LimparCache()
        _clientes = Nothing
        _metas = Nothing
        _colaboradores = Nothing
    End Sub

    ' ------------------------------------------------------------ gestão (gestor e diretoria)
    ''' <summary>Quem abre "Anotações por período" e "Gerenciamento": gestor e diretoria, a mesma regra do "administrativo".</summary>
    Public Function PodeGerenciar() As Boolean
        Return VeAtividadesAdministrativas()
    End Function

    Private _colaboradores As List(Of ColaboradorResumo) = Nothing

    ''' <summary>Colaboradores da TbSenha (sem o ADMINISTRADOR), em cache: ativos primeiro.</summary>
    Public Function Colaboradores() As List(Of ColaboradorResumo)
        If _colaboradores IsNot Nothing Then Return _colaboradores
        Dim lista As New List(Of ColaboradorResumo)()
        Using cn As New SqlConnection(ConnString)
            cn.Open()
            Using cmd As New SqlCommand("SELECT CODIGOID, RTRIM(ISNULL(USUARIO,'')) AS USUARIO, RTRIM(ISNULL(DEPTO,'')) AS DEPTO, ISNULL(DEMITIDO,0) AS DEMITIDO " &
                                        "FROM dbo.TbSenha WHERE RTRIM(ISNULL(USUARIO,'')) <> '' AND UPPER(RTRIM(USUARIO)) <> 'ADMINISTRADOR' ORDER BY USUARIO", cn) With {.CommandTimeout = 60}
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        lista.Add(New ColaboradorResumo With {
                            .Id = Convert.ToInt32(rd("CODIGOID")), .Nome = Texto(rd("USUARIO")), .Depto = Texto(rd("DEPTO")),
                            .Desligado = Convert.ToBoolean(rd("DEMITIDO"))})
                    End While
                End Using
            End Using
        End Using
        _colaboradores = lista.OrderBy(Function(c) c.Desligado).ThenBy(Function(c) c.Nome).ToList()
        Return _colaboradores
    End Function

    ''' <summary>
    ''' Consulta Analítica: as mesmas procedures e a mesma escolha entre elas que o frmTimeConsulta faz hoje
    ''' (Geral → Time_GetByEmpresaGeral; Sem código → SC ou SCTE; Com código → CC ou CCTE; "TE" = todas as empresas).
    ''' </summary>
    Public Function ConsultaAnalitica(f As FiltroAnalitica) As List(Of LinhaAnalitica)
        Dim proc As String
        Dim parametros As New List(Of KeyValuePair(Of String, Object))()
        Dim d1 = f.De.ToString("dd/MM/yyyy")
        Dim d2 = f.Ate.ToString("dd/MM/yyyy")
        Select Case f.Modo
            Case ModoAnalitica.Geral
                proc = "dbo.Time_GetByEmpresaGeral"
                parametros.Add(New KeyValuePair(Of String, Object)("@data1", d1))
                parametros.Add(New KeyValuePair(Of String, Object)("@data2", d2))
                parametros.Add(New KeyValuePair(Of String, Object)("@usuario", f.Usuario))
                parametros.Add(New KeyValuePair(Of String, Object)("@empresa", f.Empresa))
            Case ModoAnalitica.SemCodigoEmpresa
                proc = If(f.TodasEmpresas, "dbo.Time_GetByEmpresaSCTE", "dbo.Time_GetByEmpresaSC")
                parametros.Add(New KeyValuePair(Of String, Object)("@data1", d1))
                parametros.Add(New KeyValuePair(Of String, Object)("@data2", d2))
                parametros.Add(New KeyValuePair(Of String, Object)("@empresa", f.Empresa))
                parametros.Add(New KeyValuePair(Of String, Object)("@usuario", f.Usuario))
            Case Else
                proc = If(f.TodasEmpresas, "dbo.Time_GetByEmpresaCCTE", "dbo.Time_GetByEmpresaCC")
                parametros.Add(New KeyValuePair(Of String, Object)("@data1", d1))
                parametros.Add(New KeyValuePair(Of String, Object)("@data2", d2))
                parametros.Add(New KeyValuePair(Of String, Object)("@usuario", f.Usuario))
                parametros.Add(New KeyValuePair(Of String, Object)("@empresaid", If(f.TodasEmpresas, "", f.EmpresaId)))
        End Select

        Dim lista As New List(Of LinhaAnalitica)()
        Using cn As New SqlConnection(ConnString)
            cn.Open()
            Using cmd As New SqlCommand(proc, cn) With {.CommandType = CommandType.StoredProcedure, .CommandTimeout = 120}
                For Each p In parametros
                    If p.Key = "@data1" OrElse p.Key = "@data2" Then
                        cmd.Parameters.Add(p.Key, SqlDbType.Char, 10).Value = p.Value
                    ElseIf p.Key = "@empresaid" Then
                        cmd.Parameters.Add(p.Key, SqlDbType.Char, 7).Value = p.Value
                    Else
                        cmd.Parameters.Add(p.Key, SqlDbType.NVarChar, 100).Value = p.Value
                    End If
                Next
                Using rd = cmd.ExecuteReader()
                    Dim colunas = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    For i = 0 To rd.FieldCount - 1
                        colunas.Add(rd.GetName(i))
                    Next
                    While rd.Read()
                        lista.Add(New LinhaAnalitica With {
                            .Id = Inteiro(rd, colunas, "codigoid"),
                            .Data = If(colunas.Contains("data") AndAlso rd("data") IsNot DBNull.Value, Convert.ToDateTime(rd("data")).Date, Date.MinValue),
                            .Empresa = TextoCol(rd, colunas, "empresa"),
                            .EmpresaId = TextoCol(rd, colunas, "empresaid"),
                            .Horas = Dec(rd, colunas, "tempo"),
                            .UsuarioId = Inteiro(rd, colunas, "usuarioid"),
                            .Usuario = TextoCol(rd, colunas, "usuario"),
                            .Depto = TextoCol(rd, colunas, "DEPTO"),
                            .CustoHora = Dec(rd, colunas, "custo_por_hora"),
                            .Grupo = TextoCol(rd, colunas, "DESCRIGRU")})
                    End While
                End Using
            End Using
        End Using
        Return lista.OrderByDescending(Function(l) l.Data).ThenBy(Function(l) l.Usuario).ThenBy(Function(l) l.Empresa).ToList()
    End Function

    Private Function TextoCol(rd As IDataRecord, colunas As HashSet(Of String), nome As String) As String
        If Not colunas.Contains(nome) Then Return ""
        Return Texto(rd(nome))
    End Function

    Private Function Inteiro(rd As IDataRecord, colunas As HashSet(Of String), nome As String) As Integer
        If Not colunas.Contains(nome) OrElse rd(nome) Is DBNull.Value Then Return 0
        Return Convert.ToInt32(rd(nome))
    End Function

    Private Function Dec(rd As IDataRecord, colunas As HashSet(Of String), nome As String) As Decimal
        If Not colunas.Contains(nome) OrElse rd(nome) Is DBNull.Value Then Return 0D
        Return Convert.ToDecimal(rd(nome))
    End Function

    ''' <summary>Atividades internas do Outros.txt; sem "administrativo" para quem não é gestor/diretoria.</summary>
    Public Function AtividadesInternas() As List(Of String)
        Dim linhas As New List(Of String)()
        For Each caminho In ARQUIVOS_OUTROS
            Try
                If File.Exists(caminho) Then
                    linhas = File.ReadAllLines(caminho).Select(Function(l) l.Trim()).Where(Function(l) l <> "").Distinct().ToList()
                    Exit For
                End If
            Catch
            End Try
        Next
        If Not VeAtividadesAdministrativas() Then
            linhas = linhas.Where(Function(l) Not l.ToLowerInvariant().Contains("administrativo")).ToList()
        End If
        Return linhas
    End Function

    ''' <summary>True quando o Outros.txt foi encontrado (senão a tela avisa e só permite clientes).</summary>
    Public Function ArquivoAtividadesDisponivel() As Boolean
        Return ARQUIVOS_OUTROS.Any(Function(c) File.Exists(c))
    End Function

    ' ------------------------------------------------------------ metas do mês
    ''' <summary>Horas da competência "MM/yyyy" na TbHorasMes (0 se não cadastrada).</summary>
    Public Function MetaMes(competencia As String) As Decimal
        If _metas Is Nothing Then
            _metas = New Dictionary(Of String, Decimal)()
            Using cn As New SqlConnection(ConnString)
                cn.Open()
                Using cmd As New SqlCommand("dbo.HorasMes_GetList", cn) With {.CommandType = CommandType.StoredProcedure}
                    cmd.Parameters.Add("@competencia", SqlDbType.Char, 7).Value = ""
                    Using rd = cmd.ExecuteReader()
                        While rd.Read()
                            Dim comp = Texto(rd("COMPETENCIA"))
                            If comp <> "" AndAlso Not _metas.ContainsKey(comp) Then
                                _metas(comp) = If(rd("HORASMES") Is DBNull.Value, 0D, Convert.ToDecimal(rd("HORASMES")))
                            End If
                        End While
                    End Using
                End Using
            End Using
        End If
        Dim v As Decimal = 0D
        _metas.TryGetValue(competencia, v)
        Return v
    End Function

    ''' <summary>Totais e dias pendentes do usuário logado, com a mesma regra do frmMain.</summary>
    Public Function CalcularResumo() As ResumoHoras
        Dim hoje = Date.Today
        Dim itens = ListarDesde(ResumoHoras.InicioNecessario(hoje))
        Return ResumoHoras.Calcular(itens, AddressOf MetaMes, Admissao(), hoje)
    End Function

    ' ------------------------------------------------------------ utilidades
    Private Function Texto(v As Object) As String
        If v Is Nothing OrElse v Is DBNull.Value Then Return ""
        Return v.ToString().Trim()
    End Function

    Private Function Cortar(s As String, tamanho As Integer) As String
        Dim t = If(s, "").Trim()
        Return If(t.Length > tamanho, t.Substring(0, tamanho), t)
    End Function

End Module

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
            Using cmd As New SqlCommand("SELECT CODIGOID, RTRIM(ISNULL(USUARIO,'')) AS USUARIO, RTRIM(ISNULL(DEPTO,'')) AS DEPTO, RTRIM(ISNULL(FUNCAO,'')) AS FUNCAO, ISNULL(DEMITIDO,0) AS DEMITIDO " &
                                        "FROM dbo.TbSenha WHERE RTRIM(ISNULL(USUARIO,'')) <> '' AND UPPER(RTRIM(USUARIO)) <> 'ADMINISTRADOR' ORDER BY USUARIO", cn) With {.CommandTimeout = 60}
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        lista.Add(New ColaboradorResumo With {
                            .Id = Convert.ToInt32(rd("CODIGOID")), .Nome = Texto(rd("USUARIO")), .Depto = Texto(rd("DEPTO")), .Funcao = Texto(rd("FUNCAO")),
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

    ''' <summary>
    ''' Consulta Sintética: as mesmas procedures e a mesma escolha entre elas que o frmTimeConsultaAcumulado faz hoje:
    '''   Geral → Time_GetByEmpresaAcumuladoGeralComSem (usuário, empresa, grupo);
    '''   grupo escolhido (qualquer modo) → Time_GetByEmpresaAcumuladoGeral (usuário, empresa, grupo);
    '''   Sem código → SCTE (todas) ou SC (uma); Com código → CCTE (todas) ou CC (uma).
    ''' Código e CNPJ que a procedure não traz são completados pelo nome na TbClientes, como a tela antiga faz.
    ''' O "Acumular por Empresa" é feito depois, em LinhaSintetica.AcumularPorEmpresa.
    ''' </summary>
    Public Function ConsultaSintetica(f As FiltroSintetica) As List(Of LinhaSintetica)
        Dim proc As String
        Dim parametros As New List(Of KeyValuePair(Of String, Object))()
        Dim d1 = f.De.ToString("dd/MM/yyyy")
        Dim d2 = f.Ate.ToString("dd/MM/yyyy")
        If f.Modo = ModoAnalitica.Geral OrElse Not f.TodosGrupos Then
            proc = If(f.Modo = ModoAnalitica.Geral, "dbo.Time_GetByEmpresaAcumuladoGeralComSem", "dbo.Time_GetByEmpresaAcumuladoGeral")
            parametros.Add(New KeyValuePair(Of String, Object)("@data1", d1))
            parametros.Add(New KeyValuePair(Of String, Object)("@data2", d2))
            parametros.Add(New KeyValuePair(Of String, Object)("@usuario", f.Usuario))
            parametros.Add(New KeyValuePair(Of String, Object)("@empresa", f.Empresa))
            parametros.Add(New KeyValuePair(Of String, Object)("@descrigru", If(f.TodosGrupos, FiltroSintetica.TODOS_GRUPOS, f.Grupo)))
        ElseIf f.Modo = ModoAnalitica.SemCodigoEmpresa Then
            proc = If(f.TodasEmpresas, "dbo.Time_GetByEmpresaAcumuladoSCTE", "dbo.Time_GetByEmpresaAcumuladoSC")
            parametros.Add(New KeyValuePair(Of String, Object)("@data1", d1))
            parametros.Add(New KeyValuePair(Of String, Object)("@data2", d2))
            parametros.Add(New KeyValuePair(Of String, Object)("@empresa", f.Empresa))
            parametros.Add(New KeyValuePair(Of String, Object)("@usuario", f.Usuario))
        Else
            proc = If(f.TodasEmpresas, "dbo.Time_GetByEmpresaAcumuladoCCTE", "dbo.Time_GetByEmpresaAcumuladoCC")
            parametros.Add(New KeyValuePair(Of String, Object)("@data1", d1))
            parametros.Add(New KeyValuePair(Of String, Object)("@data2", d2))
            parametros.Add(New KeyValuePair(Of String, Object)("@usuario", f.Usuario))
            parametros.Add(New KeyValuePair(Of String, Object)("@empresaid", If(f.TodasEmpresas, "", f.EmpresaId)))
        End If

        Dim porNome = Clientes().GroupBy(Function(c) c.Nome.Trim().ToUpperInvariant()).ToDictionary(Function(g) g.Key, Function(g) g.First())
        Dim lista As New List(Of LinhaSintetica)()
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
                        Dim l As New LinhaSintetica With {
                            .Codigo = TextoCol(rd, colunas, "TRADUTOR"),
                            .Cnpj = TextoCol(rd, colunas, "cpf"),
                            .Empresa = TextoCol(rd, colunas, "EMPRESA"),
                            .Grupo = TextoCol(rd, colunas, "descrigru"),
                            .UsuarioId = Inteiro(rd, colunas, "USUARIOID"),
                            .Usuario = TextoCol(rd, colunas, "USUARIO"),
                            .Ano = Inteiro(rd, colunas, "ANO"),
                            .Mes = Inteiro(rd, colunas, "MES"),
                            .Horas = Dec(rd, colunas, "TOT")}
                        Dim c As ClienteResumo = Nothing
                        If porNome.TryGetValue(l.Empresa.Trim().ToUpperInvariant(), c) Then
                            If l.Codigo = "" Then l.Codigo = c.Tradutor
                            If l.Cnpj = "" Then l.Cnpj = c.Cnpj
                            If l.Grupo = "" Then l.Grupo = c.Grupo
                        End If
                        lista.Add(l)
                    End While
                End Using
            End Using
        End Using
        Return lista.OrderBy(Function(l) l.Ano).ThenBy(Function(l) l.Mes).ThenBy(Function(l) l.Empresa).ThenBy(Function(l) l.Usuario).ToList()
    End Function

    ''' <summary>
    ''' Consulta Por funcionário ("Consulta Geral", frmTimeFuncionario): horas diretas de
    ''' Time_GetByDataTudoHorasCustoDireto e indiretas de Time_GetByDataTudoHorasCustoIndireto,
    ''' somadas por USUARIOID, com departamento e função da TbSenha. Só quem tem horas no período.
    ''' A tela antiga trabalha por competência (um mês); aqui o período é livre e, com o mês
    ''' inteiro, os números são os mesmos.
    ''' </summary>
    Public Function ConsultaPorFuncionario(de As Date, ate As Date) As List(Of LinhaFuncionario)
        Dim mapa As New Dictionary(Of Integer, LinhaFuncionario)()
        LerHorasPorUsuario("dbo.Time_GetByDataTudoHorasCustoDireto", de, ate, mapa, True)
        LerHorasPorUsuario("dbo.Time_GetByDataTudoHorasCustoIndireto", de, ate, mapa, False)
        Dim pessoas = Colaboradores()
        For Each l In mapa.Values
            Dim c = pessoas.FirstOrDefault(Function(x) x.Id = l.UsuarioId)
            If c IsNot Nothing Then
                If l.Usuario = "" Then l.Usuario = c.Nome
                l.Depto = c.Depto
                l.Funcao = c.Funcao
            End If
        Next
        Return mapa.Values.Where(Function(l) l.Total > 0D).OrderBy(Function(l) l.Usuario).ToList()
    End Function

    Private Sub LerHorasPorUsuario(proc As String, de As Date, ate As Date, mapa As Dictionary(Of Integer, LinhaFuncionario), diretas As Boolean)
        Dim pessoas = Colaboradores()
        Using cn As New SqlConnection(ConnString)
            cn.Open()
            Using cmd As New SqlCommand(proc, cn) With {.CommandType = CommandType.StoredProcedure, .CommandTimeout = 120}
                cmd.Parameters.Add("@data1", SqlDbType.Char, 10).Value = de.ToString("dd/MM/yyyy")
                cmd.Parameters.Add("@data2", SqlDbType.Char, 10).Value = ate.ToString("dd/MM/yyyy")
                Using rd = cmd.ExecuteReader()
                    Dim colunas = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    For i = 0 To rd.FieldCount - 1
                        colunas.Add(rd.GetName(i))
                    Next
                    While rd.Read()
                        Dim nome = TextoCol(rd, colunas, "usuarios")
                        Dim id = Inteiro(rd, colunas, "usuarioid")
                        If id = 0 Then
                            Dim c = pessoas.FirstOrDefault(Function(x) String.Equals(x.Nome, nome, StringComparison.OrdinalIgnoreCase))
                            If c IsNot Nothing Then id = c.Id
                        End If
                        Dim l As LinhaFuncionario = Nothing
                        If Not mapa.TryGetValue(id, l) Then
                            l = New LinhaFuncionario With {.UsuarioId = id, .Usuario = nome}
                            mapa(id) = l
                        End If
                        Dim h = Math.Round(Dec(rd, colunas, "horas"), 2)
                        If diretas Then l.HorasDiretas += h Else l.HorasIndiretas += h
                    End While
                End Using
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Consulta Por usuário e data (frmTimeConsultausuario): Time_GetByEmpresaUsuario devolve o total
    ''' de horas de cada pessoa no intervalo; quem não lançou nada entra com zero (colaboradores
    ''' ativos da TbSenha), e cada linha é comparada com o esperado do intervalo.
    ''' </summary>
    Public Function ConsultaPorUsuarioData(de As Date, ate As Date) As List(Of LinhaUsuarioData)
        Dim esperado = LinhaUsuarioData.HorasEsperadas(de, ate)
        Dim pessoas = Colaboradores()
        Dim lista As New List(Of LinhaUsuarioData)()
        Using cn As New SqlConnection(ConnString)
            cn.Open()
            Using cmd As New SqlCommand("dbo.Time_GetByEmpresaUsuario", cn) With {.CommandType = CommandType.StoredProcedure, .CommandTimeout = 120}
                cmd.Parameters.Add("@data1", SqlDbType.Char, 10).Value = de.ToString("dd/MM/yyyy")
                cmd.Parameters.Add("@data2", SqlDbType.Char, 10).Value = ate.ToString("dd/MM/yyyy")
                Using rd = cmd.ExecuteReader()
                    Dim colunas = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    For i = 0 To rd.FieldCount - 1
                        colunas.Add(rd.GetName(i))
                    Next
                    While rd.Read()
                        Dim nome = TextoCol(rd, colunas, "USUARIOS")
                        Dim c = pessoas.FirstOrDefault(Function(x) String.Equals(x.Nome, nome, StringComparison.OrdinalIgnoreCase))
                        lista.Add(New LinhaUsuarioData With {
                            .UsuarioId = If(c Is Nothing, 0, c.Id), .Usuario = nome, .Depto = If(c Is Nothing, "", c.Depto),
                            .Horas = Dec(rd, colunas, "HORAS"), .Esperado = esperado})
                    End While
                End Using
            End Using
        End Using
        For Each c In pessoas.Where(Function(x) Not x.Desligado)
            Dim nome = c.Nome
            If Not lista.Any(Function(l) String.Equals(l.Usuario, nome, StringComparison.OrdinalIgnoreCase)) Then
                lista.Add(New LinhaUsuarioData With {.UsuarioId = c.Id, .Usuario = c.Nome, .Depto = c.Depto, .Horas = 0D, .Esperado = esperado})
            End If
        Next
        Return lista.OrderBy(Function(l) l.Usuario).ToList()
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

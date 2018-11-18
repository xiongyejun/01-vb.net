Imports System.Data.OleDb

''' <summary> ADO.NET的操作 </summary>
Public Class CADO
    Private str_sql As String
    Private conn As OleDbConnection = Nothing

    ''' <summary> 设置SQL语句 </summary>
    WriteOnly Property StrSql() As String
        Set(ByVal value As String)
            str_sql = value
        End Set
    End Property

    ''' <summary> 根据SoureFile设置连接字符串 </summary>
    Private Function StrConnection(ByVal str_SoureFile As String) As String
        Dim str As String = ""
        str_SoureFile = LCase(str_SoureFile)

        If Right(str_SoureFile, 4) = ".xls" Then
            str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & str_SoureFile
            str &= ";Extended Properties=""Excel 8.0;HDR=YES"";"

        ElseIf Right(str_SoureFile, 4) = ".mdb" Then
            str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & str_SoureFile

        ElseIf Right(str_SoureFile, 5) = ".xlsx" Or Right(str_SoureFile, 5) = ".xlsm" Or Right(str_SoureFile, 5) = ".xlsb" Then
            str = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" & str_SoureFile
            str &= ";Extended Properties=""Excel 12.0;HDR=YES"";"

        ElseIf Right(str_SoureFile, 6) = ".accdb" Then
            str = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" & str_SoureFile & ";Jet OLEDB:Database Password=dimiaslong" 'dimiaslong

        ElseIf Right(str_SoureFile, 4) = ".txt" Then
            'str_SoureFile应该为文件的路径
            str = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" & Left(str_SoureFile, InStrRev(str_SoureFile, "\"))
            str &= ";Extended Properties=""TEXT;HDR=YES"";"
        End If

        Return str
    End Function

    ''' <summary> 打开数据库 </summary>
    ''' <returns>打开出错就返回Nothing</returns>
    Private Function ConnOpen(ByVal str_conn As String) As OleDb.OleDbConnection
        Try
            conn = New OleDb.OleDbConnection(str_conn)
            conn.Open()
        Catch ex As Exception
            MsgBox(ex.Message & vbNewLine & vbNewLine & "可能是文件加密了或者文件后缀与实际文件格式不一致。")
            conn = Nothing
        End Try

        Return conn
    End Function

    ''' <summary> 获取数据库中的表名称 </summary>
    Function GetTables() As String()
        If conn Is Nothing Then Return Nothing

        Dim table As DataTable = conn.GetSchema("Tables")
        Dim str() As String = Nothing
        Dim k As Long = 0

        For Each row As DataRow In table.Rows
            If row.Item(3) = "TABLE" Then
                ReDim Preserve str(k)
                str(k) = row.Item(2).ToString
                'Console.WriteLine(str(k))
                k += 1
            End If
        Next

        table.Dispose()
        If conn.State = 1 Then conn.Close()

        Return str
    End Function

    ''' <summary>
    ''' 执行SQL语句，出错返回False
    ''' </summary>
    Function ExcuteSql() As Boolean
        If conn Is Nothing Then Return Nothing
        Dim adapter As OleDbDataAdapter = Nothing
        Dim table As DataTable = New DataTable

        Try
            adapter = New OleDbDataAdapter(str_sql, conn)
            adapter.Fill(table)
        Catch ex As Exception
            Return False
        End Try

        table.Dispose()
        table = Nothing
        adapter.Dispose()
        adapter = Nothing
        If conn.State = 1 Then conn.Close()

        Return True
    End Function
    ' ''' <summary>
    ' ''' 执行SQL语句，出错返回False
    ' ''' </summary>
    'Function UpdateData(ByVal dt As DataTable) As Boolean
    '    If conn Is Nothing Then Return Nothing
    '    Dim cmd As OleDbCommand = New OleDbCommand()
    '    cmd.Connection = conn

    '    Dim arr_parameters(dt.Columns.Count - 1) As String
    '    Dim arr_type(dt.Columns.Count - 1) As System.Data.OleDb.OleDbType
    '    Dim arr_field(dt.Columns.Count - 1) As String
    '    For i As Integer = 0 To arr_field.Length - 1
    '        arr_field(i) = dt.Columns(i).ColumnName
    '        '对于 OleDbDataAdapter 对象和 OdbcDataAdapter 对象，必须使用问号 (?) 占位符来标识参数
    '        arr_parameters(i) = "?" ' "@COL_" & i.ToString
    '        If i = 2 Then
    '            arr_type(i) = OleDbType.Integer
    '        Else
    '            arr_type(i) = OleDbType.VarChar
    '        End If

    '    Next
    '    Dim str_cmdText As String = "Insert Into " & dt.TableName & "(" & Join(arr_field, ",") & ") Values (" & Join(arr_parameters, ",") & ")"
    '    cmd.CommandText = str_cmdText

    '    For i As Integer = 0 To arr_type.Length - 1
    '        cmd.Parameters.Add(New OleDbParameter("@" & arr_field(i), arr_type(i)))
    '    Next

    '    Try
    '        Dim adapter As OleDbDataAdapter = New OleDbDataAdapter("Select * From " & dt.TableName & " Where 1=2", conn)
    '        adapter.InsertCommand = New OleDbCommand(str_cmdText)
    '        For i As Integer = 0 To arr_type.Length - 1
    '            adapter.InsertCommand.Parameters.Add(arr_field(i), arr_type(i), 100, arr_field(i))
    '        Next
    '        Dim ds As DataSet = New DataSet
    '        adapter.Fill(ds)
    '        For Each r As DataRow In dt.Rows
    '            ds.Tables(0).Rows.Add(r)
    '        Next

    '        adapter.Update(ds)
    '    Catch ex As Exception
    '        Return False
    '    End Try


    '    If conn.State = 1 Then conn.Close()

    '    Return True
    'End Function
    ''' <summary>
    ''' 更新数据，OleDbCommandBuilder会根据dataRow的RowState属性来判断是insert还是delete还是update
    ''' </summary>
    Function UpdateData(ByVal dt As DataTable, ByVal TableName As String) As Boolean
        If conn Is Nothing Then Return Nothing
        Dim adapter As OleDbDataAdapter = Nothing

        Try
            adapter = New OleDbDataAdapter("Select * From " & TableName & " Where 1=2", conn)
            Dim cmd As OleDbCommandBuilder = New OleDbCommandBuilder(adapter)
            adapter.Update(dt)
            cmd = Nothing
        Catch ex As Exception
            Return False
        End Try

        adapter.Dispose()
        adapter = Nothing
        If conn.State = 1 Then conn.Close()

        Return True
    End Function

    Function GetData() As DataTable
        If conn Is Nothing Then Return Nothing

        Dim adapter As OleDbDataAdapter = New OleDbDataAdapter(str_sql, conn)
        adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey '为了获取主键信息
        Dim table As DataTable = New DataTable
        adapter.Fill(table)
        adapter.Dispose()
        If conn.State = 1 Then conn.Close()

        Return table
    End Function

    Protected Overrides Sub Finalize()
        conn = Nothing
        MyBase.Finalize()
    End Sub

    Public Sub New(ByVal SoureFile As String)
        Dim str As String = StrConnection(SoureFile)
        If str <> "" Then ConnOpen(str)
    End Sub

    ''' <summary>
    ''' 创建新表
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>"Create Table 人员 (姓名 Char, 年龄 Integer)"</remarks>
    Function CreateTable() As Boolean
        Try
            Dim command As New OleDbCommand(str_sql, conn)
            command.ExecuteNonQuery()
            If conn.State = 1 Then conn.Close()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function


End Class

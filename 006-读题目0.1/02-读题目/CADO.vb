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

    Private If_Err As Boolean
    ReadOnly Property IfErr() As Boolean
        Get
            IfErr = If_Err
        End Get
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
            str = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" & str_SoureFile ' & ";Jet OLEDB:Database Password=dimiaslong=1"

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
            MsgBox(ex.Message)
            conn = Nothing
            If_Err = True
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
    ''' OleDbCommandBuilder会自动根据datarowStates来判断是插入、删除、更新
    ''' </summary>
    Function UpdateData(ByVal dt As DataTable, TableName As String) As Boolean
        If conn Is Nothing Then Return Nothing
        Dim adapter As OleDbDataAdapter = Nothing

        Try
            adapter = New OleDbDataAdapter("Select * From " & TableName & " Where 1=2", conn)
            Dim cmd As OleDbCommandBuilder = New OleDbCommandBuilder(adapter)
            adapter.Update(dt)
            cmd = Nothing
        Catch ex As Exception
            MsgBox("数据添加出错" & ex.Message)
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
        If_Err = False
        Dim str As String = StrConnection(SoureFile)
        If str <> "" Then ConnOpen(str)
    End Sub


    ''' <summary>
    ''' 执行sql
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>"Create Table 人员 (姓名 Char, 年龄 Integer)"</remarks>
    Function ExcuteSql() As Boolean
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

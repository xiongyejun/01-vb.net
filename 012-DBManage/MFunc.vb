Imports System.Windows.Forms.Form
Imports System.Threading.Thread

Module MFunc
    ''' <summary>
    ''' 读取设置
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function ReadSet()
        Dim c_set As New CSet(FileSet)
        DicSet = c_set.Read()
        c_set = Nothing

        Return 1
    End Function
    ''' <summary>
    ''' 保存设置
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function WriteSet()
        Dim c_set As New CSet(FileSet)
        c_set.Write(DicSet)
        c_set = Nothing


        Return 1
    End Function

    ''' <summary>
    ''' 选择数据库文件
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function SelectDB() As Boolean
        Dim fd As OpenFileDialog = New OpenFileDialog

        fd.Filter = "Access(*.mdb;*.accdb)|*.mdb;*.accdb"
        If fd.ShowDialog = vbOK Then
            DB_Info.Path = fd.FileName
            InitDBInfo()

            Return True
        End If

        Return False
    End Function

    Function InitDBInfo() As Integer
        DB_Info.DicTableIndex = New Dictionary(Of String, Integer)
        DB_Info.Tables = GetTables(DB_Info.Path)
        GetPointer()
        ExtendField()

        Return 1
    End Function


    ''' <summary>
    ''' 获取表格的信息：表格名称、字段、字段类型、主键等
    ''' </summary>
    ''' <param name="DBPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetTables(ByVal DBPath As String) As TableInfo()
        Dim c_ado As New CADO(DBPath)
        Dim Arr() As String = c_ado.GetTables()
        Dim Tables(Arr.Length - 1) As TableInfo
        Dim dic As New Dictionary(Of String, Integer)

        For i As Integer = 0 To Arr.Length - 1
            Dim table_info As TableInfo = Nothing
            table_info.Name = Arr(i)
            DB_Info.DicTableIndex(table_info.Name) = i

            c_ado.StrSql = "Select * From [" & table_info.Name & "] Where 1=2"
            Dim dt As DataTable = c_ado.GetData()

            ReDim table_info.Field.Name(dt.Columns.Count - 1)
            ReDim table_info.Field.Type(dt.Columns.Count - 1)
            ReDim table_info.Field.Pointer(dt.Columns.Count - 1)
            For j As Integer = 0 To dt.Columns.Count - 1
                '字段名称
                table_info.Field.Name(j) = dt.Columns(j).ColumnName
                '字段类型
                table_info.Field.Type(j) = dt.Columns(j).DataType
                dic(table_info.Field.Name(j)) = j
            Next
            '获取主键信息
            Dim cols() As DataColumn = dt.PrimaryKey
            ReDim table_info.Field.PrimaryKey(cols.Length - 1)
            ReDim table_info.Field.PrimaryKeyIndex(cols.Length - 1)
            For j As Integer = 0 To cols.Length - 1
                table_info.Field.PrimaryKey(j) = cols(j).ColumnName
                table_info.Field.PrimaryKeyIndex(j) = dic(cols(j).ColumnName)
            Next

            Tables(i) = table_info
        Next

        Return Tables
    End Function
    ''' <summary>
    ''' 字段是否链接了其他表的信息，-1是没有，大于-1就是其他表的下标
    ''' 不放到GetTables里，是因为表出现的顺序是不确定的，DB_Info.DicTableIndex有可能还没完成初始化
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetPointer() As Integer
        For i As Integer = 0 To DB_Info.Tables.Length - 1
            For j As Integer = 0 To DB_Info.Tables(i).Field.Name.Length - 1
                If DB_Info.Tables(i).Field.Name(j) Like "*?ID" Then
                    Dim tmpTable As String = DB_Info.Tables(i).Field.Name(j).Substring(0, DB_Info.Tables(i).Field.Name(j).Length - 2)
                    Dim tmpTableIndex As Integer = DB_Info.DicTableIndex(tmpTable)
                    DB_Info.Tables(i).Field.Pointer(j) = tmpTableIndex
                    '记录引用了ID的表格
                    Dim k As Integer = 0
                    If DB_Info.Tables(tmpTableIndex).bUseMyIdTables Then
                        k = DB_Info.Tables(tmpTableIndex).UseMyIdTables.Length
                    Else
                        DB_Info.Tables(tmpTableIndex).bUseMyIdTables = True
                    End If
                    ReDim Preserve DB_Info.Tables(tmpTableIndex).UseMyIdTables(k)
                    DB_Info.Tables(tmpTableIndex).UseMyIdTables(k) = i
                Else
                    DB_Info.Tables(i).Field.Pointer(j) = -1
                End If
            Next
        Next
        Return 1
    End Function

    ''' <summary>
    ''' 把xxID的字段扩展出来，并构建sql
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ExtendField() As Integer
        For i As Integer = 0 To DB_Info.Tables.Length - 1
            ExtendFieldItem(DB_Info.Tables(i))
        Next

        Return 1
    End Function
    Private Function ExtendFieldItem(ByRef Table_Info As TableInfo) As Integer
        '报价产品：Select A.名称,B.名称 As 用户,C.名称 As 产品名称,C.型号 As 产品型号,A.数量,A.金额,A.时间,A.备注 From 报价产品 A,用户 B,产品 C Where A.用户ID=B.ID And A.产品ID=C.ID
        Dim strSelect As String = "Select "
        Dim strTables As String = " From " & Table_Info.Name
        Dim strWhere As String = " Where 1=1 "
        Dim extend_field(ExtendFieldEnum.Count - 1) As ArrayList '记录扩展出来的字段，最后再存到ExtendFieldName里

        For i As Integer = 0 To extend_field.Length - 1
            extend_field(i) = New ArrayList
        Next

        DGExtendFieldItem(Table_Info,
                          strSelect,
                          strTables,
                         strWhere,
                         extend_field)

        '去掉select最后的逗号
        Table_Info.SqlExtend = strSelect.Substring(0, strSelect.Length - 1) & strTables & strWhere
        '将ArrayList的数据转化为数组
        ReDim Table_Info.ExtendField.Tables((extend_field(0).Count - 1))
        ReDim Table_Info.ExtendField.ExtendFieldName((extend_field(0).Count - 1))
        ReDim Table_Info.ExtendField.FieldName((extend_field(0).Count - 1))
        ReDim Table_Info.ExtendField.ExtendFieldType((extend_field(0).Count - 1))
        For i As Integer = 0 To extend_field(0).Count - 1
            Table_Info.ExtendField.Tables(i) = extend_field(ExtendFieldEnum.Tables)(i).ToString
            Table_Info.ExtendField.FieldName(i) = extend_field(ExtendFieldEnum.FieldName)(i).ToString
            Table_Info.ExtendField.ExtendFieldName(i) = extend_field(ExtendFieldEnum.ExtendFieldName)(i).ToString
            Table_Info.ExtendField.ExtendFieldType(i) = CType(extend_field(ExtendFieldEnum.ExtendFieldType).Item(i), System.Type)
        Next

        Return 1
    End Function
    ''' <summary>
    ''' 递归处理 xxID 的字段
    ''' </summary>
    ''' <param name="strSelect"></param>
    ''' <param name="extend_field"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DGExtendFieldItem(ByRef Table_Info As TableInfo,
                                       ByRef strSelect As String,
                                       ByRef strTables As String,
                                       ByRef strWhere As String,
                                       ByRef extend_field() As ArrayList) As Integer

        Dim tmpTable As String '= strFieldName.Substring(0, strFieldName.Length - 2)

        For i As Integer = 0 To Table_Info.Field.Name.Length - 1
            Dim strFieldName As String = Table_Info.Field.Name(i)

            If Table_Info.Field.Pointer(i) > -1 Then
                'xxID指向的table
                tmpTable = DB_Info.Tables(Table_Info.Field.Pointer(i)).Name
                '获取xxID所有的字段
                DGExtendFieldItem(DB_Info.Tables(Table_Info.Field.Pointer(i)),
                                  strSelect,
                                  strTables,
                                  strWhere,
                                  extend_field)

                strTables &= ("," & tmpTable)
                strWhere &= "And " & Table_Info.Name & "." & strFieldName & "=" & tmpTable & ".ID "
            Else
                extend_field(ExtendFieldEnum.Tables).Add(Table_Info.Name)
                extend_field(ExtendFieldEnum.FieldName).Add(strFieldName)
                extend_field(ExtendFieldEnum.ExtendFieldName).Add(Table_Info.Name & "." & strFieldName)
                extend_field(ExtendFieldEnum.ExtendFieldType).Add(Table_Info.Field.Type(i))

                strSelect &= (Table_Info.Name & "." & strFieldName & " As " & Table_Info.Name & strFieldName & ",")
            End If

        Next

        Return 1
    End Function

    ''' <summary>
    ''' 执行查找
    ''' </summary>
    ''' <param name="DBPath"></param>
    ''' <param name="StrSql"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function DoSearch(ByVal DBPath As String, ByVal StrSql As String) As DataTable
        Dim c_ado As New CADO(DBPath)
        c_ado.StrSql = StrSql
        Dim dt As DataTable = c_ado.GetData()
        c_ado = Nothing

        Return dt
    End Function

    ''' <summary>
    ''' 备份文件
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function BackupFile(ByVal FileName As String) As Boolean
        Dim strNow As String = Strings.Format(Now(), "yyyyMMddHHmmss")
        Dim fileExt As String = IO.Path.GetExtension(FileName)
        Dim newFile As String = FileName.Substring(0, FileName.Length - fileExt.Length) & strNow & fileExt

        IO.File.Copy(FileName, newFile)

        Return True
    End Function
End Module

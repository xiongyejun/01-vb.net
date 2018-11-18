''' <summary>
''' 从excel中导入数据
''' </summary>
''' <remarks></remarks>

Structure ImportData
    Dim Arr(,) As Object
    Dim LBound As Integer
    Dim Rows As Integer
    Dim Cols As Integer
    Dim NextCol As Integer '下一个要处理的列
    Dim Result(,) As String
End Structure


Module MImportFromExcel
    Function ImportFromExcel() As Boolean
        Dim rng As Object = GetRngFromExcel()
        If rng Is Nothing Then
            Return -1
        End If

        '初始化数据
        Dim Import_Data As ImportData
        Import_Data.Arr = rng.Value
        Import_Data.LBound = 1
        Import_Data.Rows = Import_Data.Arr.GetUpperBound(0)
        Import_Data.Cols = Import_Data.Arr.GetUpperBound(1)
        Import_Data.NextCol = 1
        ReDim Import_Data.Result(Import_Data.Rows - 1, 0)
        Import_Data.Result(0, 0) = "结果"
        '检查标题
        If Not CheckField(Import_Data) Then
            MsgBox("标题对应不上。")
            Return False
        End If

        If Not AddData(Import_Data, DB_Info.Tables(DB_Info.ActivateTableIndex)) Then
            Return False
        End If

        rng.Columns(1).Offset(0, rng.Columns.Count).Value = Import_Data.Result

        Return True
    End Function

    ''' <summary>
    ''' 添加数据
    ''' 遍历每个字段，如果Pointer为1，表示指向了另外一个table，就递归过去
    ''' </summary>
    ''' <param name="Import_Data"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function AddData(ByRef Import_Data As ImportData, ByRef Table As TableInfo) As Boolean
        Dim TablePointerArr(Table.Field.Name.Length - 1) As Integer  '当前表的列，对应的Arr中的列

        For i As Integer = 0 To Table.Field.Name.Length - 1
            TablePointerArr(i) = Import_Data.NextCol

            If Table.Field.Pointer(i) = -1 Then
                Import_Data.NextCol += 1
            Else
                '指向了其他的表，递归过去
                If Not AddData(Import_Data, DB_Info.Tables(Table.Field.Pointer(i))) Then
                    Return False
                End If
            End If
        Next

        Dim IsActivateTable As Boolean = (Table.Name = DB_Info.ActivateTable)
        '获取表的结构
        Dim c_ado As New CADO(DB_Info.Path)
        '将表的主键放入字典，item存放id，
        c_ado.StrSql = "Select * From [" & Table.Name & "]"
        Dim dt As DataTable = c_ado.GetData()
        Dim dic As New Dictionary(Of String, Integer)
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim key As String = GetKeyByDt(dt, Table, i)
            If IsActivateTable Then
                dic(key) = 0 '当前添加的表，仅需要记录是否存在
            Else
                dic(key) = dt.Rows(i).Item("ID") '引用的其他表格是要获取他的ID
            End If
        Next

        Dim dtCount As Integer = dt.Rows.Count '记录原来的记录数
        Dim dtCount2 As Integer = dtCount
        Dim startCol As Integer = TablePointerArr(0)
        '判断Arr中的数据是否已经出现了, 出现的就获取ID，没有的就放入dt中update
        For i As Integer = Import_Data.LBound + 1 To Import_Data.Rows
            Dim key As String = GetKeyByArr(Import_Data, Table, i, TablePointerArr)

            If key.Length = 0 Then
                MsgBox("[" & Table.Name & "]主键为空")
                Return False
            End If

            If dic.ContainsKey(key) Then
                Import_Data.Arr(i, startCol) = dic(key)
            Else
                Dim r As DataRow = dt.NewRow()

                Dim startIndex As Integer = 0 '当前表格的所有列都要记录
                If Not IsActivateTable Then startIndex = 1 '其他表格第0列是ID，不用赋值，自动增加的

                For j As Integer = startIndex To TablePointerArr.Length - 1
                    If Import_Data.Arr(i, TablePointerArr(j)) Is Nothing Then
                        r.Item(j) = ""
                    Else
                        r.Item(j) = Import_Data.Arr(i, TablePointerArr(j))
                    End If
                Next
                dt.Rows.Add(r)
                dtCount2 += 1
                Import_Data.Result(i - 1, 0) &= "Add To [" & Table.Name & "]、"

                If IsActivateTable Then
                    dic(key) = 0
                Else
                    dic(key) = dt.Rows(dtCount2 - 1).Item("ID") '记录新加的ID
                End If

            End If
        Next

        '如果有新增的，就更新
        If dtCount2 > dtCount Then
            c_ado.UpdateData(dt, Table.Name)

            If Not IsActivateTable Then
                '其他表格的情况下要获取他的ID
                For i As Integer = Import_Data.LBound + 1 To Import_Data.Rows
                    If Import_Data.Arr(i, startCol) Is Nothing Then '对ID还是空的进行赋值
                        Dim key As String = GetKeyByArr(Import_Data, Table, i, TablePointerArr)
                        Import_Data.Arr(i, startCol) = dic(key)
                    End If
                Next
            End If

        End If

        c_ado = Nothing

        Return True
    End Function

    ''' <summary>
    ''' 从arr中，根据列来获取主键的信息
    ''' </summary>
    ''' <param name="Import_Data"></param>
    ''' <param name="Table"></param>
    ''' <param name="RowIndex"></param>
    ''' <param name="TablePointerArr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetKeyByArr(ByRef Import_Data As ImportData, ByRef Table As TableInfo, ByVal RowIndex As Integer, ByRef TablePointerArr() As Integer) As String
        Dim key As String = ""

        For j As Integer = 0 To Table.Field.PrimaryKeyIndex.Length - 1
            'Table.Field.PrimaryKeyIndex(j) 主键对应的table的下标
            'TablePointerArr    转化为Arr的下标
            If Import_Data.Arr(RowIndex, TablePointerArr(Table.Field.PrimaryKeyIndex(j))) IsNot Nothing Then
                key &= Import_Data.Arr(RowIndex, TablePointerArr(Table.Field.PrimaryKeyIndex(j))).ToString
            End If
        Next

        Return key
    End Function

    Private Function GetKeyByDt(ByRef dt As DataTable, ByRef Table As TableInfo, ByVal RowIndex As Integer) As String
        Dim key As String = ""

        For j As Integer = 0 To Table.Field.PrimaryKeyIndex.Length - 1
            key &= dt.Rows(RowIndex).Item(Table.Field.PrimaryKeyIndex(j)).ToString
        Next

        Return key
    End Function

    ''' <summary>
    ''' 检查标题，按扩展的标题检查
    ''' </summary>
    ''' <param name="Import_Data"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckField(ByRef Import_Data As ImportData) As Boolean
        For i As Integer = Import_Data.LBound To Import_Data.Cols
            If Import_Data.Arr(Import_Data.LBound, i).ToString <> DB_Info.Tables(DB_Info.ActivateTableIndex).ExtendField.ExtendFieldName(i - 1) Then
                Return False
            End If
        Next

        Return True
    End Function

    ''' <summary>
    ''' 从excel获取rng
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetRngFromExcel() As Object
        Dim c_excel As New CExcel
        c_excel.GetExcel()
        Dim rng As Object = c_excel.GetRng()

        Return rng
    End Function
End Module

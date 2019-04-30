''' <summary>
''' 从excel中导入数据
''' </summary>
''' <remarks></remarks>

Structure ImportData
    Dim CExcel As CExcel
    Dim Arr(,) As Object
    Dim LBoundRow As Integer
    Dim LBoundCol As Integer
    Dim UBoundRow As Integer
    Dim UBoundCol As Integer

    Dim NextCol As Integer '下一个要处理的列
    Dim Result(,) As String

    Dim rng As Object
    Dim f As FMain

    Dim stmtHandleGetID() As Integer
    Dim stmtHandleInsert() As Integer
End Structure

Module MImportFromExcel
    Private Const CONTEXT_FH As String = "|" 'key连接的时候的分隔符号
    Private ds As ImportData = Nothing

    Function ImportFromExcel(ByRef f As FMain) As Boolean
        ds.rng = GetRngFromExcel(ds)
        If ds.rng Is Nothing Then
            Return False
        End If

        ds.f = f
        ds.f.Show()
        '初始化数据
        ds.Arr = ds.rng.Value

        ds.LBoundRow = ds.Arr.GetLowerBound(0) + 1 '第1行是标题
        ds.LBoundCol = ds.Arr.GetLowerBound(1)

        ds.UBoundRow = ds.Arr.GetUpperBound(0)
        ds.UBoundCol = ds.Arr.GetUpperBound(1)

        ds.NextCol = ds.LBoundCol

        Dim t As Threading.Thread = New Threading.Thread(AddressOf StartImport)
        t.Start()

        Return True
    End Function

    Sub StartImport()
        '将Arr中的空白，都改成null，让sqlite放default value，如果是不允许null，也正好让sqlite检测
        For i As Integer = ds.LBoundRow To ds.UBoundRow
            For j As Integer = ds.LBoundCol To ds.UBoundCol
                If ds.Arr(i, j) Is Nothing OrElse ds.Arr(i, j).ToString = "" Then
                    ds.Arr(i, j) = "null"
                End If

            Next

            If i Mod 10 = 0 Then
                ds.f.Text = String.Format("PriceDB:正在处理数据:{0}/{1}", i, ds.UBoundRow)
            End If
        Next

        ReDim ds.Result(ds.UBoundRow, 0)
        ds.Result(0, 0) = "结果"

        '检查标题
        If Not CheckField(ds) Then
            MsgBox("标题对应不上。")
            DisposData(ds)
            Return
        End If

        If cdb.ExecuteNonQuery("begin") Then
            MsgBox("ExecuteNonQuery(""begin"")出错了。" & vbNewLine & cdb.GetErr)
            Return
        End If

        ReDim ds.stmtHandleGetID(DB_Info.Tables.Length - 1)
        ReDim ds.stmtHandleInsert(DB_Info.Tables.Length - 1)

        '获取每个有ID项的表的最后1个ID
        For i As Integer = 0 To DB_Info.Tables.Length - 1
            Dim values(DB_Info.Tables(i).Field.Name.Length - 1) As String
            For j As Integer = 0 To values.Length - 1
                values(j) = "?"
            Next
            Dim ret As Integer = cdb.Prepare(String.Format("insert into {0} values ({1})", DB_Info.Tables(i).Name, Join(values, ",")), ds.stmtHandleInsert(i))
            If ret Then
                MsgBox(cdb.GetErr)
                cdb.ExecuteNonQuery("rollback")
                GoTo Exit_sub
            End If

            ReDim values(DB_Info.Tables(i).Field.PrimaryKey.Length - 1)
            For j As Integer = 0 To values.Length - 1
                values(j) = DB_Info.Tables(i).Field.PrimaryKey(j) & "=?"
            Next
            ret = cdb.Prepare(String.Format("select ID from {0} where {1}", DB_Info.Tables(i).Name, Join(values, " and ")), ds.stmtHandleGetID(i))
            If ret Then
                MsgBox(cdb.GetErr)
                cdb.ExecuteNonQuery("rollback")
                GoTo Exit_sub
            End If

            If DB_Info.Tables(i).bHasID Then
                DB_Info.Tables(i).LastID = cdb.GetColZeroValue(String.Format("select max(ID) from {0}", DB_Info.Tables(i).Name))
                If DB_Info.Tables(i).LastID = -1 Then
                    MsgBox(DB_Info.Tables(i).Name & " LastID获取出错了。" & vbNewLine & cdb.GetErr)
                    GoTo Exit_sub
                End If

            End If
        Next

        '添加数据
        If Not AddData(ds, DB_Info.Tables(DB_Info.ActivateTableIndex), True) Then
            If cdb.ExecuteNonQuery("rollback") Then
                MsgBox("ExecuteNonQuery(""rollback"")出错了。" & vbNewLine & cdb.GetErr)
                Return
            End If

            DisposData(ds)
            Return
        End If

        If cdb.ExecuteNonQuery("commit") Then
            MsgBox("ExecuteNonQuery(""commit"")出错了。" & vbNewLine & cdb.GetErr)
            Return
        End If

        ds.rng.Columns(1).Offset(0, ds.rng.Columns.Count).Value = ds.Result

Exit_sub:
        For i As Integer = 0 To ds.stmtHandleGetID.Length - 1
            cdb.stmtFinalize(ds.stmtHandleGetID(i))
            cdb.stmtFinalize(ds.stmtHandleInsert(i))
        Next

        DisposData(ds)

        Return
    End Sub

    Private Function DisposData(ByRef ds As ImportData) As Boolean
        ds.f.Text = "PriceDB"
        ds.CExcel = Nothing

        Return True
    End Function

    ''' <summary>
    ''' 添加数据，就是要构造出1个完整的表，当前表的每个字段都应该要有对应的值
    ''' 但是如果字段是ID的话，就可以忽略
    ''' 如果字段是xxID的话，就要先去xx表把那个ID找到（方法是先按主键查找，如果没找到，就插入数据再查找）
    ''' 遍历每个字段，如果Pointer为1，表示指向了另外一个table，就递归过去
    ''' </summary>
    ''' <param name="ds"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function AddData(ByRef ds As ImportData, ByRef Table As CSQLite.TableInfo, Optional ByVal IsActivateTable As Boolean = False) As Boolean
        Dim TablePointerArr(Table.Field.Name.Length - 1) As Integer  '当前表的列，对应的Arr中的列

        For i As Integer = 0 To Table.Field.Name.Length - 1
            TablePointerArr(i) = ds.NextCol

            If Table.Field.Pointer(i) = -1 Then
                ds.NextCol += 1
            Else
                '指向了其他的表，递归过去
                If Not AddData(ds, DB_Info.Tables(Table.Field.Pointer(i))) Then
                    Return False
                End If
            End If
        Next

        Dim tableIndex As Integer = DB_Info.DicTableIndex(Table.Name)
        '找到当前表和数组对应的关系后，就开始找ID或者插入数据
        If IsActivateTable Then
            '当前表    就是要插入数据
            ds.f.Text = String.Format("PriceDB:开始添加数据:{0}/{1}", 0, ds.UBoundRow)

            For j As Integer = ds.LBoundRow To ds.UBoundRow
                If j Mod 10 = 0 Then
                    ds.f.Text = String.Format("PriceDB:正在添加数据:{0}/{1}", j, ds.UBoundRow)
                End If

                Dim ret As Integer = 0
                '绑定每一列的值
                For k As Integer = 0 To TablePointerArr.Length - 1
                    Dim tmp As String

                    If Table.Field.Name(k) = "ID" Then
                        Table.LastID += 1
                        tmp = Table.LastID.ToString
                    Else
                        tmp = ds.Arr(j, TablePointerArr(k)).ToString
                    End If

                    If tmp <> "null" Then ' 不绑定默认就是null                    
                        ret = cdb.BindData(k + 1, tmp, Table.Field.Type(k).Name, ds.stmtHandleInsert(tableIndex))
                        If ret Then
                            MsgBox(String.Format("表[{0}]数据Bind出错，出错行号{1}", Table.Name, j) & vbNewLine & cdb.GetErr)
                            Return False
                        End If
                    End If
                Next

                 If Not DoAfterBind(Table.Name, j, ds.stmtHandleInsert(tableIndex)) Then Return False
            Next

            Return True
        Else
            Dim dic As New Dictionary(Of String, Integer) '记录arr中，已经找到了的ID，实际导入的时候，很多都是一样的
            '其他情况就是当前表有xxID字段，需要获取xx表的ID，获取到的ID存放到ds.Arr中，这样最后就为当前表构建了完整信息
            '把ds里的Arr的ID列给获取到
            Dim ret As Integer = 0
            For j As Integer = ds.LBoundRow To ds.UBoundRow
                If j Mod 10 = 0 Then
                    ds.f.Text = String.Format("PriceDB:正在处理表[{0}]ID:{1}/{2}", Table.Name, j, ds.UBoundRow)
                End If

                '首先根据主键尝试获取ID
                Dim strKey As String = GetKeyByArr(ds, Table, j, TablePointerArr)
                If dic.ContainsKey(strKey) Then
                    '将ID放入到数组中
                    ds.Arr(j, TablePointerArr(0)) = dic(strKey)
                Else

                    For k As Integer = 0 To Table.Field.PrimaryKey.Length - 1
                        Dim tmp As String = ds.Arr(j, TablePointerArr(Table.Field.PrimaryKeyIndex(k))).ToString
                        ret = cdb.BindData(k + 1, tmp, Table.Field.Type(Table.Field.PrimaryKeyIndex(k)).Name, ds.stmtHandleGetID(tableIndex))
                        If ret Then
                            MsgBox(String.Format("表[{0}]数据Bind出错，出错行号{1}", Table.Name, j) & vbNewLine & cdb.GetErr)
                            Return False
                        End If
                    Next
                    ret = cdb.GetStep(ds.stmtHandleGetID(tableIndex))

                    Dim id As Integer = 0
                    If ret = 100 Then
                        id = cdb.ColumnValue(0, cdb.ColumnType(0, ds.stmtHandleGetID(tableIndex)), ds.stmtHandleGetID(tableIndex))
                    ElseIf ret = 101 Then
                        id = -1
                    ElseIf ret > 0 Then
                        MsgBox(String.Format("表[{0}]GetID出错，出错行号{1}", Table.Name, j) & vbNewLine & cdb.GetErr)
                        Return False
                    End If

                    ret = cdb.ClearBindings(ds.stmtHandleGetID(tableIndex))
                    If ret Then
                        MsgBox(String.Format("表[{0}]ClearBindings出错，出错行号{1}", Table.Name, j) & vbNewLine & cdb.GetErr)
                        Return False
                    End If

                    ret = cdb.Reset(ds.stmtHandleGetID(tableIndex))
                    If ret Then
                        MsgBox(String.Format("表[{0}]数据Reset出错，出错行号{1}", Table.Name, j) & vbNewLine & cdb.GetErr)
                        Return False
                    End If

                    If id = -1 Then
                        '表的ID增加1
                        Table.LastID += 1
                        id = Table.LastID
                        '将ID放入到数组中
                        ds.Arr(j, TablePointerArr(0)) = id
                        For k As Integer = 0 To TablePointerArr.Length - 1
                            Dim tmp As String = ds.Arr(j, TablePointerArr(k)).ToString
                            If tmp <> "null" Then ' 不绑定默认就是null
                                If tmp <> "null" Then ' 不绑定默认就是null                    
                                    ret = cdb.BindData(k + 1, tmp, Table.Field.Type(k).Name, ds.stmtHandleInsert(tableIndex))
                                    If ret Then
                                        MsgBox(String.Format("表[{0}]数据Bind出错，出错行号{1}", Table.Name, j) & vbNewLine & cdb.GetErr)
                                        Return False
                                    End If
                                End If
                            End If

                            If ret Then
                                MsgBox(String.Format("表[{0}]数据Bind出错，出错行号{1}", Table.Name, j) & vbNewLine & cdb.GetErr)
                                Return False
                            End If
                        Next

                        If Not DoAfterBind(Table.Name, j, ds.stmtHandleInsert(tableIndex)) Then Return False
                        ds.Result(j - 1, 0) &= "Add To [" & Table.Name & "]、"
                    Else
                        '将ID放入到数组中
                        ds.Arr(j, TablePointerArr(0)) = id
                    End If
                    dic(strKey) = id
                End If

            Next
        End If

        Return True
    End Function

    Function DoAfterBind(ByVal tableName As String, ByVal Index As Integer, ByVal stmtHandle As Integer) As Boolean
        Dim ret As Integer

        ret = cdb.GetStep(stmtHandle)
        If ret <> 101 Then
            MsgBox(String.Format("表[{0}]数据insert出错，出错行号{1}", tableName, Index) & vbNewLine & cdb.GetErr)
            Return False
        End If

        ret = cdb.Reset(stmtHandle)
        If ret Then
            MsgBox(String.Format("表[{0}]数据Reset出错，出错行号{1}", tableName, Index) & vbNewLine & cdb.GetErr)
            Return False
        End If

        ret = cdb.ClearBindings(stmtHandle)
        If ret Then
            MsgBox(String.Format("表[{0}]ClearBindings出错，出错行号{1}", tableName, Index) & vbNewLine & cdb.GetErr)
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' 根据主键，获取where的sql 'Select ID From xx Where F1=X AND F2=Y……
    ''' </summary>
    ''' <param name="ds"></param>
    ''' <param name="Table"></param>
    ''' <param name="RowIndex"></param>
    ''' <param name="TablePointerArr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetPrimaryKeySQL(ByRef ds As ImportData, ByRef Table As CSQLite.TableInfo, ByVal RowIndex As Integer, ByRef TablePointerArr() As Integer) As String
        Dim keys(Table.Field.PrimaryKeyIndex.Length - 1) As String
        Dim keysAllBlack As Boolean = True '主键不能全部为空

        For j As Integer = 0 To Table.Field.PrimaryKeyIndex.Length - 1
            'Table.Field.PrimaryKeyIndex(j) 主键对应的table的下标
            Dim index As Integer = Table.Field.PrimaryKeyIndex(j)
            Dim arrCol As Integer = TablePointerArr(index) 'Arr的列

            If ds.Arr(RowIndex, arrCol).ToString <> "" Then
                keysAllBlack = False
            End If

            If Table.Field.Type(index).Name = "DateTime" Then
                keys(j) = Table.Field.Name(index) & "=#" & CDate(ds.Arr(RowIndex, arrCol)).ToString("yyyy-MM-dd") & "#"
            ElseIf Table.Field.Type(index).Name = "String" Then
                keys(j) = ds.Arr(RowIndex, arrCol).ToString
                If InStr(keys(j), """") Then
                    '如果字符串中包含了“"”,就替换掉
                    keys(j) = keys(j).Replace("""", "")
                    ds.Arr(RowIndex, arrCol) = keys(j)
                End If

                If InStr(keys(j), "'") Then
                    keys(j) = keys(j).Replace("'", "")
                    ds.Arr(RowIndex, arrCol) = keys(j)
                End If

                keys(j) = Table.Field.Name(index) & "='" & keys(j) & "'"


            Else
                keys(j) = Table.Field.Name(index) & "=" & ds.Arr(RowIndex, arrCol).ToString
            End If
        Next
        If keysAllBlack Then Return ""

        Return Join(keys, " And ")
    End Function

    ''' <summary>
    ''' 从arr中，根据列来获取主键的信息
    ''' </summary>
    ''' <param name="ds"></param>
    ''' <param name="Table"></param>
    ''' <param name="RowIndex"></param>
    ''' <param name="TablePointerArr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetKeyByArr(ByRef ds As ImportData, ByRef Table As CSQLite.TableInfo, ByVal RowIndex As Integer, ByRef TablePointerArr() As Integer) As String
        Dim key As String = ""

        For j As Integer = 0 To Table.Field.PrimaryKeyIndex.Length - 1
            'Table.Field.PrimaryKeyIndex(j) 主键对应的table的下标
            'TablePointerArr    转化为Arr的下标
            If ds.Arr(RowIndex, TablePointerArr(Table.Field.PrimaryKeyIndex(j))) IsNot Nothing Then
                key &= ds.Arr(RowIndex, TablePointerArr(Table.Field.PrimaryKeyIndex(j))).ToString
                key &= CONTEXT_FH
            End If
        Next

        Return key
    End Function

    Private Function GetKeyByDt(ByRef dt As DataTable, ByRef Table As TableInfo, ByVal RowIndex As Integer) As String
        Dim key As String = ""

        For j As Integer = 0 To Table.Field.PrimaryKeyIndex.Length - 1
            key &= dt.Rows(RowIndex).Item(Table.Field.PrimaryKeyIndex(j)).ToString
            key &= CONTEXT_FH
        Next

        Return key
    End Function

    ''' <summary>
    ''' 检查标题，按扩展的标题检查
    ''' </summary>
    ''' <param name="ds"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckField(ByRef ds As ImportData) As Boolean
        If ds.UBoundCol - ds.LBoundCol + 1 <> DB_Info.Tables(DB_Info.ActivateTableIndex).ExtendField.ExtendFieldName.Length Then
            MsgBox("Excel的列数不对。")
            Return False
        End If

        For i As Integer = ds.LBoundCol To ds.UBoundCol
            If ds.Arr(ds.LBoundRow - 1, i).ToString <> DB_Info.Tables(DB_Info.ActivateTableIndex).ExtendField.ExtendFieldName(i - ds.LBoundCol) Then
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
    Private Function GetRngFromExcel(ByRef ds As ImportData) As Object
        ds.CExcel = New CExcel
        ds.CExcel.GetExcel()
        Dim rng As Object = ds.CExcel.GetRng()

        Return rng
    End Function
End Module

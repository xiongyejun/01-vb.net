Public Class FMain

#Region "Form"
    Private F_ExcuteSQL As New FExcuteSQL

    Private WithEvents btnSearch As Button
    Private WithEvents btnAdd As Button
    Private WithEvents btnUpdate As Button
    Private WithEvents btnDelete As Button


    Private WithEvents btnFirst As Button
    Private WithEvents btnPre As Button
    Private WithEvents btnNext As Button
    Private WithEvents btnLast As Button

    Private WithEvents dgv As DataGridView
    ''' <summary>
    ''' dgv控件关联的数据
    ''' </summary>
    ''' <remarks></remarks>
    Private dt As DataTable
    Private cpage As CPages
    Private sqlSearch As String

    Private ArrMeSet() As String = New String() {"WindowState", "Height", "Width", "Left", "Top"}

    Private Sub FMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If DicSet IsNot Nothing Then
            DicSet.Clear()

            For Each Str As String In ArrMeSet
                DicSet(Str) = CallByName(Me, Str, CallType.Get)
            Next

            DicSet("DBPath") = DB_Info.Path

            WriteSet()
        End If
    End Sub

    Private Sub FMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim iLeft As Integer = 5
        Dim iTop As Integer = 25

        btnSearch = New Button
        With btnSearch
            .Left = iLeft
            .Top = iTop
            .Text = "查询"
            .Width = 50
            .Height = 50
        End With
        Me.Controls.Add(btnSearch)

        dgv = New DataGridView
        With dgv
            .Left = iLeft + btnSearch.Height
            .Top = iTop
            .ScrollBars = ScrollBars.Both
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .AllowUserToResizeRows = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        End With
        Me.Controls.Add(dgv)

        iTop += btnSearch.Height
        btnAdd = New Button
        With btnAdd
            .Left = iLeft
            .Top = iTop
            .Text = "添加"
            .Width = 50
            .Height = 50
        End With
        Me.Controls.Add(btnAdd)

        iTop += btnSearch.Height
        btnUpdate = New Button
        With btnUpdate
            .Left = iLeft
            .Top = iTop
            .Text = "更新"
            .Width = 50
            .Height = 50
        End With
        Me.Controls.Add(btnUpdate)

        iTop += btnUpdate.Height
        btnDelete = New Button
        With btnDelete
            .Left = iLeft
            .Top = iTop
            .Text = "删除"
            .Width = 50
            .Height = 50
            .Enabled = False
        End With
        Me.Controls.Add(btnDelete)

        iTop += btnDelete.Height
        btnFirst = New Button
        With btnFirst
            .Left = iLeft
            .Top = iTop
            .Text = "第1页"
            .Width = 50
            .Height = 50
            '.Enabled = False
        End With
        Me.Controls.Add(btnFirst)
        iTop += btnFirst.Height

        btnPre = New Button
        With btnPre
            .Left = iLeft
            .Top = iTop
            .Text = "上一页"
            .Width = 50
            .Height = 50
            '.Enabled = False
        End With
        Me.Controls.Add(btnPre)

        iTop += btnPre.Height
        btnNext = New Button
        With btnNext
            .Left = iLeft
            .Top = iTop
            .Text = "下一页"
            .Width = 50
            .Height = 50
            '.Enabled = False
        End With
        Me.Controls.Add(btnNext)

        iTop += btnNext.Height
        btnLast = New Button
        With btnLast
            .Left = iLeft
            .Top = iTop
            .Text = "最后1页"
            .Width = 50
            .Height = 50
            '.Enabled = False
        End With
        Me.Controls.Add(btnLast)

        Me.Text = "PriceDB"

        If DicSet IsNot Nothing Then
            If Not DicSet.ContainsKey("WindowState") Then
                DicSet("WindowState") = FormWindowState.Maximized
            End If
            For Each Str As String In ArrMeSet
                If DicSet.ContainsKey(Str) Then
                    CallByName(Me, Str, CallType.Set, DicSet(Str))
                End If
            Next
        End If

        '添加表菜单
        SetDBPath()

    End Sub

    Private Sub FMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Try
            dtvResize()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub dtvResize()
        With dgv
            .Height = Me.Height - .Top - 60
            .Width = Me.Width - .Left - 10
        End With
    End Sub

#End Region

#Region "dgv"
    ''' <summary>
    ''' 添加行号
    ''' </summary>
    Private Sub dgv_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles dgv.RowPostPaint
        'Dim rect As Drawing.Rectangle = New Drawing.Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, dgvTest.RowHeadersWidth - 4, e.RowBounds.Height)
        'TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), dgvTest.RowHeadersDefaultCellStyle.Font, rect, dgvTest.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
        Using b As SolidBrush = New SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                                  dgv.DefaultCellStyle.Font,
                                  b,
                                  e.RowBounds.Location.X,
                                  e.RowBounds.Location.Y + 4)

        End Using
    End Sub

    Private Sub dgv_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgv.SelectionChanged
        btnDelete.Enabled = True
    End Sub
#End Region

#Region "Menu File"
    Private Sub 选择文件SToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemSelectFile.Click
        If Not SelectDB() Then Return
        InitDBInfo()
        SetDBPath()
    End Sub

    Function SetDBPath() As Integer
        If DB_Info.Path <> "" Then
            '把数据库的表都添加到菜单上
            Me.MenuItemTables.DropDownItems.Clear()
            For i As Integer = 0 To DB_Info.Tables.Length - 1
                Dim tmp As ToolStripItem = Me.MenuItemTables.DropDownItems.Add(DB_Info.Tables(i).Name)
                tmp.Tag = i '记录表在db_info里tables的下标
                AddHandler tmp.Click, AddressOf MenuItem_Click
            Next
            '新选择了数据库文件就重新设置状态栏
            Me.status.Text = "DBPath：" & DB_Info.Path
            Me.statusTable.Text = " Table："
            Me.MenuItemActions.Enabled = True
        End If

        Return 1
    End Function

    Private Sub MenuItemBackup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemBackup.Click
        BackupFile(DB_Info.Path)
    End Sub

    ''' <summary>
    ''' 执行一个SQL
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub MenuItemExcuteSQL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemExcuteSQL.Click
        F_ExcuteSQL.ShowDialog(Me)
        Dim strSQL As String = F_ExcuteSQL.GetSQL()
        If strSQL.Length Then
            Dim ret As Integer = cdb.ExecuteNonQuery(strSQL)
            If ret Then
                MsgBox(cdb.GetErr)
            End If
        End If
        F_ExcuteSQL.Hide()
    End Sub

    Private Sub MenuItemOpenDBPath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemOpenDBPath.Click
        Dim strFolder As String = DB_Info.Path.Substring(0, DB_Info.Path.LastIndexOfAny("\"))
        System.Diagnostics.Process.Start(strFolder)
    End Sub

    Private Sub 退出XToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemQuit.Click
        Me.Close()
    End Sub
#End Region

#Region "Tables"
    ''' <summary>
    ''' 每个[表]都作为1个下拉按钮，单击的时候选择该表
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub MenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim tmp As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        '为了实现单选功能，先都设置为不选中，
        For Each t As ToolStripMenuItem In Me.MenuItemTables.DropDownItems
            t.Checked = False
        Next
        '再选中当前的
        tmp.Checked = True
        DB_Info.ActivateTableIndex = CLng(tmp.Tag)
        DB_Info.ActivateTable = tmp.Text
        Me.statusTable.Text = " Table：" & DB_Info.ActivateTable
        Me.statusFields.Text = " Fields：" & Join(DB_Info.Tables(DB_Info.ActivateTableIndex).Field.Name, ",") & " PrimaryKey:" & Join(DB_Info.Tables(DB_Info.ActivateTableIndex).Field.PrimaryKey, ",")
        '把显示的数据也清空
        dgv.DataSource = Nothing
    End Sub
#End Region

#Region "Button"
    Private Sub btnSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        '是否设置了数据库
        If Strings.Len(DB_Info.Path) Then
            '是否选择了表格
            If Strings.Len(DB_Info.ActivateTable) Then
                Dim strSql As String = ""
                Dim strCondition As String = ""
                '是否需要展开ID
                If MenuItemExtendID.Checked Then
                    strSql = DB_Info.Tables(DB_Info.ActivateTableIndex).SqlExtend ' GetSql(DB_Info.Tables(DB_Info.ActivateTableIndex))
                Else
                    strSql = "Select * From [" & DB_Info.ActivateTable & "]"
                End If

                '是否需要设置查询条件
                If MenuItemSearchCondition.Checked Then
                    Dim F_SearchCondition As New FSearchCondition

                    F_SearchCondition = New FSearchCondition
                    '是否需要展开ID
                    If MenuItemExtendID.Checked Then
                        F_SearchCondition.DBFields = DB_Info.Tables(DB_Info.ActivateTableIndex).ExtendField.ExtendFieldName
                        F_SearchCondition.DBFieldsType = DB_Info.Tables(DB_Info.ActivateTableIndex).ExtendField.ExtendFieldType
                    Else
                        F_SearchCondition.DBFields = DB_Info.Tables(DB_Info.ActivateTableIndex).Field.Name
                        F_SearchCondition.DBFieldsType = DB_Info.Tables(DB_Info.ActivateTableIndex).Field.Type
                        strSql &= " Where 1=1 "
                    End If

                    F_SearchCondition.ShowDialog(Me)
                    strCondition = F_SearchCondition.ReturnValue
                    F_SearchCondition.Hide()
                End If
                '如果是空，可能是没进行过设置。也有可能是取消了，如果是取消了就不进行查询
                If Not (strCondition = "" AndAlso MenuItemSearchCondition.Checked) Then
                    strSql &= strCondition
                    '默认是按照主键排序的，如果有ID，让表按照ID来排序
                    If DB_Info.Tables(DB_Info.ActivateTableIndex).bHasID Then strSql &= " Order By " & DB_Info.Tables(DB_Info.ActivateTableIndex).Name & ".ID"
                    sqlSearch = strSql
                    Dim tmpCounts As Integer = cdb.GetColZeroValue(String.Format("select count(*) from ({0})", sqlSearch))
                    cpage = New CPages(tmpCounts, PAGE_NUM)
                    dt = cdb.ExecuteQuery(strSql & cpage.GetLimitOffset)
                    SetDGV()
                End If

            Else
                MsgBox("请选择表。")
            End If
        Else
            MsgBox("请选择数据库。")
        End If
    End Sub

    Private Sub btnAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Dim f_add As New FAdd

        f_add.DBFields = DB_Info.Tables(DB_Info.ActivateTableIndex).Field.Name
        f_add.DBFieldsType = DB_Info.Tables(DB_Info.ActivateTableIndex).Field.Type
        f_add.ShowDialog(Me)

    End Sub

    Private Sub btnUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Dim arr(20) As Integer
        Dim parr As Integer = 0

        For i As Integer = 0 To dt.Rows.Count - 1
            Select Case dt.Rows(i).RowState
                Case Data.DataRowState.Modified
                    arr(parr) = i
                    parr += 1
            End Select
        Next

        If parr Then
            Dim icount As Integer = DB_Info.Tables(DB_Info.ActivateTableIndex).Field.Name.Length - 1 'ID不需要，根据的就是ID来更新
            Dim values(icount - 1) As String
            For i As Integer = 0 To icount - 1
                values(i) = DB_Info.Tables(DB_Info.ActivateTableIndex).Field.Name(i + 1) & "=?"
            Next


            If cdb.ExecuteNonQuery("begin") Then
                MsgBox("ExecuteNonQuery(""begin"")出错了。" & vbNewLine & cdb.GetErr)
                Return
            End If

            Dim stmtHandle As Integer = 0
            Dim ret As Integer = 0
            ret = cdb.Prepare(String.Format("update {0} set {1} where ID=?", DB_Info.ActivateTable, Join(values, ",")), stmtHandle)
            If ret Then
                MsgBox(cdb.GetErr)
                cdb.ExecuteNonQuery("rollback")
                Return
            End If

            For i As Integer = 0 To parr - 1
                Dim j As Integer = 0
                For j = 1 To DB_Info.Tables(DB_Info.ActivateTableIndex).Field.Name.Length - 1
                    ret = cdb.BindData(j, dt.Rows(arr(i)).Item(j), DB_Info.Tables(DB_Info.ActivateTableIndex).Field.Type(j).Name, stmtHandle)
                    If ret Then
                        MsgBox(String.Format("数据Bind出错，出错行号{0}", i) & vbNewLine & cdb.GetErr)
                        Return
                    End If
                Next
                '最后1个是ID
                ret = cdb.BindData(j, dt.Rows(arr(i)).Item(0), DB_Info.Tables(DB_Info.ActivateTableIndex).Field.Type(0).Name, stmtHandle)
                If ret Then
                    MsgBox(String.Format("数据Bind出错，出错行号{0}", i) & vbNewLine & cdb.GetErr)
                    Return
                End If

                If Not DoAfterBind(DB_Info.ActivateTable, i, stmtHandle) Then cdb.ExecuteNonQuery("rollback")
            Next

            If cdb.ExecuteNonQuery("commit") Then
                MsgBox("ExecuteNonQuery(""commit"")出错了。" & vbNewLine & cdb.GetErr)
                Return
            End If
            dt.AcceptChanges()
        End If

    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim index As Integer = dgv.SelectedCells(0).RowIndex
        Dim ret As Integer = cdb.ForeignKeys(True)
        If ret Then
            MsgBox(cdb.GetErr)
            Return
        End If

        Dim strSql As String = String.Format("delete from {0} where ID={1}", DB_Info.ActivateTable, dt.Rows(index).Item("ID"))
        ret = cdb.ExecuteNonQuery(strSql)
        If ret Then
            MsgBox(cdb.GetErr)
            Return
        End If

        cdb.ForeignKeys(False)
    End Sub

    Private Sub btnPre_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPre.Click
        btnPre.Enabled = cpage.Pre
        dt = cdb.ExecuteQuery(sqlSearch & cpage.GetLimitOffset)
        SetDGV()

        btnNext.Enabled = True
    End Sub

    Private Sub btnNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNext.Click
        btnNext.Enabled = cpage.NextP
        dt = cdb.ExecuteQuery(sqlSearch & cpage.GetLimitOffset())
        SetDGV()

        btnPre.Enabled = True
    End Sub

    Private Sub btnFirst_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFirst.Click
        cpage.First()
        dt = cdb.ExecuteQuery(sqlSearch & cpage.GetLimitOffset())
        SetDGV()
    End Sub

    Private Sub btnLast_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLast.Click
        cpage.Last()
        dt = cdb.ExecuteQuery(sqlSearch & cpage.GetLimitOffset())
        SetDGV()
    End Sub
#End Region

#Region "Option"
    Private Sub MenuItemExtendID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemExtendID.Click
        MenuItemExtendID.Checked = Not MenuItemExtendID.Checked
        btnUpdate.Enabled = Not MenuItemExtendID.Checked
    End Sub
    Private Sub MenuItemSearchCondition_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemSearchCondition.Click
        MenuItemSearchCondition.Checked = Not MenuItemSearchCondition.Checked
    End Sub
#End Region

#Region "Menu Import"
    '导入
    Private Sub MenuItemImportFromExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemImportFromExcel.Click
        Me.Hide()

        Try
            ImportFromExcel(Me)
        Catch ex As Exception
            MsgBox("MenuItemImportFromExcel_Click出错：" & vbNewLine & ex.Message)
        End Try

        Me.Show()
    End Sub
#End Region

#Region "Menu Struct"
    Private Sub MenuItemGetField_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemGetField.Click
        Me.Hide()
        Dim c_excel As New CExcel
        c_excel.GetExcel()
        Dim rng As Object = c_excel.GetRng()

        If rng Is Nothing Then
            Me.Show()
            Return
        End If
        c_excel = Nothing

        rng.resize(1, DB_Info.Tables(DB_Info.ActivateTableIndex).ExtendField.ExtendFieldName.Length).value = DB_Info.Tables(DB_Info.ActivateTableIndex).ExtendField.ExtendFieldName
        '标识表的主键为红色底色
        SignPrimaryKey(rng, 0, DB_Info.Tables(DB_Info.ActivateTableIndex))
        MsgBox("标红项为主键，1个表中的主键不能【全部】为空。")
        Me.Show()
    End Sub

    ''' <summary>
    ''' 标识表的主键为红色底色
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="iOffset"></param>
    ''' <param name="Table"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SignPrimaryKey(ByRef rng As Object, ByVal iOffset As Integer, ByRef Table As CSQLite.TableInfo) As Boolean
        Dim extendCol As Integer = 0 '因为有扩展表需要偏移的列，那种xxID的字段

        Dim iCol As Integer = 0
        '循环主键下标
        For i As Integer = 0 To Table.Field.PrimaryKeyIndex.Length - 1
            '0-主键1；主键1+1-主键2；主键2+1-主键3……
            For j As Integer = iCol To Table.Field.PrimaryKeyIndex(i)
                If Table.Field.Pointer(j) > -1 Then
                    '指向了其他的表, 递归过去
                    SignPrimaryKey(rng, iOffset + extendCol + j, DB_Info.Tables(Table.Field.Pointer(j)))
                    '偏移指向的那个表的扩展字段的长度
                    extendCol += DB_Info.Tables(Table.Field.Pointer(j)).ExtendField.ExtendFieldName.Length
                    '本身也要占1个位置
                    extendCol -= 1
                End If
            Next
            If Table.Field.Pointer(Table.Field.PrimaryKeyIndex(i)) = -1 Then
                rng.Offset(0, iOffset + extendCol + Table.Field.PrimaryKeyIndex(i)).Interior.Color = 255
            End If
            '主键n+1
            iCol = Table.Field.PrimaryKeyIndex(i) + 1
        Next

        '主键后的也要处理，有可能还有其他表的主键
        For i As Integer = iCol To Table.Field.Name.Length - 1
            If Table.Field.Pointer(i) = -1 Then

            Else
                SignPrimaryKey(rng, iOffset + extendCol + i, DB_Info.Tables(Table.Field.Pointer(i)))
                extendCol += DB_Info.Tables(Table.Field.Pointer(i)).Field.Name.Length
                extendCol -= 1 '本身也要占1个位置
            End If
        Next

        Return True
    End Function

    Private Sub MenuItemCopyFields_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemCopyFields.Click
        If DB_Info.ActivateTableIndex > -1 Then
            My.Computer.Clipboard.SetText(Join(DB_Info.Tables(DB_Info.ActivateTableIndex).Field.Name, ","))
        Else
            MsgBox("请先选择表。")
        End If
    End Sub

    Private Sub MenuItemCopyExtendFields_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemCopyExtendFields.Click
        If DB_Info.ActivateTableIndex > -1 Then
            My.Computer.Clipboard.SetText(Join(DB_Info.Tables(DB_Info.ActivateTableIndex).ExtendField.ExtendFieldName, ","))
        Else
            MsgBox("请先选择表。")
        End If
    End Sub


    Private Sub MenuItemCopyExtendSQL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemCopyExtendSQL.Click
        If DB_Info.ActivateTableIndex > -1 Then
            My.Computer.Clipboard.SetText(DB_Info.Tables(DB_Info.ActivateTableIndex).SqlExtend)
        Else
            MsgBox("请先选择表。")
        End If
    End Sub
#End Region

    Sub SetDGV()
        If dt Is Nothing Then Return

        dt.AcceptChanges()
        dgv.DataSource = dt
        If Not Me.MenuItemExtendID.Checked Then
            dgv.Columns("ID").ReadOnly = True
            For i As Integer = 0 To DB_Info.Tables(DB_Info.ActivateTableIndex).Field.PrimaryKeyIndex.Length - 1
                With dgv.Columns(DB_Info.Tables(DB_Info.ActivateTableIndex).Field.PrimaryKeyIndex(i))
                    .ReadOnly = True
                    .HeaderCell.ToolTipText = "Primary Key"
                End With
            Next
        End If

        dgv.AutoResizeColumns()
    End Sub
End Class

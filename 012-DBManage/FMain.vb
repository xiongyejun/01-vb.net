Public Class FMain

#Region "Form"
    Private WithEvents btnSearch As Button
    Private WithEvents btnAdd As Button
    Private WithEvents btnUpdate As Button
    Private WithEvents btnDelete As Button

    Private WithEvents dgv As DataGridView
    Private dt As New DataTable

    Private ArrMeSet() As String = New String() {"WindowState", "Height", "Width", "Left", "Top"}

    Private Sub FMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DicSet.Clear()

        For Each Str As String In ArrMeSet
            DicSet(Str) = CallByName(Me, Str, CallType.Get)
        Next

        DicSet("DBPath") = DB_Info.Path

        WriteSet()
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

        Me.Text = "PriceDB"

        '读取设置，并应用设置
        ReadSet()
        If Not DicSet.ContainsKey("WindowState") Then
            DicSet("WindowState") = FormWindowState.Maximized
        End If
        For Each Str As String In ArrMeSet
            If DicSet.ContainsKey(Str) Then
                CallByName(Me, Str, CallType.Set, DicSet(Str))
            End If
        Next

        If DicSet.ContainsKey("DBPath") Then
            DB_Info.Path = DicSet("DBPath")
            InitDBInfo()
            SetDBPath()
        End If
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
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture), _
                                  dgv.DefaultCellStyle.Font, _
                                  b, _
                                  e.RowBounds.Location.X, _
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
            Me.statusDBPath.Text = "DBPath：" & DB_Info.Path
            Me.statusTable.Text = " Table："
            Me.MenuItemActions.Enabled = True
        End If

        Return 1
    End Function

    Private Sub MenuItemBackup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemBackup.Click
        BackupFile(DB_Info.Path)
    End Sub

    Private Sub 退出XToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemQuit.Click
        Me.Close()
    End Sub
#End Region

#Region "Tables"
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
                    Dim f As New FSearchCondition
                    '是否需要展开ID
                    If MenuItemExtendID.Checked Then
                        f.DBFields = DB_Info.Tables(DB_Info.ActivateTableIndex).ExtendField.ExtendFieldName
                        f.DBFieldsType = DB_Info.Tables(DB_Info.ActivateTableIndex).ExtendField.ExtendFieldType
                    Else
                        f.DBFields = DB_Info.Tables(DB_Info.ActivateTableIndex).Field.Name
                        f.DBFieldsType = DB_Info.Tables(DB_Info.ActivateTableIndex).Field.Type
                        strSql &= " Where 1=1 "
                    End If
                   
                    f.ShowDialog(Me)

                    strCondition = f.ReturnValue
                    f.Close()
                End If
                '如果是空，可能是没进行过设置。也有可能是取消了，如果是取消了就不进行查询
                If Not (strCondition = "" AndAlso MenuItemSearchCondition.Checked) Then
                    strSql &= strCondition
                    dt = MFunc.DoSearch(DB_Info.Path, strSql)
                    dgv.DataSource = dt
                    dgv.AutoResizeColumns()
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
        Dim c_ado As New CADO(DB_Info.Path)
        c_ado.UpdateData(dt, DB_Info.ActivateTable)
        c_ado = Nothing
    End Sub


    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim index As Integer = dgv.SelectedCells(0).RowIndex
        Dim strSql As String = ""
        Dim table As TableInfo = DB_Info.Tables(DB_Info.ActivateTableIndex)
        Dim c_ado As New CADO(DB_Info.Path)

        If dgv.Columns(0).HeaderText = "ID" Then
            Dim oldID As String = CStr(dgv.Rows(index).Cells(0).Value)
            '有ID的要判断是否有引用了他的ID的表，先替换，再删除
            If table.bUseMyIdTables Then
                '选择一个替换的ID
                Dim newID As String = ""
                MsgBox("删除前请先选择一个代替者。")
                Dim f As New FSelectID

                f.SetFormText = "选择" & table.Name & "ID"
                f.ShowDialog(Me)

                Dim tmp As Integer = f.ReturnValue
                If tmp > -1 Then newID = tmp.ToString
                f.Close()

                If newID = "" Then
                    c_ado = Nothing
                    Return
                ElseIf newID = oldID Then
                    MsgBox("不能选择自己。")
                    Return
                Else
                    '更新其他表格的ID
                    For i As Integer = 0 To table.UseMyIdTables.Length - 1
                        strSql = "Update [" & DB_Info.Tables(table.UseMyIdTables(i)).Name & "] Set " & table.Name & "ID=" & newID & " Where " & table.Name & "ID=" & oldID
                        c_ado.StrSql = strSql
                        c_ado.ExcuteSql()
                    Next
                End If
            End If
            strSql = "ID=" & oldID
        Else
            '没有ID字段的，没有其他表引用，直接删除
            Dim arr() As Integer = table.Field.PrimaryKeyIndex
            Dim sqlArr(arr.Length - 1) As String
            For i As Integer = 0 To arr.Length - 1
                If table.Field.Type(arr(i)).Name = "String" Then
                    sqlArr(i) = table.Field.Name(arr(i)) & "='" & dgv.Rows(index).Cells(arr(i)).Value & "'"
                ElseIf table.Field.Type(arr(i)).Name = "DateTime" Then
                    sqlArr(i) = table.Field.Name(arr(i)) & "=#" & dgv.Rows(index).Cells(arr(i)).Value & "#"
                Else
                    sqlArr(i) = table.Field.Name(arr(i)) & "=" & dgv.Rows(index).Cells(arr(i)).Value
                End If
            Next
            strSql = Join(sqlArr, " And ")
        End If

        c_ado.StrSql = "Delete * From [" & DB_Info.ActivateTable & "] Where " & strSql
        c_ado.ExcuteSql()
        c_ado = Nothing
    End Sub
#End Region

#Region "Option"
    Private Sub MenuItemOption_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemExtendID.Click, MenuItemSearchCondition.Click
        sender.Checked = Not sender.Checked
        btnUpdate.Enabled = Not MenuItemExtendID.Checked
    End Sub
#End Region

#Region "Menu Import"
    '导入
    Private Sub MenuItemImportFromExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemImportFromExcel.Click
        Me.Hide()

        ImportFromExcel()

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

        Me.Show()
    End Sub
#End Region

    Private Sub MenuItemTmp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemTmp.Click
        Debug.Print(DB_Info.Tables(DB_Info.ActivateTableIndex).SqlExtend)
    End Sub


End Class

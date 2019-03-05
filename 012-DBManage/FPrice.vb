Public Class FPrice
    Private Enum PD
        项目明细
        价格形式
        价格时间时间
        价格时间备注
    End Enum

    Private Structure PriceData
        Dim TableName As String
        Dim lbName As String
        Dim tbReadOnly As Boolean
        Dim ReturnCol As Integer
        Dim ID As Integer
        Dim tb As TextBox
    End Structure
    Private pds(3) As PriceData

    Private dt As DataTable
    Private WithEvents dgv As DataGridView
    Private WithEvents btnOK As Button
    Private WithEvents btnCancel As Button
    Private WithEvents btnFromExcel As Button

    Private Sub FPrice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim pdCount As Integer = 0
        pds(pdCount).lbName = "项目明细.代码"
        pds(pdCount).TableName = "项目明细"
        pds(pdCount).tbReadOnly = True
        pds(pdCount).ReturnCol = 1
        pdCount += 1

        pds(pdCount).lbName = "价格形式.名称"
        pds(pdCount).TableName = "价格形式"
        pds(pdCount).tbReadOnly = True
        pds(pdCount).ReturnCol = 1
        pdCount += 1

        pds(pdCount).lbName = "价格时间.时间"
        pds(pdCount).tbReadOnly = False
        pdCount += 1

        pds(pdCount).lbName = "价格时间.备注"
        pds(pdCount).tbReadOnly = False
        pdCount += 1
        '以上可以确定一个价格时间

        Const LB_WIDTH As Integer = 100
        Const TB_WIDTH As Integer = 300
        Dim iLeft As Integer = 5
        Dim iTop As Integer = 5
        For i As Integer = 0 To pdCount - 1
            Dim lb As Label = New Label
            With lb
                .Text = pds(i).lbName
                .Left = iLeft
                .Width = LB_WIDTH
                .Top = iTop
            End With

            pds(i).tb = New TextBox
            With pds(i).tb
                .ReadOnly = pds(i).tbReadOnly
                .Tag = i
                .Left = iLeft + LB_WIDTH
                .Top = iTop
                .Width = TB_WIDTH

                If pds(i).lbName = "价格时间.时间" Then .Text = Format(Now(), "yyyy-MM-dd")
            End With

            If pds(i).tbReadOnly Then
                AddHandler pds(i).tb.Click, AddressOf tb_Click
            End If

            iTop += pds(i).tb.Height
            iTop += 5
            Me.Controls.Add(lb)
            Me.Controls.Add(pds(i).tb)
        Next

        dt = New DataTable
        Dim c As DataColumn = dt.Columns.Add()
        c.ColumnName = "序号"
        c.DataType = Type.GetType("System.String")

        c = dt.Columns.Add()
        c.ColumnName = "项目"
        c.DataType = Type.GetType("System.String")

        c = dt.Columns.Add()
        c.ColumnName = "价格"
        c.DataType = Type.GetType("System.Double")

        c = dt.Columns.Add()
        c.ColumnName = "单位"
        c.DataType = Type.GetType("System.String")

        c = dt.Columns.Add()
        c.ColumnName = "备注"
        c.DataType = Type.GetType("System.String")
        Dim r As DataRow = dt.NewRow
        r.Item("序号") = "1"
        r.Item("项目") = "价格"
        r.Item("单位") = "元"

        dt.Rows.Add(r)


        dgv = New DataGridView
        With dgv
            .Left = iLeft
            .Top = iTop
            .ScrollBars = ScrollBars.Both
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .AllowUserToResizeRows = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .Width = 600
            .Height = 300
        End With
        iTop += dgv.Height
        Me.Controls.Add(dgv)
        dgv.DataSource = dt

        iTop += 10
        btnOK = New Button
        With btnOK
            .Text = "确定"
            .Width = (LB_WIDTH + TB_WIDTH) / 3
            .Left = iLeft
            iLeft += .Width
            .Top = iTop
        End With
        Me.Controls.Add(btnOK)

        btnCancel = New Button
        With btnCancel
            .Text = "取消"
            .Width = (LB_WIDTH + TB_WIDTH) / 3
            .Left = iLeft
            iLeft += .Width
            .Top = iTop
        End With
        Me.Controls.Add(btnCancel)

        btnFromExcel = New Button
        With btnFromExcel
            .Text = "从Excel导入"
            .Width = (LB_WIDTH + TB_WIDTH) / 3
            .Left = iLeft
            iLeft += .Width
            .Top = iTop
        End With
        Me.Controls.Add(btnFromExcel)
        iTop += btnFromExcel.Height

        Me.Width = dgv.Width + 20
        Me.Height = iTop + 50

        SetFromPos(Me)
    End Sub

    Private Sub tb_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim f As New FSelectItem
        Dim t As TextBox = CType(sender, TextBox)
        Dim index As Integer = Int(t.Tag)

        f.SetFormText = "选择" & pds(index).lbName
        f.TableName = pds(index).TableName
        f.ReturnCol = pds(index).ReturnCol
        f.ShowDialog(Me)

        Dim tmp As String = f.ReturnValue
        If tmp <> "" Then t.Text = tmp
        pds(index).ID = f.ReturnZeroCol

        f.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If pds(PD.项目明细).ID < 1 Then
            MsgBox("请先选择[项目明细]")
            Return
        End If
        If pds(PD.价格形式).ID < 1 Then
            MsgBox("请先选择[价格形式]")
            Return
        End If

        Dim 价格构成项ID(dt.Rows.Count - 1) As Integer
        Dim bHasPrice As Boolean = False
        '检查是否有[价格构成项]中的[价格]，并检查项目是否都存在
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim str As String = dt.Rows(i).Item("项目")
            价格构成项ID(i) = cdb.GetColZeroValue(String.Format("select id from 价格构成项 where 名称='{0}'", str))
            If 价格构成项ID(i) = -1 Then
                MsgBox("请先添加[价格构成项]" & str)
                Return
            End If

            If str = "价格" Then bHasPrice = True
        Next
        If Not bHasPrice Then
            MsgBox("必须要有[价格]这个项目。")
            Return
        End If


        Dim ret As Integer = 0
        Dim priceTimeID As Integer = 0

        priceTimeID = GetId("价格时间")
        If priceTimeID = -1 Then Return

        If cdb.ExecuteNonQuery("begin") Then
            MsgBox("ExecuteNonQuery(""begin"")出错了。" & vbNewLine & cdb.GetErr)
            Return
        End If
        '先添加到[价格时间]
        'ID,项目明细ID,时间,价格形式ID,inserttime,备注
        Dim values As String = String.Format("{0},{1},'{2}',{3},'{4}'", priceTimeID, pds(PD.项目明细).ID, pds(PD.价格时间时间).tb.Text, pds(PD.价格形式).ID, pds(PD.价格时间备注).tb.Text)
        ret = cdb.ExecuteNonQuery(String.Format("insert into {0} ({1}) values ({2})", "价格时间", "ID,项目明细ID,时间,价格形式ID,备注", values))
        If ret Then
            MsgBox("insert into [价格时间] 出错了。" & vbNewLine & cdb.GetErr)
            cdb.ExecuteNonQuery("rollback")
            Return
        End If

        Dim 价格数据ID As Integer = 0
        价格数据ID = GetId("价格数据")
        If 价格数据ID = -1 Then Return
        '再添加到[价格数据]
        'ID, 价格时间ID, 序号, 价格构成项ID, 价格, 单位, inserttime, 备注
        For i As Integer = 0 To dt.Rows.Count - 1
            values = String.Format("{0},{1},'{2}',{3},{4},'{5}','{6}'", 价格数据ID, priceTimeID, dt.Rows(i).Item("序号"), 价格构成项ID(i), dt.Rows(i).Item("价格"), dt.Rows(i).Item("单位"), dt.Rows(i).Item("备注"))
            ret = cdb.ExecuteNonQuery(String.Format("insert into {0} ({1}) values ({2})", "价格数据", "ID, 价格时间ID, 序号, 价格构成项ID, 价格, 单位, 备注", values))
            If ret Then
                MsgBox("insert into [价格数据] 出错了。" & vbNewLine & cdb.GetErr)
                cdb.ExecuteNonQuery("rollback")
                Return
            End If
            价格数据ID += 1
        Next

        If cdb.ExecuteNonQuery("commit") Then
            MsgBox("ExecuteNonQuery(""commit"")出错了。" & vbNewLine & cdb.GetErr)
            Return
        End If

        MsgBox("OK")
    End Sub

    Private Function GetId(ByVal tableName As String) As Integer
        DB_Info.Tables(DB_Info.ActivateTableIndex).LastID = cdb.GetColZeroValue(String.Format("select max(ID) from {0}", tableName))
        If DB_Info.Tables(DB_Info.ActivateTableIndex).LastID = -1 Then
            MsgBox(DB_Info.Tables(DB_Info.ActivateTableIndex).Name & " LastID获取出错了。" & vbNewLine & cdb.GetErr)
            Return -1
        End If
        Return DB_Info.Tables(DB_Info.ActivateTableIndex).LastID + 1
    End Function


    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnFromExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFromExcel.Click
        Dim ex As CExcel = New CExcel
        ex.GetExcel()

        Me.Hide()
        Dim rng As Object = ex.GetRng
        Me.Show()

        If rng Is Nothing Then Return

        Dim Arr(,) As Object = rng.value
        If Arr.GetUpperBound(1) <> dt.Columns.Count Then
            MsgBox("必须是5列数据")
            Return
        End If

        dt.Clear()
        For i As Integer = 1 To Arr.GetUpperBound(0) - 1 '跳过标题
            Dim r As DataRow = dt.NewRow
            For j As Integer = 0 To dt.Columns.Count - 1
                r.Item(j) = Arr(i + 1, j + 1)
            Next
            dt.Rows.Add(r)
        Next
    End Sub
End Class
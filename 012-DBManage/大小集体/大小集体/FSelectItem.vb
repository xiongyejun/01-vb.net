''' <summary>
''' 针对*ID这样的字段，需要打开对应的table的数据来方便选择id
''' </summary>
''' <remarks></remarks>

Public Class FSelectItem
    Private WithEvents btnSetConditon As Button '设置查询条件

    Private WithEvents btnFirst As Button
    Private WithEvents btnPre As Button
    Private WithEvents btnNext As Button
    Private WithEvents btnLast As Button

    Private cpage As CPages
    Private sqlSearch As String
    Private strCondition As String

    Private WithEvents dgv As DataGridView
    Private dt As DataTable
    Private pageIndex As Integer = 0

    '0列都是ID，也要返回值
    Private return_ZeroCol As Integer = -1
    ReadOnly Property ReturnZeroCol() As Integer
        Get
            Return return_ZeroCol
        End Get
    End Property

    Private return_Value As String
    '返回值
    ReadOnly Property ReturnValue() As String
        Get
            ReturnValue = return_Value
        End Get
    End Property

    ''' <summary>
    ''' 需要读取第几列，不设置就是默认0，一般ID就在0
    ''' </summary>
    ''' <remarks></remarks>
    Private Return_Col As Integer
    WriteOnly Property ReturnCol() As Integer
        '返回值
        Set(ByVal value As Integer)
            Return_Col = value
        End Set
    End Property

    Private Table_Name As String
    WriteOnly Property TableName() As String
        '返回值
        Set(ByVal value As String)
            Table_Name = value
        End Set
    End Property


    ''' <summary>
    ''' 设置窗体名称，也是为了把table名称传入
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    WriteOnly Property SetFormText() As String
        Set(ByVal value As String)
            Me.Text = value
        End Set
    End Property

    Private Sub FSelectID_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim iTop As Integer = 5
        Dim iLeft As Integer = 5

        btnSetConditon = New Button
        With btnSetConditon
            .Left = iLeft
            .Top = iTop
            .Text = "设置条件"
            .Width = 50
            .Height = 50
            '.Enabled = False
        End With
        Me.Controls.Add(btnSetConditon)
        iTop += btnSetConditon.Height

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

        iLeft += btnNext.Width
        dgv = New DataGridView
        With dgv
            .Left = iLeft
            .Top = 5
            .ReadOnly = True
        End With
        Me.Controls.Add(dgv)

        Dim tmpCounts As Integer = cdb.GetColZeroValue(String.Format("Select count(*) From [{0}]", Table_Name))
        cpage = New CPages(tmpCounts, PAGE_NUM)
        '表数据
        sqlSearch = String.Format("Select * From [{0}]", Table_Name)
        dt = cdb.ExecuteQuery(String.Format("{0} limit {1} offset {2}", sqlSearch, PAGE_NUM, (PAGE_NUM * pageIndex)))

        dgv.DataSource = dt
        dgv.AutoResizeColumns()
        return_Value = ""

        Me.Width = 600
        Me.Height = 600

        SetFromPos(Me)
    End Sub

    Private Sub FSelectID_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Try
            dgvResize()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub dgvResize()
        With dgv
            .Height = Me.Height - .Top - 60
            .Width = Me.Width - .Left - 10
        End With
    End Sub

    ''' <summary>
    ''' 获取ID
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgv_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv.CellDoubleClick
        Dim c As DataGridViewCell = dgv.SelectedCells(0)
        Dim index As Integer = c.RowIndex
        return_Value = dgv.Rows(index).Cells(Return_Col).Value
        return_ZeroCol = Int(dgv.Rows(index).Cells(0).Value)
        Me.Hide()
    End Sub

    Private Sub btnPre_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPre.Click
        btnPre.Enabled = cpage.Pre
        dt = cdb.ExecuteQuery(sqlSearch & cpage.GetLimitOffset)
        dgv.DataSource = dt
        btnNext.Enabled = True
    End Sub

    Private Sub btnNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNext.Click
        btnNext.Enabled = cpage.NextP
        dt = cdb.ExecuteQuery(sqlSearch & cpage.GetLimitOffset)
        dgv.DataSource = dt
        btnPre.Enabled = True
    End Sub
    Private Sub btnFirst_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFirst.Click
        cpage.First()
        dt = cdb.ExecuteQuery(sqlSearch & cpage.GetLimitOffset())
        dgv.DataSource = dt
    End Sub

    Private Sub btnLast_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLast.Click
        cpage.Last()
        dt = cdb.ExecuteQuery(sqlSearch & cpage.GetLimitOffset())
        dgv.DataSource = dt
    End Sub
    Private Sub btnSetConditon_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSetConditon.Click
        Dim TableIndex As Integer = DB_Info.DicTableIndex(Table_Name)

        Dim F_SearchCondition As New FSearchCondition

        F_SearchCondition = New FSearchCondition
        F_SearchCondition.DBFields = DB_Info.Tables(TableIndex).Field.Name
        F_SearchCondition.DBFieldsType = DB_Info.Tables(TableIndex).Field.Type

        F_SearchCondition.ShowDialog(Me)
        strCondition = F_SearchCondition.ReturnValue
        F_SearchCondition.Hide()

        If strCondition.Length Then
            sqlSearch = String.Format("select * From [{0}] where 1=1 {1}", Table_Name, strCondition)
            Dim tmpCounts As Integer = cdb.GetColZeroValue(String.Format("Select count(*) From ({0})", sqlSearch))
            cpage = New CPages(tmpCounts, PAGE_NUM)
            '表数据
            dt = cdb.ExecuteQuery(String.Format("{0} limit {1} offset {2}", sqlSearch, PAGE_NUM, (PAGE_NUM * pageIndex)))

            dgv.DataSource = dt
            dgv.AutoResizeColumns()
        End If
    End Sub
End Class
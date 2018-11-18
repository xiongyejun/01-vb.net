''' <summary>
''' 针对*ID这样的字段，需要打开对应的table的数据来方便选择id
''' </summary>
''' <remarks></remarks>

Public Class FSelectID

    Private WithEvents dgv As DataGridView

    Private return_Value As Integer
    '返回值，ID
    ReadOnly Property ReturnValue() As Integer
        Get
            ReturnValue = return_Value
        End Get
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
        dgv = New DataGridView
        With dgv
            .Left = 5
            .Top = 5
            .ReadOnly = True
        End With
        Me.Controls.Add(dgv)
        '获取表名称
        Dim TableName As String = Me.Text.Substring(2, Len(Me.Text) - 4)
        '表下标
        Dim TableIndex As Integer = DB_Info.DicTableIndex(TableName)
        '表数据
        If DB_Info.Tables(TableIndex).dt Is Nothing Then
            DB_Info.Tables(TableIndex).dt = MFunc.DoSearch(DB_Info.Path, "Select * From [" & TableName & "]")
        End If

        dgv.DataSource = DB_Info.Tables(TableIndex).dt
        dgv.AutoResizeColumns()
        return_Value = -1

        Me.Width = 600
        Me.Height = 500
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
        return_Value = dgv.Rows(index).Cells("ID").Value
        Me.Hide()
    End Sub
End Class
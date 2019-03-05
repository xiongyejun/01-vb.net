Public Class FExcuteSQL
    Private strSQL As String

    ReadOnly Property GetSQL() As String
        Get
            Return strSQL
        End Get
    End Property

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        strSQL = rtbSQL.Text
        Me.Hide()
    End Sub
    ''' <summary>
    ''' CancelButton
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        strSQL = ""
    End Sub

    ''' <summary>
    ''' 把数据库所有表名添加到lbTables中
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub FExcuteSQL_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.lbTables.Items.Clear()
        For i As Integer = 0 To DB_Info.Tables.Length - 1
            Me.lbTables.Items.Add(DB_Info.Tables(i).Name)
        Next
    End Sub
    ''' <summary>
    ''' lbTables单击时，在lbFields显示当前单击表的字段
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub lbTables_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbTables.Click
        Me.lbFields.Items.Clear()
        Dim index As Integer = lbTables.SelectedIndex
        If index >= 0 Then
            For i As Integer = 0 To DB_Info.Tables(index).Field.Name.Length - 1
                Me.lbFields.Items.Add(DB_Info.Tables(index).Field.Name(i))
            Next
        End If

    End Sub
    ''' <summary>
    ''' lbTables双击时，把双击表的表名替换rtbSQL选中的文本
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub lbTables_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbTables.DoubleClick
        Dim lb As ListBox = CType(sender, ListBox)
        Me.rtbSQL.SelectedText = lb.SelectedItem.ToString
    End Sub
    ''' <summary>
    ''' lbFields双击时，把双击字段的字段名替换rtbSQL选中的文本
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub lbFields_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbFields.DoubleClick
        Dim lb As ListBox = CType(sender, ListBox)
        Me.rtbSQL.SelectedText = lbTables.SelectedItem.ToString & "." & lb.SelectedItem.ToString
    End Sub

  
End Class
''' <summary>
''' 使用窗体输入数据，添加到数据库
''' </summary>
''' <remarks></remarks>

Public Class FAdd
    ''' <summary>
    ''' 确认添加
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents btnOK As Button
    ''' <summary>
    ''' 取消添加
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents btnCancel As Button
    ''' <summary>
    ''' 继续添加，先将记录添加到dt中，再清空tb
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents btnGoOn As Button
    ''' <summary>
    ''' 有多少个字段就添加多少个tb
    ''' </summary>
    ''' <remarks></remarks>
    Private tb() As TextBox

    ''' <summary>
    ''' 数据库当前表的字段
    ''' </summary>
    ''' <remarks></remarks>
    Private DB_Fields() As String
    ''' <summary>
    ''' 数据库当前表字段的类型
    ''' </summary>
    ''' <remarks></remarks>
    Private DB_FieldsType() As System.Type

    ''' <summary>
    ''' 添加到表中使用的ado
    ''' </summary>
    ''' <remarks></remarks>
    Dim c_ado As New CADO(DB_Info.Path)
    ''' <summary>
    ''' 记录添加数据
    ''' </summary>
    ''' <remarks></remarks>
    Dim dt As DataTable
    ''' <summary>
    ''' 设置字段
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    WriteOnly Property DBFields() As String()
        Set(ByVal value() As String)
            DB_Fields = value
        End Set
    End Property
    ''' <summary>
    ''' 设置字段的类型
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    WriteOnly Property DBFieldsType() As System.Type()
        Set(ByVal value() As System.Type)
            DB_FieldsType = value
        End Set
    End Property

    Private Sub FAdd_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Const LB_WIDTH As Integer = 150
        Const TB_WIDTH As Integer = 200
        Dim iTop As Integer = 5
        Dim iLeft As Integer = 5

        ReDim tb(DB_Fields.Length - 1)
        For i As Integer = 0 To DB_Fields.Length - 1
            Dim lb As New Label

            lb.Text = DB_Fields(i) & "(" & DB_FieldsType(i).Name & ")"
            lb.Left = 5
            lb.Top = iTop
            lb.Width = LB_WIDTH

            tb(i) = New TextBox
            tb(i).Width = TB_WIDTH
            tb(i).Left = 5 + lb.Width
            tb(i).Top = iTop
            tb(i).Tag = i
            If DB_Fields(i).Length > 2 Then
                If Strings.Right(DB_Fields(i), 2) = "ID" Then
                    tb(i).ReadOnly = True
                    AddHandler tb(i).Click, AddressOf tb_Click
                End If
            End If


            iTop += lb.Height
            Me.Controls.Add(lb)
            Me.Controls.Add(tb(i))
        Next

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

        btnGoOn = New Button
        With btnGoOn
            .Text = "继续添加"
            .Width = (LB_WIDTH + TB_WIDTH) / 3
            .Left = iLeft
            iLeft += .Width
            .Top = iTop
        End With
        Me.Controls.Add(btnGoOn)

        btnCancel = New Button
        With btnCancel
            .Text = "取消"
            .Width = (LB_WIDTH + TB_WIDTH) / 3
            .Left = iLeft
            iLeft += .Width
            .Top = iTop
        End With
        Me.Controls.Add(btnCancel)

        Me.MaximizeBox = False
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
        Me.Width = LB_WIDTH + TB_WIDTH + 20
        Me.Height = iTop + 70

        '初始化dt的结构
        c_ado.StrSql = "Select * From [" & DB_Info.ActivateTable & "] Where 1=2"
        dt = c_ado.GetData()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        AddRowToDt()
        '更新数据
        c_ado.UpdateData(dt, DB_Info.ActivateTable)
        Me.Close()
    End Sub

    Private Sub btnGoOn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGoOn.Click
        AddRowToDt()
        '清空tb
        For i As Integer = 0 To tb.Length - 1
            tb(i).Clear()
        Next
    End Sub
    ''' <summary>
    ''' 将记录添加到dt中
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function AddRowToDt() As Long
        Dim r As DataRow = dt.NewRow()
        For i As Integer = 0 To tb.Length - 1
            If DB_Fields(i) <> "ID" Then
                r.Item(i) = tb(i).Text
            End If
        Next
        dt.Rows.Add(r)

        Return 1
    End Function


    Private Sub tb_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim f As New FSelectID
        Dim t As TextBox = CType(sender, TextBox)

        f.SetFormText = "选择" & DB_Fields(t.Tag)
        f.ShowDialog(Me)

        Dim tmp As Integer = f.ReturnValue
        If tmp Then t.Text = tmp

        f.Close()
    End Sub
End Class
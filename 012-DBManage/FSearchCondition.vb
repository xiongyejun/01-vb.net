''' <summary>
''' 设置查询的条件
''' </summary>
''' <remarks></remarks>

Public Class FSearchCondition
    Private WithEvents tb As TextBox

    ''' <summary>
    ''' 确认
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents btnOK As Button
    ''' <summary>
    ''' 取消
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents btnCancel As Button

    ''' <summary>
    ''' 1个lable显示字段的名称
    ''' 1个ComboBox选择条件
    ''' 1个textbox输入条件
    ''' </summary>
    ''' <remarks></remarks>
    Structure ControlStruct
        Dim lb As Label
        Dim cb As ComboBox
        Dim tb As TextBox

        Dim Result As String '结果
    End Structure

    ''' <summary>
    ''' 有多少个字段就添加多少组控件
    ''' </summary>
    ''' <remarks></remarks>
    Private ctls() As ControlStruct

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

    Private return_Value As String
    '返回设置的条件
    ReadOnly Property ReturnValue() As String
        Get
            ReturnValue = return_Value
        End Get
    End Property

    ''' <summary>
    ''' cb选择的那些东西
    ''' </summary>
    ''' <remarks></remarks>
    Structure Conditions
        Dim Text As String
        Dim Symbol As String '需要添加的字符，> < Like

        Dim AddBefore As String
        Dim AddAfter As String '主要是Like的%
    End Structure
    Private cdsVal(6) As Conditions
    Private cdsString(6) As Conditions

    ''' <summary>
    ''' 初始化cb条件选择
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function InitConditions() As Boolean
        Dim i As Integer = 0
        cdsVal(i).Text = "" : cdsVal(i).Symbol = "" : i += 1
        cdsVal(i).Text = "大于" : cdsVal(i).Symbol = ">" : i += 1
        cdsVal(i).Text = "小于" : cdsVal(i).Symbol = "<" : i += 1
        cdsVal(i).Text = "等于" : cdsVal(i).Symbol = "=" : i += 1
        cdsVal(i).Text = "不等于" : cdsVal(i).Symbol = "<>" : i += 1
        cdsVal(i).Text = "大于等于" : cdsVal(i).Symbol = ">=" : i += 1
        cdsVal(i).Text = "小于等于" : cdsVal(i).Symbol = "<=" : i += 1

        i = 0
        cdsString(i).Text = "" : cdsString(i).Symbol = "" : cdsString(i).AddBefore = "'" : cdsString(i).AddAfter = "'" : i += 1
        cdsString(i).Text = "等于" : cdsString(i).Symbol = "=" : cdsString(i).AddBefore = "'" : cdsString(i).AddAfter = "'" : i += 1
        cdsString(i).Text = "不等于" : cdsString(i).Symbol = "<>" : cdsString(i).AddBefore = "'" : cdsString(i).AddAfter = "'" : i += 1
        cdsString(i).Text = "开头是" : cdsString(i).Symbol = " Like " : cdsString(i).AddBefore = "'" : cdsString(i).AddAfter = "%'" : i += 1
        cdsString(i).Text = "结尾是" : cdsString(i).Symbol = " Like " : cdsString(i).AddBefore = "'%" : cdsString(i).AddAfter = "'" : i += 1
        cdsString(i).Text = "包含" : cdsString(i).Symbol = " Like " : cdsString(i).AddBefore = "'%" : cdsString(i).AddAfter = "%'" : i += 1
        cdsString(i).Text = "不包含" : cdsString(i).Symbol = " Not Like " : cdsString(i).AddBefore = "'%" : cdsString(i).AddAfter = "%'" : i += 1

        Return True
    End Function

    Private Sub FCondition_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Const LB_WIDTH As Integer = 150
        Const CB_WIDTH As Integer = 80
        Const TB_WIDTH As Integer = 200
        Dim iTop As Integer = 5
        Dim iLeft As Integer = 5

        InitConditions()

        ReDim ctls(DB_Fields.Length - 1)
        For i As Integer = 0 To DB_Fields.Length - 1
            ctls(i).lb = New Label
            With ctls(i).lb
                .Text = DB_Fields(i) & "(" & DB_FieldsType(i).Name & ")"
                .Left = 5
                .Top = iTop
                .Width = LB_WIDTH
            End With

            ctls(i).cb = New ComboBox
            With ctls(i).cb
                .Width = CB_WIDTH
                .Left = ctls(i).lb.Left + ctls(i).lb.Width
                .Top = iTop
                .Tag = i
                .DropDownStyle = ComboBoxStyle.DropDownList
                With .Items
                    If DB_FieldsType(i).Name = "String" Then
                        For j As Integer = 0 To cdsString.Length - 1
                            .Add(cdsString(j).Text)
                        Next
                    Else
                        For j As Integer = 0 To cdsVal.Length - 1
                            .Add(cdsVal(j).Text)
                        Next
                    End If

                End With
                AddHandler .SelectedIndexChanged, AddressOf cb_SelectedIndexChanged
            End With

            ctls(i).tb = New TextBox
            With ctls(i).tb
                .Width = TB_WIDTH
                .Left = 5 + ctls(i).cb.Left + ctls(i).cb.Width
                .Top = iTop
                .Tag = i

                AddHandler .TextChanged, AddressOf tb_TextChanged
            End With

            iTop += ctls(i).lb.Height
            Me.Controls.Add(ctls(i).lb)
            Me.Controls.Add(ctls(i).cb)
            Me.Controls.Add(ctls(i).tb)
        Next

        iTop += 10
        btnOK = New Button
        With btnOK
            .Text = "确定"
            .Width = (LB_WIDTH + TB_WIDTH + CB_WIDTH) / 2
            .Left = iLeft
            iLeft += .Width
            .Top = iTop
        End With
        Me.Controls.Add(btnOK)

        btnCancel = New Button
        With btnCancel
            .Text = "取消"
            .Width = (LB_WIDTH + TB_WIDTH + CB_WIDTH) / 2
            .Left = iLeft
            iLeft += .Width
            .Top = iTop
        End With
        Me.Controls.Add(btnCancel)

        Me.MaximizeBox = False
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
        Me.Width = LB_WIDTH + TB_WIDTH + CB_WIDTH + 20
        Me.Height = iTop + 70

    End Sub


    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        For i As Integer = 0 To ctls.Length - 1
            If ctls(i).Result <> "" Then
                return_Value &= (" And " & ctls(i).Result)
            End If
        Next
        Me.Hide()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Hide()
    End Sub

    Private Sub cb_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim index As Integer
        index = Val(sender.tag)
        GetResult(index, ctls(index).cb.SelectedIndex)
    End Sub

    Private Sub tb_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb.TextChanged
        Dim index As Integer
        index = Val(sender.tag)
        GetResult(index, ctls(index).cb.SelectedIndex)
    End Sub

    Private Function GetResult(ByVal index As Integer, ByVal selectIndex As Integer) As Boolean
        'cb选择了，并且tb填写了数据
        If selectIndex > 0 AndAlso ctls(index).tb.Text <> "" Then
            If DB_FieldsType(index).Name = "DateTime" Then
                ctls(index).Result = DB_Fields(index) & cdsVal(selectIndex).Symbol & "#" & ctls(index).tb.Text & "#"
            ElseIf DB_FieldsType(index).Name = "String" Then
                ctls(index).Result = DB_Fields(index) & cdsString(selectIndex).Symbol & cdsString(selectIndex).AddBefore & ctls(index).tb.Text & cdsString(selectIndex).AddAfter
            Else
                ctls(index).Result = DB_Fields(index) & cdsVal(selectIndex).Symbol & ctls(index).tb.Text
            End If
        Else
            ctls(index).Result = ""
        End If

        Return True
    End Function
End Class
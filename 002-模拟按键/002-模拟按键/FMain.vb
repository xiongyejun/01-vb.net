Imports System.IO

Public Class FMain
    Dim iSleep As Integer = 0
    Const USER_NAME As String = "029013、033065、"  'system.Environment.UserName
    Const FU_HAO As String = "※"
    Const FILE_NAME As String = "ZhiLing.txt"

    '鼠标键盘钩子
    Friend WithEvents MyHk As New CHook(False)


    Friend lbZhiLing As Label            '指令标签
    Friend WithEvents btnClearRTB As Button         '清除指令按钮
    Friend WithEvents rtbZhiLing As RichTextBox     '指令文本框
    Friend WithEvents lvKeys As ListView            '功能键列表框

    Friend WithEvents tbWenBen As TextBox           '添加指令的文本框
    Friend WithEvents btnAddWenBen As Button        '添加指令文本的按钮
    Friend WithEvents btnZuoBiao As Button          '控制鼠标钩子的按钮

    Friend lbSleep As Label              '延迟时间标签
    Friend nudSleep As NumericUpDown     '延迟时间设置
    Friend WithEvents btnStart As Button            '开始按钮

    Friend lbLoop As Label              '循环次数
    Friend nudLoop As NumericUpDown     '循环次数

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MyHk.KeyHookEnabled = False
        MyHk.MouseHookEnabled = False

        Dim filename As String = Application.StartupPath & "\" & FILE_NAME
        If Me.rtbZhiLing.Text <> "" Then
            WriteText(filename, Me.rtbZhiLing.Text)
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim iTop As Integer = 10
        Dim iLeft As Integer = 10
        Const iStep As Integer = 30

        With Me
            .Size = New Size(400, 400)
            .FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
            .MaximizeBox = False
            .Text = "模拟按键"
            '.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        End With

        lbZhiLing = New Label
        With lbZhiLing
            .Location = New Point(iLeft, iTop)
            .Text = "按键指令:"
        End With
        btnClearRTB = New Button
        With btnClearRTB
            .Location = New Point(iLeft + lbZhiLing.Width + lbZhiLing.Left, iTop)
            .Text = "清除指令"
        End With

        '指令文本框
        iTop += lbZhiLing.Height
        rtbZhiLing = New RichTextBox
        With rtbZhiLing
            .Location = New Point(iLeft, iTop)
            .Width = Me.Width - 20
            '.ReadOnly = True
        End With

        '功能按键选择列表
        iTop += rtbZhiLing.Height
        lvKeys = New ListView
        With lvKeys
            .Location = New Point(iLeft, iTop)
            .Width = 230
            .Height = Me.Height - rtbZhiLing.Height - lbZhiLing.Height - 50

            .Columns.Add("序号", 40, HorizontalAlignment.Left)
            .Columns.Add("键", 80, HorizontalAlignment.Left)
            .Columns.Add("代码", 80)
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True

            .Items.AddRange(GetKeys)
        End With

        '文本编辑
        iLeft += lvKeys.Width + 5

        tbWenBen = New TextBox
        With tbWenBen
            .Location = New Point(iLeft, iTop)
            .Width = Me.Height - lvKeys.Height - 30
        End With

        iTop += iStep
        btnAddWenBen = New Button
        With btnAddWenBen
            .Location = New Point(iLeft, iTop)
            .Text = "添加文本"
        End With

        '添加坐标
        iTop += iStep
        btnZuoBiao = New Button
        With btnZuoBiao
            .Location = New Point(iLeft, iTop)
            .Text = "启动MouseHook"
            .Width = 100
        End With

        '设置延迟时间
        iTop += iStep
        lbSleep = New Label
        With lbSleep
            .Location = New Point(iLeft, iTop + 5)
            .Text = "设置延迟时间(毫秒)"
            .Width = 120
        End With

        iTop += iStep
        nudSleep = New NumericUpDown
        With nudSleep
            .Location = New Point(iLeft, iTop)
            .Width = 50
            .Minimum = 200
            .Maximum = 10000
            .Increment = 200
            .Value = 1000
        End With

        iTop += iStep
        lbLoop = New Label
        With lbLoop
            .Location = New Point(iLeft, iTop + 5)
            .Text = "设置循环次数"
            .Width = 120
        End With

        iTop += iStep
        nudLoop = New NumericUpDown
        With nudLoop
            .Location = New Point(iLeft, iTop)
            .Width = 50
            .Minimum = 1
            .Maximum = 1000
            .Increment = 1
            .Value = 1
        End With

        '开始按钮
        iTop += iStep
        btnStart = New Button
        With btnStart
            .Location = New Point(iLeft, iTop)
            .Text = "Start"
        End With


        With Me.Controls
            .Add(lbZhiLing)
            .Add(btnClearRTB)
            .Add(rtbZhiLing)
            .Add(lvKeys)

            .Add(tbWenBen)
            .Add(btnAddWenBen)

            .Add(btnZuoBiao)

            .Add(btnStart)
            .Add(nudSleep)
            .Add(lbSleep)

            .Add(nudLoop)
            .Add(lbLoop)
        End With

        Dim filename As String = Application.StartupPath & "\" & FILE_NAME
        If File.Exists(filename) Then
            Me.rtbZhiLing.Text = ReadText(filename)
        End If
    End Sub

    Private Sub lvKeys_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvKeys.DoubleClick
        Dim str As String = Me.lvKeys.SelectedItems(0).SubItems(2).Text
        rtbTextAdd(str)
    End Sub

    Private Sub btnAddWenBen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddWenBen.Click
        Dim str As String = Me.tbWenBen.Text
        rtbTextAdd(str)
    End Sub

    Private Sub btnZuoBiao_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnZuoBiao.Click
        With Me.btnZuoBiao
            If .Text = "启动MouseHook" Then
                .Text = "关闭MouseHook"
                MyHk.KeyHookEnabled = True
                MyHk.MouseHookEnabled = True
            Else
                .Text = "启动MouseHook"
                MyHk.KeyHookEnabled = False
                MyHk.MouseHookEnabled = False
            End If
        End With
    End Sub

    Private Sub btnClearRTB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClearRTB.Click
        Me.rtbZhiLing.Text = ""
    End Sub

    Private Sub btnStart_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Dim Arr() As String = Split(Me.rtbZhiLing.Text, FU_HAO)
        iSleep = Me.nudSleep.Value
        Me.WindowState = FormWindowState.Minimized
        System.Threading.Thread.Sleep(iSleep)

        For j As Integer = 0 To Val(nudLoop.Value) - 1
            For i As Integer = 0 To Arr.Length - 1
                If Strings.Left(Arr(i), 2) = "X(" Then
                    Dim strTmp As String = Strings.Mid(Arr(i), 3, Len(Arr(i)) - 2)
                    Dim temp() As String = Split(strTmp, ",")
                    Screen_Click(Val(temp(1)), Val(temp(2)))
                    System.Threading.Thread.Sleep(Val(temp(0)))
                Else
                    ShuRu(Arr(i))
                    System.Threading.Thread.Sleep(iSleep)
                End If

            Next
        Next

        Erase Arr
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub rtbTextAdd(ByVal str As String)
        Dim Str1 As String = ""
        Dim i As Integer = Me.rtbZhiLing.SelectionStart

        With Me.rtbZhiLing
            If .Text = "" Then
                .Text = str
            Else
                If i = 0 Then
                    str = str & FU_HAO
                Else
                    str = FU_HAO & str
                End If

                .Text = .Text.Insert(i, str)   '添加符号
            End If

            i = i + Len(str)
            .SelectionStart = i

        End With
    End Sub

    Private Sub MyHk_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyHk.KeyDown
        Dim iCode As Integer = e.KeyCode
        If iCode = 112 Then
            Me.btnZuoBiao.Text = "启动MouseHook"
            MyHk.KeyHookEnabled = False
            MyHk.MouseHookEnabled = False
        End If
    End Sub



    Private Sub MyHk_MouseActivity(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyHk.MouseActivity

        Dim x As Integer = System.Windows.Forms.Control.MousePosition.X
        Dim y As Integer = System.Windows.Forms.Control.MousePosition.Y
        Me.tbWenBen.Text = "X(3000," & x & "," & y & ")Y"

    End Sub
End Class

Imports System.IO
Imports System.Threading.Thread

Public Class Form1
    Private WithEvents lbNowPoint As System.Windows.Forms.Label
    Private WithEvents tbNowPoint As System.Windows.Forms.TextBox

    Private WithEvents lbOffset As System.Windows.Forms.Label
    Private WithEvents nudOffset As System.Windows.Forms.NumericUpDown

    Private WithEvents lbSourePoint As System.Windows.Forms.Label
    Private WithEvents tbSourePoint As System.Windows.Forms.TextBox

    Private WithEvents lbDesPoint As System.Windows.Forms.Label
    Private WithEvents tbDesPoint As System.Windows.Forms.TextBox

    Private WithEvents lbState As System.Windows.Forms.Label

    Private WithEvents btnKey As System.Windows.Forms.Button
    Private lbSleep As System.Windows.Forms.Label
    Private nud As System.Windows.Forms.NumericUpDown

    Private WithEvents btnSun As System.Windows.Forms.Button
    Private WithEvents tGame As New System.Windows.Forms.Timer

    Private WithEvents kc As New hookClass(False, False)


    Const KEY_str As String = "F1：MouseHook开关，F2：KeyHook开关,F3:添加坐标,F4:停止（先开KeyHook）"
    Dim txt As String = Application.StartupPath & "\shezhi.txt"
    Dim t As Threading.Thread
    Dim p As New PlantsVsZombies

    Private Declare Function GetPixel Lib "gdi32" Alias "GetPixel" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Integer
    '获取指定窗口的设备场景
    Private Declare Function GetDC Lib "user32" Alias "GetDC" (ByVal hwnd As Integer) As Integer
    '释放由调用GetDC或GetWindowDC函数获取的指定设备场景。它对类或私有设备场景无效（但这样的调用不会造成损害）
    Private Declare Function ReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hwnd As Integer, ByVal hdc As Integer) As Integer

    Function getPix(ByVal x As Integer, ByVal y As Integer, ByVal pixInt As Integer) As Boolean
        If pixInt = GetPixel(GetDC(0), x, y) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim iLeft As Integer = 5
        Dim iTop As Integer = 5

        lbNowPoint = New Label
        Func.labelAdd(Me, lbNowPoint, "当前坐标:", iLeft, iTop)
        iLeft += lbNowPoint.Width
        tbNowPoint = New TextBox
        Func.textBoxAdd(Me, tbNowPoint, "", iLeft, iTop)

        iLeft += tbNowPoint.Width
        iLeft += 5

        lbOffset = New Label
        Func.labelAdd(Me, lbOffset, "偏移坐标:", iLeft, iTop)
        iLeft += lbOffset.Width
        nudOffset = New NumericUpDown
        Func.numericUpDownAdd(Me, nudOffset, iLeft, iTop)
        nudOffset.Minimum = 0
        nudOffset.Maximum = 10000
        nudOffset.Increment = 10
        nudOffset.Value = 50

        iTop += tbNowPoint.Height
        iTop += 5
        iLeft = 5

        lbSourePoint = New Label
        Func.labelAdd(Me, lbSourePoint, "  源坐标:", iLeft, iTop)
        iLeft += lbSourePoint.Width
        tbSourePoint = New TextBox
        Func.textBoxAdd(Me, tbSourePoint, "", iLeft, iTop, 100, 100)
        tbSourePoint.Multiline = True
        iLeft += tbSourePoint.Width

        iLeft += 5

        lbDesPoint = New Label
        Func.labelAdd(Me, lbDesPoint, "目标坐标:", iLeft, iTop)
        iLeft += lbDesPoint.Width
        tbDesPoint = New TextBox
        Func.textBoxAdd(Me, tbDesPoint, "", iLeft, iTop, tbSourePoint.Width, tbSourePoint.Height)
        tbDesPoint.Multiline = True

        iTop += tbDesPoint.Height
        iTop += 5
        iLeft = 5
        btnKey = New Button
        Func.btnAdd(Me, btnKey, "开始按键", iLeft, iTop)

        iLeft += btnKey.Width
        lbSleep = New Label
        Func.labelAdd(Me, lbSleep, "iSleep", iLeft, iTop)

        iLeft += lbSleep.Width
        nud = New NumericUpDown
        Func.numericUpDownAdd(Me, nud, iLeft, iTop)
        nud.Minimum = 10
        nud.Maximum = 10000
        nud.Increment = 10
        nud.Value = 500

        iTop += btnKey.Height
        iTop += 5
        iLeft = 5
        btnSun = New Button
        Func.btnAdd(Me, btnSun, "Sun", iLeft, iTop)

        iTop += btnSun.Height
        iTop += 5
        iLeft = 5
        lbState = New Label
        Func.labelAdd(Me, lbState, KEY_str, iLeft, iTop, 500, 100)
        'lbState.AutoSize = True



        Me.Width = Me.lbState.Width + 20
        Me.Left = Screen.PrimaryScreen.Bounds.Width - Me.Width
        Me.Top = Screen.PrimaryScreen.Bounds.Height - Me.Height - 30

        readTxt()
    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown,
                                                            lbNowPoint.KeyDown, lbDesPoint.KeyDown, lbSourePoint.KeyDown, lbState.KeyDown,
                                                            tbNowPoint.KeyDown, tbDesPoint.KeyDown, tbSourePoint.KeyDown
        Select Case e.KeyCode
            Case Keys.F1
                kc.MouseHookEnabled = Not kc.MouseHookEnabled
            Case Keys.F2
                kc.KeyHookEnabled = Not kc.KeyHookEnabled
            Case Keys.F3
                Me.tbDesPoint.Text = Me.tbDesPoint.Text & Me.tbNowPoint.Text
        End Select

        Me.lbState.Text = getState()
    End Sub

    Function getState() As String
        Dim str As String = KEY_str

        str &= Chr(10) & "状态:" & Chr(10)
        str &= "MouseHookEnabled = " & kc.MouseHookEnabled & Chr(10)
        str &= "KeyHookEnabled =" & kc.KeyHookEnabled & Chr(10)

        Return str
    End Function

    Private Sub kc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles kc.KeyDown
        Select Case e.KeyCode
            Case Keys.F4
                btnKeyClick()
        End Select
    End Sub

    Private Sub kc_MouseActivity(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles kc.MouseActivity
        Me.tbNowPoint.Text = e.Location.ToString
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        kc.KeyHookEnabled = False
        kc.MouseHookEnabled = False

        Dim str As String = Me.tbSourePoint.Text
        str &= vbNewLine
        str &= Me.tbDesPoint.Text

        str &= vbNewLine
        str &= Me.nud.Value

        str &= vbNewLine
        str &= Me.nudOffset.Value
        Func.WriteText(txt, str)

        Try
            t.Abort()
        Catch ex As Exception

        End Try
    End Sub

    Function readTxt()
        If File.Exists(txt) Then
            Dim str As String = Func.ReadText(txt)
            Dim arr() As String = Split(str, vbNewLine)

            Me.tbSourePoint.Text = arr(0)
            Me.tbDesPoint.Text = arr(1)
            Me.nud.Value = Val(arr(2))
            Me.nudOffset.Value = Val(arr(3))
        End If

        Return 0
    End Function

    Sub start()
        Dim x1 As Integer = 0, y1 As Integer = 0
        Dim x2 As Integer = 0, y2 As Integer = 0
        Dim iSleep As Integer = Me.nud.Value
        Dim iOffset As Integer = Me.nudOffset.Value

        Dim arrSoure() As String = Split(Me.tbSourePoint.Text, "{")
        Dim arrDes() As String = Split(Me.tbDesPoint.Text, "{")

        Do While True

            'X=355 Y=560 Pixel=8107759
            If Me.getPix(355, 560, 8107759) Then    '判断是否到了图鉴页面
                Const clickSleep As Integer = 2000
                Func.Screen_Click(716, 612)       'X=716 Y=612关闭
                Threading.Thread.Sleep(clickSleep)

                Threading.Thread.Sleep(clickSleep)
                Func.Screen_Click(264, 329)       'X=264 Y=329辣椒
                Threading.Thread.Sleep(clickSleep)
                Func.Screen_Click(264, 329)       'X=264 Y=329辣椒
                Threading.Thread.Sleep(clickSleep)


                Threading.Thread.Sleep(clickSleep)
                Func.Screen_Click(52, 329)       'X=52 Y=327浮萍
                Threading.Thread.Sleep(clickSleep)

                Threading.Thread.Sleep(clickSleep)
                Func.Screen_Click(162, 398)       'X=162 Y=397向日葵--打空中的，为了防止出提示
                Threading.Thread.Sleep(clickSleep)

                For i As Integer = 50 To 350 Step 50
                    Func.Screen_Click(i, 190)
                    Threading.Thread.Sleep(clickSleep)
                Next
                Func.Screen_Click(240, 597)          'X=240 Y=597开始战斗
                Threading.Thread.Sleep(clickSleep)
                Func.Screen_Click(240, 597)          'X=240 Y=597开始战斗
                Threading.Thread.Sleep(clickSleep)
                Func.Screen_Click(240, 597)          'X=240 Y=597开始战斗
                Threading.Thread.Sleep(clickSleep)

            Else
                For j As Integer = 1 To arrDes.Length - 1
                    getPoint(x2, y2, arrDes(j))
                    For i As Integer = 1 To arrSoure.Length - 1
                        getPoint(x1, y1, arrSoure(i))

                        Func.Screen_Click(x1, y1)
                        Threading.Thread.Sleep(iSleep)

                        x2 += iOffset * (i - 1)
                        Func.Screen_Click(x2, y2)
                        Threading.Thread.Sleep(iSleep)
                    Next
                Next
            End If

        Loop
    End Sub

    Function getPoint(ByRef x As Integer, ByRef y As Integer, ByVal str As String) 'X=235,Y=260}
        Dim arr() As String = Split(str, ",")

        x = CInt(Replace(arr(0), "X=", ""))

        y = CInt(Replace(Replace(arr(1), "Y=", ""), "}", ""))
        Return 0
    End Function

    Private Sub btnKey_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnKey.Click
        btnKeyClick()
    End Sub

    Sub btnKeyClick()
        If Me.btnKey.Text = "开始按键" Then
            t = New Threading.Thread(AddressOf start)
            t.Start()
            Me.btnKey.Text = "结束按键"
        Else
            Try
                t.Abort()
            Catch ex As Exception

            End Try
            Me.btnKey.Text = "开始按键"
        End If

    End Sub

    Private Sub btnSun_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSun.Click
        Me.tGame.Interval = 100
        Me.tGame.Enabled = Not Me.tGame.Enabled
    End Sub

    Private Sub tGame_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tGame.Tick
        p.SunshineValue = 9999
        p.AllPlantsCoolDownDisEnable()
    End Sub
End Class

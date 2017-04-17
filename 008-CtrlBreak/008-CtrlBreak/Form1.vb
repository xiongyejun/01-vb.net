Public Class Form1

    Private cb_window As System.Windows.Forms.ComboBox
    Private WithEvents btn_click As System.Windows.Forms.Button
    Private WithEvents btn_getwindow As System.Windows.Forms.Button
    Private num As System.Windows.Forms.NumericUpDown

    '在窗口列表中寻找与指定条件相符的第一个子窗口
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer

    '寻找窗口列表中第一个符合指定条件的顶级窗口（在vb里使用：FindWindow最常见的一个用途是获得ThunderRTMain类的隐藏窗口的句柄；该类是所有运行中vb执行程序的一部分。获得句柄后，可用api函数GetWindowText取得这个窗口的名称；该名也是应用程序的标题）
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer

    '枚举窗口列表中的所有父窗口（顶级和被所有窗口）
    Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Integer) As Integer
    '取得一个窗体的标题（caption）文字，或者一个控件的内容（在vb里使用：使用vb窗体或控件的caption或text属性）
    Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Integer, ByVal lpString As System.Text.StringBuilder, ByVal cch As Integer) As Integer
    '调查窗口标题文字或控件内容的长短（在vb里使用：直接使用vb窗体或控件的caption或text属性）
    Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Integer) As Integer

    Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Integer) As Integer
    Declare Sub keybd_event Lib "user32" (
                                ByVal bVk As Byte,
                                ByVal bScan As Byte,
                                ByVal dwFlags As Integer,
                                ByVal dwExtraInfo As Integer)

    Const KEYEVENT_KEYUP = &H2
    Const VK_CANCEL = &H3 'Control-break processing


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        cb_window = New ComboBox

        With cb_window
            .Left = 5
            .Top = 5
            .Height = 50
            .Width = 500
        End With

        btn_click = New Button
        With btn_click
            .Left = 5
            .Top = 10 + Me.cb_window.Height
            .Text = "点击"
            .Height = 30
        End With


        btn_getwindow = New Button
        With btn_getwindow
            .Left = 5 + btn_click.Width
            .Top = 10 + Me.cb_window.Height
            .Text = "获取窗口"
            .Height = 30
        End With

        num = New NumericUpDown
        With num
            .Value = 3
            .Left = 5 + btn_getwindow.Width + btn_getwindow.Left
            .Top = btn_getwindow.Top + 3
        End With

        Me.Controls.Add(cb_window)
        Me.Controls.Add(btn_click)
        Me.Controls.Add(btn_getwindow)
        Me.Controls.Add(num)

        Me.Width = Me.cb_window.Width + 50
    End Sub

    Private Sub btn_click_Click(sender As Object, e As EventArgs) Handles btn_click.Click
        Dim str_window As String = Me.cb_window.Text
        AppActivate(str_window)
        ctrl_break(Me.num.Value)
    End Sub

    Function ctrl_break(Optional k As Integer = 1)
        For i As Integer = 1 To k
            keybd_event(VK_CANCEL, 0, 0, 0)       '按下
            keybd_event(VK_CANCEL, 0, KEYEVENT_KEYUP, 0)    '松开
        Next i
        Return 0
    End Function

    Private Sub btn_getwindow_Click(sender As Object, e As EventArgs) Handles btn_getwindow.Click
        Me.cb_window.Items.Clear()
        'EnumWindow(FindWindowEx(0, 0, 0, 0))
        EnumWindow(FindWindow("Shell_TrayWnd", vbNullString))
    End Sub

    Function EnumWindow(ByVal hwnd As Integer) As Integer
        'Dim hwnd1 As Integer = FindWindowEx(0, 0, vbNullString, vbNullString)

        Dim k As Integer = 1
        Dim myHwnd As Integer = 0 ' hwnd1

        Do
            myHwnd = FindWindowEx(0, myHwnd, vbNullString, vbNullString)
            'Dim iText As Integer = GetWindowTextLength(myHwnd)
            If IsWindowVisible(myHwnd) Then
                Dim Str As New System.Text.StringBuilder(256)
                GetWindowText(myHwnd, Str, Str.Capacity)

                If Str.ToString <> "" Then
                    Me.cb_window.Items.Add(Str.ToString)
                    k = k + 1
                End If
            End If
        Loop Until (myHwnd = 0)

        Return 0
    End Function
End Class

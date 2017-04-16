Public Class Form1

    Private cb_window As System.Windows.Forms.ComboBox
    Private WithEvents btn_click As System.Windows.Forms.Button
    Private WithEvents btn_getwindow As System.Windows.Forms.Button

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
            .Height = 30
            .Width = 250
        End With

        btn_click = New Button
        With btn_click
            .Left = 5
            .Top = 10 + Me.cb_window.Height
            .Text = "点击"
        End With


        btn_getwindow = New Button
        With btn_getwindow
            .Left = 5 + btn_click.Width
            .Top = 10 + Me.cb_window.Height
            .Text = "获取窗口"
        End With

        Me.Controls.Add(cb_window)
        Me.Controls.Add(btn_click)
        Me.Controls.Add(btn_getwindow)

        GetWindow()
    End Sub

    Function GetWindow()
        cb_window.Items.Clear()

        For Each p As System.Diagnostics.Process In System.Diagnostics.Process.GetProcesses
            If p.MainWindowTitle <> "" Then
                Me.cb_window.Items.Add(p.MainWindowTitle)
            End If

        Next

        Return 0
    End Function

    Private Sub btn_click_Click(sender As Object, e As EventArgs) Handles btn_click.Click
        Dim str_window As String = Me.cb_window.Text
        AppActivate(str_window)
        ctrl_break()
    End Sub

    Function ctrl_break()

        keybd_event(VK_CANCEL, 0, 0, 0)       '按下
        keybd_event(VK_CANCEL, 0, KEYEVENT_KEYUP, 0)    '松开
        Return 0
    End Function

    Private Sub btn_getwindow_Click(sender As Object, e As EventArgs) Handles btn_getwindow.Click
        GetWindow()
    End Sub
End Class

Imports System.Net
Imports System.IO

Module ControlAdd
    Public Const MOUSEEVENTF_MOVE = &H1 '移动鼠标
    Public Const MOUSEEVENTF_ABSOLUTE = &H8000 '指定鼠标使用绝对坐标系，此时，屏幕在水平和垂直方向上均匀分割成65535×65535个单元
    Public Const MOUSEEVENTF_LEFTDOWN = &H2 '模拟鼠标左键按下
    Public Const MOUSEEVENTF_LEFTUP = &H4 '模拟鼠标左键抬起
    Public Const KEYEVENTF_KEYUP = &H2
    Public Const MOUSEEVENTF_WHEEL As Integer = &H800

    '这个函数模拟了键盘行动
    Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Integer, ByVal dwExtraInfo As Integer)
    '模拟一次鼠标事件
    Public Declare Sub mouse_event Lib "user32" Alias "mouse_event" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)

    '在窗口列表中寻找与指定条件相符的第一个子窗口
    Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwnd1 As Integer, ByVal hwnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    '控制窗口的可见性（在vb里使用：针对vb窗体及控件，请使用对应的vb属性）
    Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer
    Public Const SW_SHOWMAXIMIZED As Integer = 3 '最大化窗口
    '将窗口设为系统的前台窗口。这个函数可用于改变用户目前正在操作的应用程序
    Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Integer) As Integer
    '获得活动窗口的句柄
    Public Declare Function GetActiveWindow Lib "user32" Alias "GetActiveWindow" () As Integer
    '取得一个窗体的标题（caption）文字，或者一个控件的内容（在vb里使用：使用vb窗体或控件的caption或text属性）
    Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Integer, ByVal lpString As String, ByVal cch As Integer) As Integer

    Public Sub MoseMove()
        Dim x As Integer, y As Integer
        x = System.Windows.Forms.Control.MousePosition.X
        y = System.Windows.Forms.Control.MousePosition.Y
        Dim w As Integer = Screen.PrimaryScreen.Bounds.Width
        Dim h As Integer = Screen.PrimaryScreen.Bounds.Height

        mouse_event(MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, (x + 100) / w * 65535, (y + 100) / h * 65535, 0, 0)
        mouse_event(MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, x / w * 65535, y / h * 65535, 0, 0)
    End Sub

    Sub Screen_Click(ByVal X As Integer, ByVal Y As Integer, Optional ByVal n As Integer = 1)  '按鼠标
        Dim mw As Integer, mh As Integer
        mw = X / Screen.PrimaryScreen.Bounds.Width * 65535
        mh = Y / Screen.PrimaryScreen.Bounds.Height * 65535

        '    If IsMissing(n) Then n = 1
        Dim i As Integer
        For i = 1 To n Step 1
            mouse_event(MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, mw, mh, 0, 0)
            mouse_event(MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, mw, mh, 0, 0)
        Next i
    End Sub

    Sub btnAdd(ByVal form As Object, ByVal btn As System.Windows.Forms.Button, ByVal btnText As String, Optional ByVal iLeft As Integer = 5, Optional ByVal iTop As Integer = 0, Optional ByVal iWidth As Integer = 100, Optional ByVal iHeight As Integer = 30)
        With btn
            .Left = iLeft
            .Top = iTop
            .Width = iWidth
            .Height = iHeight
            .Text = btnText
        End With
        form.Controls.Add(btn)
    End Sub

    Sub checkBoxAdd(ByVal form As Object, ByVal cb As System.Windows.Forms.CheckBox, ByVal text As String, Optional ByVal iLeft As Integer = 5, Optional ByVal iTop As Integer = 0, Optional ByVal iWidth As Integer = 200, Optional ByVal iHeight As Integer = 30)
        With cb
            .Left = iLeft
            .Top = iTop
            .Width = iWidth
            .Height = iHeight
            .Text = text
        End With
        form.Controls.Add(cb)
    End Sub

    Sub comboBoxAdd(ByVal form As Object, ByVal cb As System.Windows.Forms.ComboBox, ByVal text As String, Optional ByVal iLeft As Integer = 5, Optional ByVal iTop As Integer = 0, Optional ByVal iWidth As Integer = 100, Optional ByVal iHeight As Integer = 30)
        With cb
            .Left = iLeft
            .Top = iTop
            .Width = iWidth
            .Height = iHeight
            .Text = text
        End With
        form.Controls.Add(cb)
    End Sub
    Sub groupBoxAdd(ByVal form As Object, ByVal gp As System.Windows.Forms.GroupBox, ByVal text As String, Optional ByVal iLeft As Integer = 5, Optional ByVal iTop As Integer = 0, Optional ByVal iWidth As Integer = 100, Optional ByVal iHeight As Integer = 30)
        With gp
            .Left = iLeft
            .Top = iTop + 5
            .Width = iWidth
            .Height = iHeight
            .Text = text
        End With
        form.Controls.Add(gp)
    End Sub

    Sub labelBoxAdd(ByVal form As Object, ByVal label As System.Windows.Forms.Label, ByVal labelText As String, Optional ByVal iLeft As Integer = 5, Optional ByVal iTop As Integer = 0, Optional ByVal iWidth As Integer = 100, Optional ByVal iHeight As Integer = 30)
        With label
            .Left = iLeft
            .Top = iTop + 5
            .Width = iWidth
            .Height = iHeight
            .Text = labelText
            '.AutoSize = True
        End With
        form.Controls.Add(label)
    End Sub

    Sub listViewAdd(ByVal form As Object, ByVal lv As System.Windows.Forms.ListView, Optional ByVal iLeft As Integer = 5, Optional ByVal iTop As Integer = 0, Optional ByVal iWidth As Integer = 100, Optional ByVal iHeight As Integer = 30)
        With lv
            .Left = iLeft
            .Top = iTop
            .Width = iWidth
            .Height = iHeight
        End With
        form.Controls.Add(lv)
    End Sub

    Sub numericUpDownAdd(ByVal form As Object, ByVal num As System.Windows.Forms.NumericUpDown, Optional ByVal iLeft As Integer = 5, Optional ByVal iTop As Integer = 0, Optional ByVal iWidth As Integer = 60, Optional ByVal iHeight As Integer = 30)
        With num
            .Left = iLeft
            .Top = iTop + 5
            .Width = iWidth
            .Height = iHeight
        End With
        form.Controls.Add(num)
    End Sub

    Sub radioButtonAdd(ByVal form As Object, rb As System.Windows.Forms.RadioButton, ByVal textBoxText As String, Optional ByVal iLeft As Integer = 5, Optional ByVal iTop As Integer = 0, Optional ByVal iWidth As Integer = 100, Optional ByVal iHeight As Integer = 30, Optional val As Boolean = False)
        With rb
            .Left = iLeft
            .Top = iTop
            .Width = iWidth
            .Height = iHeight
            .Text = textBoxText
            .Checked = val
            .AutoSize = True
        End With
        form.Controls.Add(rb)
    End Sub

    Sub textBoxAdd(ByVal form As Object, ByVal textBox As System.Windows.Forms.TextBox, ByVal textBoxText As String, Optional ByVal iLeft As Integer = 5, Optional ByVal iTop As Integer = 0, Optional ByVal iWidth As Integer = 100, Optional ByVal iHeight As Integer = 30)
        With textBox
            .Left = iLeft
            .Top = iTop
            .Width = iWidth
            .Height = iHeight
            .Text = textBoxText
            .AutoSize = True
        End With
        form.Controls.Add(textBox)
    End Sub

    Sub richTextBoxAdd(ByVal form As Object, ByVal rtb As System.Windows.Forms.RichTextBox, ByVal text As String, Optional ByVal iLeft As Integer = 5, Optional ByVal iTop As Integer = 0, Optional ByVal iWidth As Integer = 100, Optional ByVal iHeight As Integer = 30)
        With rtb
            .Left = iLeft
            .Top = iTop
            .Width = iWidth
            .Height = iHeight
            .Text = text
        End With
        form.Controls.Add(rtb)
    End Sub

    Function readHtml(ByVal Url As String) As String
        Dim httpReq As System.Net.HttpWebRequest
        Dim httpResp As System.Net.HttpWebResponse
        Dim httpURL As New System.Uri(Url)
        httpReq = CType(WebRequest.Create(httpURL), HttpWebRequest)
        httpReq.Method = "GET"
        httpResp = CType(httpReq.GetResponse(), HttpWebResponse)
        httpReq.KeepAlive = False ' 获取或设置一个值，该值指示是否与 Internet资源建立持久连接。
        Dim reader As StreamReader = New StreamReader(httpResp.GetResponseStream, System.Text.Encoding.GetEncoding(-0))
        Dim respHTML As String = reader.ReadToEnd()

        Return respHTML
    End Function

    Function downFile(ByVal file As String, ByVal saveFile As String)
        Dim DownApp As New Microsoft.VisualBasic.Devices.Network ' System.Net.WebClient 'Microsoft.VisualBasic.Devices.Network
        DownApp.DownloadFile(file, saveFile, "", "", False, 5000, True)
        Return 0
    End Function

    Function ReadText(ByVal file As String) As String
        Dim sr As StreamReader = New StreamReader(file)
        Dim str As String = sr.ReadToEnd
        sr.Close()
        Return str
    End Function

    Sub WriteText(ByVal file As String, ByVal str As String)
        Dim sw As StreamWriter = New StreamWriter(file)
        sw.Write(str)
        sw.Close()
    End Sub

End Module

Imports system.IO

Module Module1
    Public Const MOUSEEVENTF_MOVE = &H1 '移动鼠标
    Public Const MOUSEEVENTF_ABSOLUTE = &H8000 '指定鼠标使用绝对坐标系，此时，屏幕在水平和垂直方向上均匀分割成65535×65535个单元
    Public Const MOUSEEVENTF_LEFTDOWN = &H2 '模拟鼠标左键按下
    Public Const MOUSEEVENTF_LEFTUP = &H4 '模拟鼠标左键抬起
    Public Const KEYEVENTF_KEYUP = &H2
    Public Const MOUSEEVENTF_WHEEL As Integer = &H800
    Public Declare Sub mouse_event Lib "user32" Alias "mouse_event" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)

    Sub Screen_Click(ByVal X As Integer, ByVal Y As Integer, Optional ByVal n As Integer = 1)  '按鼠标
        Dim mw As Integer, mh As Integer
        mw = X / Screen.PrimaryScreen.Bounds.Width * 65535
        mh = Y / Screen.PrimaryScreen.Bounds.Height * 65535

        Dim i As Integer
        For i = 1 To n Step 1
            mouse_event(MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, mw, mh, 0, 0)
            mouse_event(MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, mw, mh, 0, 0)
        Next i

    End Sub

    Private Sub TabJian()
        SendKeys.SendWait("{TAB}")
        'System.Threading.Thread.Sleep(iSleep)
    End Sub
    Sub ShuRu(ByVal Str As String)
        SendKeys.SendWait(Str)

    End Sub

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
    Function GetKeys()
        Dim Items(38) As ListViewItem

        Items(0) = New ListViewItem(New String() {1, "Backspace", "{BACKSPACE}"})
        Items(1) = New ListViewItem((New String() {2, "Break", "{BREAK}"}))
        Items(2) = New ListViewItem((New String() {3, "Caps Lock", "{CAPSLOCK}"}))
        Items(3) = New ListViewItem((New String() {4, "Delete", "{DELETE}"}))
        Items(4) = New ListViewItem((New String() {5, "End", "{END}"}))
        Items(5) = New ListViewItem((New String() {6, "Enter", "{ENTER}"}))
        Items(6) = New ListViewItem((New String() {7, "Esc", "{ESC}"}))
        Items(7) = New ListViewItem((New String() {8, "F1", "{F1}"}))
        Items(8) = New ListViewItem((New String() {9, "F10", "{F10}"}))
        Items(9) = New ListViewItem((New String() {10, "F11", "{F11}"}))
        Items(10) = New ListViewItem((New String() {11, "F12", "{F12}"}))
        Items(11) = New ListViewItem((New String() {12, "F13", "{F13}"}))
        Items(12) = New ListViewItem((New String() {13, "F14", "{F14}"}))
        Items(13) = New ListViewItem((New String() {14, "F15", "{F15}"}))
        Items(14) = New ListViewItem((New String() {15, "F16", "{F16}"}))
        Items(15) = New ListViewItem((New String() {16, "F2", "{F2}"}))
        Items(16) = New ListViewItem((New String() {17, "F3", "{F3}"}))
        Items(17) = New ListViewItem((New String() {18, "F4", "{F4}"}))
        Items(18) = New ListViewItem((New String() {19, "F5", "{F5}"}))
        Items(19) = New ListViewItem((New String() {20, "F6", "{F6}"}))
        Items(20) = New ListViewItem((New String() {21, "F7", "{F7}"}))
        Items(21) = New ListViewItem((New String() {22, "F8", "{F8}"}))
        Items(22) = New ListViewItem((New String() {23, "F9", "{F9}"}))
        Items(23) = New ListViewItem((New String() {24, "Help", "{HELP}"}))
        Items(24) = New ListViewItem((New String() {25, "Home", "{HOME}"}))
        Items(25) = New ListViewItem((New String() {26, "Insert", "{INSERT}"}))
        Items(26) = New ListViewItem((New String() {27, "Num Lock", "{NUMLOCK}"}))
        Items(27) = New ListViewItem((New String() {28, "Page Down", "{PGDN}"}))
        Items(28) = New ListViewItem((New String() {29, "Page Up", "{PGUP}"}))
        Items(29) = New ListViewItem((New String() {30, "Scroll Lock", "{SCROLLLOCK}"}))
        Items(30) = New ListViewItem((New String() {31, "Tab", "{TAB}"}))
        Items(31) = New ListViewItem((New String() {32, "数字键盘乘号", "{MULTIPLY}"}))
        Items(32) = New ListViewItem((New String() {33, "数字键盘除号", "{DIVIDE}"}))
        Items(33) = New ListViewItem((New String() {34, "数字键盘加号", "{ADD}"}))
        Items(34) = New ListViewItem((New String() {35, "数字键盘减号", "{SUBTRACT}"}))
        Items(35) = New ListViewItem((New String() {36, "向上键", "{UP}"}))
        Items(36) = New ListViewItem((New String() {37, "向下键", "{DOWN}"}))
        Items(37) = New ListViewItem((New String() {38, "向右键", "{RIGHT}"}))
        Items(38) = New ListViewItem((New String() {39, "向左键", "{LEFT}"}))



        Return Items
    End Function
    'Backspace	{BACKSPACE}
    'Break	{BREAK}
    'Caps Lock	{CAPSLOCK}
    'Delete	{DELETE}
    'End	{END}
    'Enter	{ENTER}
    'Esc	{ESC}
    'F1	{F1}
    'F10	{F10}
    'F11	{F11}
    'F12	{F12}
    'F13	{F13}
    'F14	{F14}
    'F15	{F15}
    'F16	{F16}
    'F2	{F2}
    'F3	{F3}
    'F4	{F4}
    'F5	{F5}
    'F6	{F6}
    'F7	{F7}
    'F8	{F8}
    'F9	{F9}
    'Help	{HELP}
    'Home	{HOME}
    'Insert	{INSERT}
    'Num Lock	{NUMLOCK}
    'Page Down	{PGDN}
    'Page Up	{PGUP}
    'Scroll Lock	{SCROLLLOCK}
    'Tab	{TAB}
    '数字键盘乘号	{MULTIPLY}
    '数字键盘除号	{DIVIDE}
    '数字键盘加号	{ADD}
    '数字键盘减号	{SUBTRACT}
    '向上键	{UP}
    '向下键	{DOWN}
    '向右键	{RIGHT}
    '向左键	{LEFT}


End Module

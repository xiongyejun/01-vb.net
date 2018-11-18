Imports System.Reflection
Imports System.Threading, System.ComponentModel
Imports System.Runtime.InteropServices


Public Class CHook

#Region "封送结构"
    Private Structure MouseHookStruct
        Dim pt As Point
        Dim hwnd As Integer
        Dim wHitTestCode As Integer
        Dim dwExtraInfo As Integer
    End Structure

    Private Structure MouseLLHookStruct
        Dim pt As Point
        Dim MouseData As Integer
        Dim Flags As Integer
        Dim Time As Integer
        Dim dwExtraInfo As Integer
    End Structure

    Private Structure KeyboardHookStruct
        Dim vkCode As Integer  '1到254间的虚拟键盘码
        Dim SCANcODE As Integer    '扫描码
        Dim flags As Integer
        Dim timer As Integer
        Dim dwExtraInfo As Integer
    End Structure

    'Private Structure POINT
    '    Public x As Integer
    '    Public y As Integer
    'End Structure

#End Region

#Region "API声明"
    '安装钩子过程
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Integer, ByVal lpfn As HookProc, ByVal hmod As IntPtr, ByVal dwThreadId As Integer) As Integer
    '从钩子链中删除钩子
    Private Declare Function UnhookWindowsHookEx Lib "user32" Alias "UnhookWindowsHookEx" (ByVal hHook As Integer) As Integer
    '调中链中的下一个挂钩过程
    Private Declare Function CallNextHookEx Lib "user32" Alias "CallNextHookEx" (ByVal hHook As Integer, ByVal ncode As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As Integer
    '根据当前的扫描码和键盘信息，将一个虚拟键转换成ASCII字符
    Private Declare Function ToAscii Lib "user32" Alias "ToAscii" (ByVal uVirtKey As Integer, ByVal uScanCode As Integer, ByVal lpbKeyState As Byte(), ByVal lpwTransKey As Byte(), ByVal fuState As Integer) As Integer
    '取得键盘上每个虚拟键当前的状态
    Private Declare Function GetKeyboardState Lib "user32" Alias "GetKeyboardState" (ByVal pbKeyState As Byte()) As Integer
    '针对已处理过的按键，在最近一次输入信息时，判断指定虚拟键的状态
    Private Declare Function GetKeyState Lib "user32" Alias "GetKeyState" (ByVal nVirtKey As Integer) As Short
    Public Delegate Function HookProc(ByVal nCode As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As Integer
    '获取一个应用程序或动态链接库的模块句柄
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As IntPtr

#End Region

#Region "常量声明"
    '钩子常量
    Private Const WH_MOUSE_LL As Integer = 14
    Private Const WH_KEYBOARD_LL As Integer = 13

    Private Const WH_MOUSE As Integer = 7
    Private Const WH_KEYBOARD As Integer = 2

    Private Const WM_MOUSEMOVE As Integer = &H200
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_LBUTTONUP As Integer = &H202
    Private Const WM_LBUTTONDBLCLK As Integer = &H203

    Private Const WM_RBUTTONDOWN As Integer = &H204
    Private Const WM_RBUTTONUP As Integer = &H205
    Private Const WM_RBUTTONDBLCLK As Integer = &H206

    Private Const WM_MBUTTONDOWN As Integer = &H207
    Private Const WM_MBUTTONUP As Integer = &H208
    Private Const WM_MBUTTONDBLCLK As Integer = &H209
    Private Const WM_MOUSEWHEEL As Integer = &H20A

    Private Const WM_KEYDOWN As Integer = &H100
    Private Const WM_KEYUP As Integer = &H101
    Private Const WM_SYSKEYDOWN As Integer = &H104
    Private Const WM_SYSKEYUP As Integer = &H105

    Private Const VK_SHIFT As Integer = &H10
    Private Const VK_CAPITAL As Integer = &H14
    Private Const VK_NUMLOCK As Integer = &H90
#End Region

#Region "事件委托"
    Private events As New System.ComponentModel.EventHandlerList

    '鼠标激活事件
    Public Custom Event MouseActivity As MouseEventHandler
        AddHandler(ByVal value As MouseEventHandler)
            events.AddHandler("MouseActivity", value)
        End AddHandler

        RemoveHandler(ByVal value As MouseEventHandler)
            events.RemoveHandler("MouseActivity", value)
        End RemoveHandler

        RaiseEvent(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
            Dim eh As MouseEventHandler = TryCast(events("MouseActivity"), MouseEventHandler)
            If eh IsNot Nothing Then eh.Invoke(sender, e)
        End RaiseEvent
    End Event

    '键盘按下事件
    Public Custom Event KeyDown As KeyEventHandler
        AddHandler(ByVal value As KeyEventHandler)
            events.AddHandler("KeyDown", value)
        End AddHandler

        RemoveHandler(ByVal value As KeyEventHandler)
            events.RemoveHandler("KeyDown", value)
        End RemoveHandler

        RaiseEvent(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            Dim eh As KeyEventHandler = TryCast(events("KeyDown"), KeyEventHandler)
            If eh IsNot Nothing Then eh.Invoke(sender, e)
        End RaiseEvent
    End Event

    '键盘输入事件
    Public Custom Event KeyPress As KeyPressEventHandler
        AddHandler(ByVal value As KeyPressEventHandler)
            events.AddHandler("KeyPress", value)
        End AddHandler

        RemoveHandler(ByVal value As KeyPressEventHandler)
            events.RemoveHandler("KeyPress", value)
        End RemoveHandler

        RaiseEvent(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
            Dim eh As KeyPressEventHandler = TryCast(events("KeyPress"), KeyPressEventHandler)
            If eh IsNot Nothing Then eh.Invoke(sender, e)
        End RaiseEvent
    End Event

    '键盘松开事件
    Public Custom Event KeyUp As KeyEventHandler
        AddHandler(ByVal value As KeyEventHandler)
            events.AddHandler("KeyUp", value)
        End AddHandler

        RemoveHandler(ByVal value As KeyEventHandler)
            events.RemoveHandler("KeyUp", value)
        End RemoveHandler

        RaiseEvent(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            Dim eh As KeyEventHandler = TryCast(events("KeyUp"), KeyEventHandler)
            If eh IsNot Nothing Then eh.Invoke(sender, e)
        End RaiseEvent
    End Event
#End Region


    '钩子句柄
    Private hMouseHook As Integer
    Private hKeyboardHook As Integer
    Private Shared MouseHookProcedure As HookProc
    Private Shared KeyboardHookProcedure As HookProc
    '事件类型

#Region "创建与析构类型"
    '创建一个全局鼠标键盘钩子（请使用Start方法开始监视）
    Sub New()
        '留空即可
    End Sub

    '创建一个全局鼠标键盘钩子，决定是否安装钩子
    'InstallAll是否立刻安装钩子
    Sub New(ByVal InstallAll As Boolean)
        If InstallAll Then StartHook(True, True)
    End Sub

    '创建一个全局鼠标键盘钩子，决定钩子类型
    '
    Sub New(ByVal InstallKeyboard As Boolean, ByVal InstallMouse As Boolean)
        StartHook(InstallKeyboard, InstallMouse)
    End Sub

    '析构函数
    Protected Overrides Sub Finalize()
        UnHook()
        MyBase.Finalize()
    End Sub
#End Region

    '开始安装系统钩子
    Public Sub StartHook(Optional ByVal InstallKeyboardHook As Boolean = True, Optional ByVal InstallMouseHook As Boolean = False)
        '注册键盘钩子
        If InstallKeyboardHook AndAlso hKeyboardHook = 0 Then
            KeyboardHookProcedure = New HookProc(AddressOf Me.KeyboardHookProc)
            'hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, KeyboardHookProcedure, Marshal.GetHINSTANCE([Assembly].GetExecutingAssembly.GetModules()(0)), 0)
            hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, KeyboardHookProcedure, GetModuleHandle(Process.GetCurrentProcess().MainModule.ModuleName), 0)
            If (hKeyboardHook) = 0 Then
                UnHook(True, False)
                Throw New Win32Exception(Marshal.GetLastWin32Error)
            End If
        End If

        '注册鼠标钩子
        If InstallMouseHook AndAlso hMouseHook = 0 Then
            MouseHookProcedure = New HookProc(AddressOf MouseHookPro)
            'hMouseHook = SetWindowsHookEx(WH_KEYBOARD_LL, MouseHookProcedure, Marshal.GetHINSTANCE([Assembly].GetExecutingAssembly.GetModules()(0)), 0)
            hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, MouseHookProcedure, GetModuleHandle(Process.GetCurrentProcess().MainModule.ModuleName), 0)
            If (hMouseHook) = 0 Then
                UnHook(False, True)
                Throw New Win32Exception(Marshal.GetLastWin32Error)
            End If
        End If
    End Sub

    '立刻卸载系统钩子
    Public Sub UnHook(Optional ByVal UnInstallKeyboardHook As Boolean = True, Optional ByVal UnInstallMouseHook As Boolean = False, Optional ByVal ThrowExceptions As Boolean = False)
        If hKeyboardHook <> 0 AndAlso UnInstallKeyboardHook Then
            Dim retKyeboard As Integer = UnhookWindowsHookEx(hKeyboardHook)
            hKeyboardHook = 0
            If ThrowExceptions AndAlso retKyeboard = 0 Then
                Throw New Win32Exception(Marshal.GetLastWin32Error)
            End If
        End If

        If hMouseHook <> 0 AndAlso UnInstallMouseHook Then
            Dim retMouse As Integer = UnhookWindowsHookEx(hMouseHook)
            hMouseHook = 0
            If ThrowExceptions AndAlso retMouse = 0 Then
                Throw New Win32Exception(Marshal.GetLastWin32Error)
            End If
        End If
    End Sub

    '键盘消息的委托处理代码
    Private Function KeyboardHookProc(ByVal nCode As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As Integer
        Static Handled As Boolean : Handled = False
        If nCode >= 0 AndAlso (events("KeyDown") IsNot Nothing OrElse events("KeyPress") IsNot Nothing OrElse events("KeyUp") IsNot Nothing) Then
            Static MykeyboardhookStruct As KeyboardHookStruct
            MykeyboardhookStruct = DirectCast(Marshal.PtrToStructure(lParam, GetType(KeyboardHookStruct)), KeyboardHookStruct)

            '激活KeyDown
            If wParam = WM_KEYDOWN OrElse wParam = WM_SYSKEYDOWN Then
                Dim e As New KeyEventArgs(MykeyboardhookStruct.vkCode)
                RaiseEvent KeyDown(Me, e)
                Handled = Handled Or e.Handled  '是否取消下一个钩子
            End If

            '激活KeyUp
            If wParam = WM_KEYUP OrElse wParam = WM_SYSKEYUP Then
                Dim e As New KeyEventArgs(MykeyboardhookStruct.vkCode)
                RaiseEvent KeyUp(Me, e)
                Handled = Handled Or e.Handled
            End If

            'keyPress
            If wParam = WM_KEYDOWN Then
                Dim isDownShift As Boolean = (GetKeyState(VK_SHIFT) & &H80 = &H80)
                Dim isDownCapslock As Boolean = (GetKeyState(VK_CAPITAL) <> 0)
                Dim keyState(256) As Byte
                GetKeyboardState(keyState)
                Dim inBuffer(2) As Byte
                If ToAscii(MykeyboardhookStruct.vkCode, MykeyboardhookStruct.SCANcODE, keyState, inBuffer, MykeyboardhookStruct.flags) = 1 Then
                    Static key As Char : key = Chr(inBuffer(0))
                    Dim e As KeyPressEventArgs = New KeyPressEventArgs(key)
                    RaiseEvent KeyPress(Me, e)
                    Handled = Handled Or e.Handled
                End If
            End If

            '取消或者激活下一个钩子
            If Handled Then
                Return 1
            Else
                Return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam)
            End If

        End If

    End Function

    '鼠标消息的委托处理代码
    Private Function MouseHookPro(ByVal nCode As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As Integer
        If nCode >= 0 AndAlso events("MouseActivity") IsNot Nothing Then
            Static MousehookStruct As MouseLLHookStruct
            MousehookStruct = DirectCast(Marshal.PtrToStructure(lParam, GetType(MouseLLHookStruct)), MouseLLHookStruct)
            Static moubut As MouseButtons : moubut = MouseButtons.None  '鼠标按键
            Static mouseDelta As Integer : mouseDelta = 0 '滚轮值
            Select Case wParam
                Case WM_LBUTTONDOWN
                    moubut = MouseButtons.Left
                Case WM_RBUTTONDOWN
                    moubut = MouseButtons.Right
                Case WM_MBUTTONDOWN
                    moubut = MouseButtons.Middle
                Case WM_MOUSEWHEEL
                    Static int As Integer : int = (MousehookStruct.MouseData >> 16) And &HFFF '本段代码CLE添加，模仿c#的short从int弃位转换
                    If int > Short.MaxValue Then
                        mouseDelta = int - 65536
                    Else
                        mouseDelta = int
                    End If

            End Select

            Dim clickCount As Integer : clickCount = 0  '单击次数
            If moubut <> MouseButtons.None Then
                If wParam = WM_LBUTTONDBLCLK OrElse wParam = WM_MBUTTONDBLCLK OrElse wParam = WM_RBUTTONDBLCLK Then
                    clickCount = 2
                Else
                    clickCount = 1
                End If
            End If

            '从回调函数中得到鼠标的消息
            Dim e As MouseEventArgs = New MouseEventArgs(moubut, clickCount, MousehookStruct.pt.X, MousehookStruct.pt.Y, mouseDelta)
            RaiseEvent MouseActivity(Me, e)
        End If
        Return CallNextHookEx(hMouseHook, nCode, wParam, lParam)
    End Function

    '键盘钩子是否有效
    Public Property KeyHookEnabled() As Boolean
        Get
            Return hKeyboardHook <> 0
        End Get
        Set(ByVal value As Boolean)
            If value Then StartHook(True, False) Else UnHook(True, False)
        End Set
    End Property

    '鼠标钩子是否有效
    Public Property MouseHookEnabled() As Boolean
        Get
            Return hMouseHook <> 0
        End Get
        Set(ByVal value As Boolean)
            If value Then StartHook(False, True) Else UnHook(False, True)
        End Set
    End Property


End Class


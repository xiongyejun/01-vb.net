Friend Class PlantsVsZombies

#Region "WinApi"
    Private Declare Function GetWindowThreadProcessId Lib "user32" _
                    Alias "GetWindowThreadProcessId" (ByVal hWnd As IntPtr, _
                                                      ByRef lpdwProcessId As Integer) As Integer

    Private Declare Function OpenProcess Lib "kernel32" _
                    Alias "OpenProcess" (ByVal dwDesiredAccess As Integer, _
                                         ByVal bInheritHandle As Boolean, _
                                         ByVal dwProcessId As Integer) As IntPtr


    Private Declare Function ReadProcessMemory Lib "kernel32" _
                    Alias "ReadProcessMemory" (ByVal hProcess As IntPtr, _
                                               ByVal ByvallpBaseAddress As Integer, _
                                               ByRef lpBuffer As Integer, _
                                               ByVal nSize As Integer, _
                                               ByRef lpNumberOfBytesWritten As Integer) As Integer


    Private Declare Function WriteProcessMemory Lib "kernel32" _
                    Alias "WriteProcessMemory" (ByVal hProcess As IntPtr, _
                                                ByVal lpBaseAddress As Integer, _
                                                ByRef lpBuffer As Integer, _
                                                ByVal nSize As Integer, _
                                                ByRef lpNumberOfBytesWritten As Integer) As Integer

    Private Declare Function FindWindow Lib "user32" _
                    Alias "FindWindowA" (ByVal lpClassName As String, _
                                         ByVal lpWindowName As String) As Integer


    Private Declare Function SetWindowText Lib "user32" _
                    Alias "SetWindowTextA" (ByVal hWnd As IntPtr, _
                                            ByVal lpString As String) As Integer

    Private Declare Function InjectDllToGame Lib "Bombs" _
                    Alias "InjectDllToGame" () As Boolean

    Private Const SYNCHRONIZE = &H100000
    Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
    Private Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
    Private Const WM_LBUTTONDOWN = &H201
    Private Const WM_LBUTTONUP = &H202

#End Region

#Region "私有变量"
    Private Const BaseAddress As Int32 = &H7794F8 '游戏基址
    Private MyhProcess As IntPtr = IntPtr.Zero    '游戏进程句柄
    Private MyHwnd As IntPtr = IntPtr.Zero        '游戏窗体句柄
    Private Const MyGameClassName As String = "Plants vs. Zombies 1.2.0.1073 RELEASE"
    Private WithEvents MyTimer As New Timer

    Sub New()
        With MyTimer
            .Enabled = True
            .Interval = 10000
        End With
        UpdataHwndAndProcess()
    End Sub

    '每秒更新一次进程句柄
    Private Sub MyTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyTimer.Tick
        UpdataHwndAndProcess()
    End Sub
#End Region


#Region "基本函数"

    Private ReadOnly Property hProcess() As IntPtr '取得游戏进程句柄
        Get
            Return MyhProcess
        End Get
    End Property

    Private ReadOnly Property Hwnd() As IntPtr
        Get
            Return MyHwnd
        End Get
    End Property


    Private Sub UpdataHwndAndProcess()
        Dim NewHwnd As IntPtr = FindWindow(vbNullString, MyGameClassName)

        If (NewHwnd <> MyHwnd) And (NewHwnd <> IntPtr.Zero) Then
            MyHwnd = NewHwnd

            Dim Pid As Integer = 0
            GetWindowThreadProcessId(MyHwnd, Pid)
            MyhProcess = OpenProcess(PROCESS_ALL_ACCESS, False, Pid)

        End If
    End Sub


    Private Function GetAddress(ByVal Offsets() As Int32) As Int32 '从基址和偏移量计算 二、三级地址
        Dim Buffer, ByteRead As Int32
        Dim NewAddress As Int32 = BaseAddress

        For i As Int32 = 0 To Offsets.Length - 1
            ReadProcessMemory(Me.hProcess(), NewAddress, Buffer, 4, ByteRead)
            NewAddress = Buffer + Offsets(i)
        Next

        Return NewAddress
    End Function

    Private Function GetValue(ByVal Address As Int32) As Int32 '从地址中取得内容
        Dim Buffer, ByteRead As Int32
        ReadProcessMemory(Me.hProcess(), Address, Buffer, 4, ByteRead)
        Return Buffer
    End Function

    Private Sub SetValue(ByVal Address As Int32, ByVal Value As Int32) '往地址中定入内容
        Dim ByteRead As Int32 = 0
        Dim Buffer As Integer = Value
        WriteProcessMemory(Me.hProcess(), Address, Buffer, 4, ByteRead)
    End Sub

#End Region



#Region "二级基本函数"

    Private ReadOnly Property SunshineAddress() As Int32 '阳光地址  
        Get
            Return GetAddress(New Int32() {&H868, &H5578})
        End Get
    End Property

    Private ReadOnly Property MoneyAddress() As Int32 '金钱地址
        Get
            Return GetAddress(New Int32() {&H950, &H50})
        End Get
    End Property




    Private ReadOnly Property CountOfPlantsAddress() As Integer '可选植物栏植物总数地址
        Get
            Return GetAddress(New Integer() {&H868, &H15C, &H24})
        End Get
    End Property


    Private ReadOnly Property CountOfPlants() As Integer '取得可选植物栏植物总数
        Get
            Return GetValue(CountOfPlantsAddress)
        End Get
    End Property

    Private ReadOnly Property PlantCoolDownTimeAddress(ByVal PlantIndex As Integer) As Integer '可选植物栏植物冷却时间地址
        Get
            Return GetAddress(New Integer() {&H868, &H15C, &H50 * (PlantIndex) + &H50})
        End Get
    End Property

    Private ReadOnly Property PlantTypeAddress(ByVal PlantIndex As Integer) As Integer '可选植物栏植物类型地址
        Get
            Return GetAddress(New Integer() {&H868, &H15C, &H50 * (PlantIndex) + &H5C})
        End Get
    End Property




    Private ReadOnly Property CountOfZombiesAddress() As Integer '僵尸数量地址
        Get
            Return GetAddress(New Integer() {&H868, &HB8})
        End Get
    End Property


    Private ReadOnly Property TotalCountOfZombiesAddress() As Integer '僵尸数量上限地址
        Get
            Return GetAddress(New Int32() {&H868, &HAC})
        End Get
    End Property



    Private ReadOnly Property CountOfVaseAddress() As Integer '花瓶数量地址
        Get
            Return GetAddress(New Integer() {&H868, &H144})
        End Get
    End Property


    Private ReadOnly Property CountOfVase() As Integer '花瓶数量
        Get
            Return GetValue(CountOfVaseAddress)
        End Get
    End Property

    Private ReadOnly Property VasePerspectiveAddress(ByVal VaseIndex As Integer) As Integer '花瓶透视地址
        Get
            Return GetAddress(New Integer() {&H868, &H134, &HEC * VaseIndex + &H4C})
        End Get
    End Property

    Private ReadOnly Property ProgressBarAddress() As Integer '进度条地址
        Get
            Return GetAddress(New Integer() {&H868, &H5628})
        End Get
    End Property

    Private ReadOnly Property PlantCurrentBloodAddress(ByVal Index As Integer) As Integer '种下的植物当前血值地址
        Get
            Return GetAddress(New Integer() {&H868, &HC4, &H14C * Index + &H40})
        End Get
    End Property

    Private ReadOnly Property PlantMaxBloodAddress(ByVal Index As Integer) As Integer '种下的植物血值上限地址
        Get
            Return GetAddress(New Integer() {&H868, &HC4, &H14C * Index + &H44})
        End Get
    End Property

    Private ReadOnly Property PlantCurrentHitTimeAddress(ByVal Index As Integer) As Integer '种下的植物当前攻击倒计时地址
        Get
            Return GetAddress(New Integer() {&H868, &HC4, &H14C * Index + &H58})
        End Get
    End Property

    Private ReadOnly Property PlantTotalHitTimeAddress(ByVal Index As Integer) As Integer '种下的植物攻击倒计时上限地址
        Get
            Return GetAddress(New Integer() {&H868, &HC4, &H14C * Index + &H5C})
        End Get
    End Property

    Private ReadOnly Property PlantGrowthIndexAddress(ByVal Index As Integer) As Integer '种下的植物的类型地址
        Get
            Return GetAddress(New Integer() {&H868, &HC4, &H14C * Index + &H24})
        End Get
    End Property



#End Region




#Region "对外开公函数"

    Public Property SunshineValue() As Int32 '获取或设置阳光
        Get
            Return GetValue(SunshineAddress)
        End Get
        Set(ByVal value As Int32)
            SetValue(SunshineAddress, value)
        End Set
    End Property

    Public Property MoneyValue() As Int32 '获取或设置金钱
        Get
            Return GetValue(MoneyAddress)
        End Get
        Set(ByVal value As Int32)
            SetValue(MoneyAddress, value)
        End Set
    End Property


    Public Sub AllPlantsCoolDownDisEnable() '将植物冷却时间设置为0
        Dim Count As Integer = CountOfPlants
        For i As Integer = 0 To Count - 1
            SetValue(PlantCoolDownTimeAddress(i), 0)
        Next
    End Sub



    Public ReadOnly Property CountOfZombies() As Integer '僵尸数量
        Get
            Return GetValue(CountOfZombiesAddress)
        End Get
    End Property

    Public ReadOnly Property TotalCountOfZombies() As Integer '僵尸上限数量
        Get
            Return GetValue(TotalCountOfZombiesAddress)
        End Get
    End Property

    Public Sub ShowTextAtGameTitle(ByVal Text As String) '在游戏标题处显示信息
        SetWindowText(Hwnd, Text)
    End Sub

    Public Sub VasePerspective()           '花瓶透视
        Dim Count As Integer = CountOfVase
        For i As Integer = 0 To Count - 1
            SetValue(VasePerspectiveAddress(i), 100)
        Next
    End Sub

    Public ReadOnly Property ProgressBarValue() '进度条
        Get
            Return GetValue(ProgressBarAddress)
        End Get
    End Property

    Public Sub FillPlantsFullBlood() '植物不死
        Dim MaxBlood As Integer = 0
        Dim PlantIndex As Integer = 0

        For i As Integer = 0 To 6 * 9 - 1 '枚举种下的植物
            MaxBlood = GetValue(PlantMaxBloodAddress(i)) '获取植物的血值上限
            SetValue(PlantCurrentBloodAddress(i), MaxBlood) '给它满血

            PlantIndex = GetValue(PlantGrowthIndexAddress(i))
            Select Case PlantIndex
                Case 0, 5, 7, 8, 10, 18, 26, 28, 32, 34, 39, 40, 44, 47  '这些植物
                    SetValue(PlantTotalHitTimeAddress(i), 40)            '40的上限倒计时
            End Select
        Next
    End Sub

#End Region


End Class
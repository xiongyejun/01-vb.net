Imports System.Runtime.InteropServices
Imports System.IO

Public Class CCompdocFile
#Region "定义"
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal Destination As IntPtr, ByVal Source As IntPtr, ByVal Length As Integer)

    Const CFHEADER_SIZE As Integer = 2 ^ 9
    Const DIR_SIZE As Integer = 128

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Public Structure CFHeader
        <VBFixedArray(7), MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)>
        Dim id() As Byte                   '文档标识id  
        <VBFixedArray(15), MarshalAs(UnmanagedType.ByValArray, SizeConst:=16)>
        Dim file_id() As Byte             '文件唯一标识 
        Dim file_format_revision As Short '文件格式修订号
        Dim file_format_version As Short  '文件格式版本号
        Dim memory_endian As Short        'FFFE表示 Little-Endian
        Dim sector_size As Short          '扇区的大小 2的幂 通常为2^9=512
        Dim short_sector_size As Short    '短扇区大小，2的幂,通常为2^6
        <VBFixedArray(9), MarshalAs(UnmanagedType.ByValArray, SizeConst:=10)>
        Dim not_used_1() As Byte           '
        Dim SAT_count As Integer               '分区表扇区的总数
        Dim dir_first_SID As Integer           '目录流第一个扇区的ID
        <VBFixedArray(3), MarshalAs(UnmanagedType.ByValArray, SizeConst:=4)>
        Dim not_used_2() As Byte                '
        Dim min_stream_size As Integer         '最小标准流
        Dim SSAT_first_SID As Integer          '短分区表的第一个扇区ID
        Dim SSAT_count As Integer              '短分区表扇区总数
        Dim MSAT_first_SID As Integer          '主分区表的第一个扇区ID
        Dim MSAT_count As Integer              '分区表的扇区总数
        <VBFixedArray(108), MarshalAs(UnmanagedType.ByValArray, SizeConst:=109)>
        Dim arr_SID() As Integer            '主分区表前109个记录  108字节
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Private Structure CFDir
        <VBFixedArray(63), MarshalAs(UnmanagedType.ByValArray, SizeConst:=64)>
        Dim dir_name() As Byte              '
        Dim len_name As Short
        Dim type As Byte                    '1仓storage 2流 5根
        Dim color As Byte                    '0红色 1黑色
        Dim left_child As Integer               '-1表示叶子
        Dim right_child As Integer
        Dim sub_dir As Integer
        <VBFixedArray(19), MarshalAs(UnmanagedType.ByValArray, SizeConst:=20)>
        Dim arr_keep() As Byte          '
        Dim time_create As Date
        Dim time_modify As Date
        Dim first_SID As Integer               '目录入口所表示的第1个扇区编码
        Dim stream_size As Integer             '目录入口流尺寸，可判断是否是短扇区
        Dim not_used As Integer
    End Structure

    '0-名称，1-字节开始的地方，2-占用的字节个数，3到n实际每个字符的地址
    Structure ModuleAddress
        Dim ModuleName As String
        Dim StartAddress As Integer
        Dim Size As Integer
        Dim ArrAddress() As Integer
    End Structure

    Dim file_byte() As Byte
    Dim FileAddress As IntPtr   'file_byte的内存地址
    Public arr_MSAT() As Integer              '主分区表数组，指向的是存储分区表的SID
    Dim arr_SAT() As Integer               '分区表数组，指向的是下一个SID
    Dim arr_SSAT() As Integer              '短分区表数据
    Dim arr_dir() As CFDir              '目录
    Public cf_header As CFHeader '文件头
    Public arr_Module() As ModuleAddress  '模块的信息
#End Region

    Private my_path As String
    Public Property path() As String
        Get
            Return my_path
        End Get
        Set(ByVal value As String)
            my_path = value
        End Set
    End Property

    Private b_ready As Boolean
    Public ReadOnly Property ready() As Boolean
        Get
            Return b_ready
        End Get
    End Property

    Public Sub New(file_name As String)
        cf_header = New CFHeader
        ReDim cf_header.id(7)
        ReDim cf_header.file_id(15)
        ReDim cf_header.not_used_1(9)
        ReDim cf_header.not_used_2(3)
        ReDim cf_header.arr_SID(108)

        Me.path = file_name

        If Me.path <> "" Then
            If GetFileByte() = 1 Then
                b_ready = True
                GetCfHeader()

                GetMSAT()

                GetSAT()

                GetDir()

                getSSAT()

            End If
        Else
            b_ready = False
        End If
    End Sub
    '获取文件的前面512个字节
    Function GetCfHeader()
        cf_header = Marshal.PtrToStructure(FileAddress, cf_header.GetType)

        'CopyMemory VarPtr(cf_header.id(0)), VarPtr(file_byte(0)), CFHEADER_SIZE
        Return 0
    End Function
    '获取主分区表
    Private Function GetMSAT()
        Dim i As Integer
        Dim arr(127) As Integer
        Dim next_SID As Integer
        Dim flag As Boolean
        Dim count_MSAT As Integer

        With cf_header
            ReDim arr_MSAT(.SAT_count - 1)

            '获取头文件中的109个
            For i = 0 To 108
                If .arr_SID(i) = -1 Then
                    '头中并没有109个
                    Return 0
                    Exit Function
                End If

                arr_MSAT(i) = .arr_SID(i)
            Next i
            '获取另外的
            count_MSAT = 109
            next_SID = .MSAT_first_SID
            flag = True

            Do
                Dim p1 As IntPtr = GCHandle.Alloc(arr, GCHandleType.Pinned).AddrOfPinnedObject()
                CopyMemory(p1, FileAddress + CFHEADER_SIZE + CFHEADER_SIZE * next_SID, CFHEADER_SIZE)

                For i = 0 To 127 - 1
                    If arr(i) = -1 Then
                        flag = False
                        Exit For
                    End If

                    arr_MSAT(count_MSAT) = arr(i)
                    count_MSAT = count_MSAT + 1
                Next i
                next_SID = arr(i)       'SID的最后4个字节存储再下一个的SID
            Loop While flag

        End With

        Return 0
    End Function
    '获取分区表
    Private Function GetSAT()
        Dim i As Integer, j As Integer
        Dim k_SAT As Integer
        Dim arr(127) As Integer

        With cf_header
            ReDim arr_SAT(.SAT_count * 128 - 1)

            k_SAT = 0
            For i = 0 To .SAT_count - 1
                Dim p1 As IntPtr = GCHandle.Alloc(arr, GCHandleType.Pinned).AddrOfPinnedObject()
                CopyMemory(p1, FileAddress + CFHEADER_SIZE + CFHEADER_SIZE * arr_MSAT(i), CFHEADER_SIZE)

                For j = 0 To 127
                        arr_SAT(k_SAT) = arr(j)
                        k_SAT = k_SAT + 1
                    Next j
                Next i

        End With

        Return 0
    End Function
    '获取目录
    Private Function GetDir() As Integer
        Dim l_SID As Integer
        Dim k As Integer
        Dim d As Date = #2017-1-1#

        l_SID = cf_header.dir_first_SID

        k = 0
        Do
            ReDim Preserve arr_dir(k)
            RedimDir(arr_dir(k))

            '首先是找到SID的地址，然后1个sector存放4个dir，1个偏移DIR_SIZE
            arr_dir(k) = Marshal.PtrToStructure(FileAddress + CFHEADER_SIZE + CFHEADER_SIZE * l_SID + DIR_SIZE * (k Mod 4), arr_dir(k).GetType)

            k = k + 1
            If k Mod 4 = 0 Then
                l_SID = arr_SAT(l_SID)
            End If
        Loop Until l_SID = -2

        Return 0
    End Function

    Private Function RedimDir(ByRef d As CFDir)
        ReDim d.dir_name(63)
        ReDim d.arr_keep(19)

        Return 0
    End Function
    '获取短扇区分区表
    Private Function getSSAT()
        Dim l_SID As Integer
        Dim k As Integer

        If cf_header.SSAT_count = 0 Then
            Return 0
            Exit Function
        End If
        '根目录的 stream_size 表示短流存放流的大小，每64个为一个short sector
        ReDim arr_SSAT(arr_dir(0).stream_size / 64 - 1)

        l_SID = arr_dir(0).first_SID    '短流起始SID
        For k = 0 To arr_dir(0).stream_size / 64 - 1
            arr_SSAT(k) = l_SID * CFHEADER_SIZE + CFHEADER_SIZE + (k Mod 8) * 64  '指向偏移地址，实际地址要加上VarPtr(file_byte(0))
            If (k + 1) Mod 8 = 0 Then  '到下一个SID
                l_SID = arr_SAT(l_SID)
            End If
        Next k

        Return 0
    End Function

    Function DirToArr(ByRef arr(,) As String) As Integer
        ReDim arr(UBound(arr_dir), 13 - 1 - 1 - 1 + 1)  '13个属性，-arr_keep,-notused +序号
        Dim i_col As Integer
        Dim k_dir As Integer

        For k_dir = 0 To UBound(arr_dir)
            i_col = 0
            With arr_dir(k_dir)
                arr(k_dir, i_col) = my_hex(k_dir)
                i_col = i_col + 1

                arr(k_dir, i_col) = System.Text.Encoding.Unicode.GetString(.dir_name).TrimEnd
                'Debug.Print(arr(k_dir, i_col))
                'arr(k_dir, i_col) = VBA.Left$(.dir_name, .len_name)
                i_col = i_col + 1

                arr(k_dir, i_col) = .len_name
                i_col = i_col + 1

                arr(k_dir, i_col) = .type
                i_col = i_col + 1

                arr(k_dir, i_col) = .color
                i_col = i_col + 1

                arr(k_dir, i_col) = my_hex(.left_child)
                i_col = i_col + 1

                arr(k_dir, i_col) = my_hex(.right_child)
                i_col = i_col + 1

                arr(k_dir, i_col) = my_hex(.sub_dir)
                i_col = i_col + 1

                arr(k_dir, i_col) = Format(.time_create, "yyyy/mm/dd")
                i_col = i_col + 1

                arr(k_dir, i_col) = Format(.time_modify, "yyyy/mm/dd")
                i_col = i_col + 1

                arr(k_dir, i_col) = my_hex(.first_SID)
                i_col = i_col + 1

                arr(k_dir, i_col) = my_hex(.stream_size)
                i_col = i_col + 1
            End With
        Next k_dir

        Return k_dir - 1
    End Function

    'arr_address 构建1个查找地址的数组，在查找模块的时候用，因为数据区域不一定是连续的
    '            第1列代表当前stream字节下标，第2列是对应的地址，如：
    '           0   100
    '           1   164
    '           2   228
    '也有可能是512的大小
    Function GetStream(dir_name As String, ByRef arr_result() As Byte, ByRef stream_len As Integer, ByRef arr_address(,) As Integer, ByRef if_short As Boolean) As Integer
        Dim i As Integer
        Dim str As String
        Dim n_size As Integer, j As Integer
        Dim short_SID As Integer
        Dim l_SID As Integer

        For i = 0 To UBound(arr_dir, 1)
            str = System.Text.Encoding.Unicode.GetString(arr_dir(i).dir_name)
            If StrComp(str, dir_name, vbTextCompare) = 0 Then Exit For
        Next i

        If i - 1 = UBound(arr_dir, 1) Then
            MsgBox("没有目录" & dir_name)
            Return -1
        End If

        If arr_dir(i).type <> 2 Then
            MsgBox("目录" & dir_name & "不是流。")
            Return -1
        End If

        If arr_dir(i).first_SID = -1 Then
            MsgBox("目录" & dir_name & "流大小为0。")
            Return 0
        End If

        stream_len = arr_dir(i).stream_size
        With arr_dir(i)
            If stream_len < cf_header.min_stream_size Then
                'short_sector
                if_short = True
                'n_size = .stream_size \ 64
                If (stream_len Mod 64) = 0 Then
                    n_size = stream_len \ 64 '512
                Else
                    n_size = stream_len \ 64 + 1
                End If

                ReDim arr_address(n_size - 1, 1)
                ReDim arr_result(n_size * 64 - 1)

                short_SID = .first_SID
                '            start_address = VarPtr(file_byte(0)) + arr_SSAT(short_SID)

                Dim p1 As IntPtr = GCHandle.Alloc(arr_result, GCHandleType.Pinned).AddrOfPinnedObject()
                For j = 1 To n_size
                    arr_address(j - 1, 0) = (j - 1)
                    arr_address(j - 1, 1) = arr_SSAT(short_SID + j - 1) 'VarPtr(file_byte(0))

                    CopyMemory(p1 + (j - 1) * 64, arr_address(j - 1, 1) + FileAddress, 64)
                Next j
            Else
                if_short = False
                If (stream_len Mod CFHEADER_SIZE) = 0 Then
                    n_size = stream_len \ CFHEADER_SIZE '512
                Else
                    n_size = stream_len \ CFHEADER_SIZE + 1
                End If

                ReDim arr_address(n_size - 1, 1)
                ReDim arr_result(n_size * CFHEADER_SIZE - 1)

                l_SID = .first_SID
                '            start_address = VarPtr(file_byte(0)) + arr_SAT(l_SID) * CFHEADER_SIZE + CFHEADER_SIZE

                Dim p1 As IntPtr = GCHandle.Alloc(arr_result, GCHandleType.Pinned).AddrOfPinnedObject()
                For j = 1 To n_size
                    'If j = 32 Then Stop
                    arr_address(j - 1, 0) = j - 1
                    arr_address(j - 1, 1) = l_SID * CFHEADER_SIZE + CFHEADER_SIZE  '+VarPtr(file_byte(0))
                    CopyMemory(p1 + (j - 1) * CFHEADER_SIZE, arr_address(j - 1, 1) + FileAddress, CFHEADER_SIZE)

                    l_SID = arr_SAT(l_SID)  'arr_SAT的下标是第i个，对应的值是下1个
                Next j

            End If
        End With

        Return n_size - 1
    End Function

    Function Workspace(arr() As String)
        '    Dim str_PROJECT As String
        '    Dim if_short As Boolean
        '    Dim str_workspace As String
        '    Dim re As Object
        '    Dim match_coll As Object
        '    Dim i As Integer
        '    Dim arr_address() As Integer

        '    str_PROJECT = Me.read_stream("PROJECT", arr_address, if_short)
        '    str_workspace = VBA.Split(str_PROJECT, "[Workspace]")(1)

        'Set re = CreateObject("VBScript.RegExp") 'Microsoft VBScript Tegular Expressions 5.5
        'With re
        '        .Global = True                  '搜索全部字符，false搜索到第1个即停止
        '        .MultiLine = False              '是否多行
        '        .IgnoreCase = False             '区分大小写
        '        .Pattern = "\S*=\d{2,3},\S*"
        '    Set match_coll = .Execute(str_workspace)            '返回MatchCollection对象
        'End With

        '    ReDim arr(match_coll.Count - 1) As String
        'For i = 0 To match_coll.Count - 1
        '        arr(i) = match_coll(i).Value
        '    Next i

        'Set re = Nothing
        'Set match_coll = Nothing

        Return 0
    End Function

    '在PROJECT的stream中，利用正则查找模块
    Function GetModule()
        Dim str_PROJECT As String
        Dim re As Object = Nothing
        Dim match_coll As Object = Nothing
        Dim i As Integer, j As Integer, k As Integer
        Dim arr_address(,) As Integer = Nothing
        Dim l_double_byte As Integer
        Dim this_double_byte As Integer
        Dim if_short As Boolean
        Dim step_address As Integer
        Dim str_hiden_module
        Dim arr_byte() As Byte = Nothing, stream_len As Integer

        '有可能存在隐藏的模块，形式如0D0A0D0A0D0A0D0A0D0A0D0A0D0A
        '至少包含8个长度(Module=)
        str_hiden_module = Chr(&HD) & Chr(&HA)
        str_hiden_module = str_hiden_module & str_hiden_module & str_hiden_module & str_hiden_module 'Module=
        str_hiden_module = str_hiden_module & "\s*"

        Me.GetStream("PROJECT", arr_byte, stream_len, arr_address, if_short)
        str_PROJECT = System.Text.Encoding.Default.GetString(arr_byte)
        'MsgBox(str_PROJECT)

        If if_short Then
            step_address = 64
        Else
            step_address = 512
        End If

        re = CreateObject("VBScript.RegExp") 'Microsoft VBScript Tegular Expressions 5.5
        With re
            .Global = True                  '搜索全部字符，false搜索到第1个即停止
            .MultiLine = False              '是否多行
            .IgnoreCase = False             '区分大小写
            .Pattern = "Module=\S*|Class=\S*|BaseClass=\S*|" & str_hiden_module       '搜素规则 |Class=.{1,}||BaseClass=.{1,}|
            match_coll = .Execute(str_PROJECT)            '返回MatchCollection对象
        End With

        If match_coll.Count = 0 Then
            MsgBox("没找到模块")
            Return 0
        End If

        ReDim arr_Module(match_coll.Count - 1) '0-名称，1-字节开始的地方，2-占用的字节个数，3到n实际每个字符的地址
        l_double_byte = 0                       '双字节字符的个数

        For i = 0 To match_coll.Count - 1
            arr_Module(i).ModuleName = match_coll(i).Value     '名称
            arr_Module(i).StartAddress = match_coll(i).FirstIndex + l_double_byte '字节开始的地方，考虑双字节情况
            this_double_byte = double_byte(arr_Module(i).ModuleName)  '当前字符的双字节字符个数
            l_double_byte = l_double_byte + this_double_byte
            arr_Module(i).Size = arr_Module(i).ModuleName.Length + this_double_byte  '占用的字节个数

            '隐藏模块的情况，包含了前后2个ODOA的位置
            If arr_Module(i).ModuleName.Substring(0, 2) = Chr(&HD) & Chr(&HA) Then
                arr_Module(i).Size = arr_Module(i).Size - 4
                arr_Module(i).ModuleName = "(隐藏的)字节长度(含Module=)=" & arr_Module(i).Size.ToString
                arr_Module(i).StartAddress = arr_Module(i).StartAddress + 2
            End If

            ReDim arr_Module(i).ArrAddress(arr_Module(i).Size - 1) '
        Next i
        '修正地址，因为有可能是不连续的，理论上1个模块还可能可能跨越2个sector
        '直接计算到每一个字符的地址
        Dim p_address As Integer                       '处在哪个档次的下标上
        Dim byte_index As Integer
        For j = 0 To i - 1
            byte_index = arr_Module(j).StartAddress

            For k = 0 To arr_Module(j).Size - 1
                p_address = (k + byte_index) \ step_address
                arr_Module(j).ArrAddress(k) = arr_address(p_address, 1) + ((byte_index + k) Mod step_address)
            Next k
        Next j

        re = Nothing
        match_coll = Nothing

        Return i - 1
    End Function
    '根据ModuleName，将找到的模块在PROJECT中的byte修改为0D0A
    Function HideModule(ModuleName As String) As Integer
        Dim arr_byte(0) As Byte

        For i As Integer = 0 To Me.arr_Module.Length - 1
            If arr_Module(i).ModuleName = ModuleName Then
                Dim fw As FileStream = New FileStream(Me.path, FileMode.Open)
                For j As Integer = 0 To arr_Module(i).Size - 1
                    If j Mod 2 = 0 Then
                        arr_byte(0) = CByte(&HD)
                        fw.Seek(arr_Module(i).ArrAddress(j), origin:=0)
                        fw.Write(arr_byte, 0, arr_byte.Length)
                    Else
                        arr_byte(0) = CByte(&HA)
                        fw.Seek(arr_Module(i).ArrAddress(j), origin:=0)
                        fw.Write(arr_byte, 0, arr_byte.Length)
                    End If
                    file_byte(arr_Module(i).ArrAddress(j)) = arr_byte(0)
                Next
                fw.Close()
                MsgBox("OK")
                Return 1
            End If
        Next

        Return 0
    End Function

    Function UnHideModule(index_arr_Module As Integer, moduleName As String) As Integer
        Dim arr_byte() As Byte = System.Text.Encoding.Default.GetBytes(moduleName)
        Dim write_len As Integer = arr_Module(index_arr_Module).Size
        If arr_byte.Length < write_len Then
            write_len = arr_byte.Length
        End If

        Dim arr_byte_input(0) As Byte
        Dim fw As FileStream = New FileStream(Me.path, FileMode.Open)
        For i As Integer = 0 To write_len - 1
            arr_byte_input(0) = arr_byte(i)
            fw.Seek(arr_Module(index_arr_Module).ArrAddress(i), origin:=0)
            fw.Write(arr_byte_input, 0, arr_byte_input.Length)

            file_byte(arr_Module(index_arr_Module).ArrAddress(i)) = arr_byte_input(0)
        Next
        fw.Close()

        Return 0
    End Function

    Function write_file_byte(arr() As Object, l_index As Integer, Optional str_un_hide As String = "")
        '    Dim num_file As Integer
        '    Dim i As Integer, k As Integer
        '    Dim arr_un_hide() As Byte
        '    Dim b_input As Byte

        '    If VBA.Len(str_un_hide) > 0 Then
        '        arr_un_hide = VBA.StrConv(str_un_hide, vbFromUnicode)
        '    End If

        '    num_file = FreeFile()
        '    Open Me.path For Binary Access Write As #num_file

        'k = 3
        '    Do Until arr(l_index, k) = 0
        '        If VBA.Len(str_un_hide) > 0 Then
        '            If k - 3 > UBound(arr_un_hide) Then Exit Do
        '            b_input = arr_un_hide(k - 3)
        '        ElseIf k Mod 2 = 1 Then
        '            b_input = VBA.CByte(&HD)
        '        Else
        '            b_input = VBA.CByte(&HA)
        '        End If

        '        Put #num_file, arr(l_index, k) + 1, b_input

        '    k = k + 1
        '    Loop
        '    Close num_file

        Return 0
    End Function

    Private Function GetFileByte() As Integer
        MFunc.read_file_to_byte(Me.path, file_byte)
        FileAddress = GCHandle.Alloc(file_byte, GCHandleType.Pinned).AddrOfPinnedObject()
        Return 1
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()

        Erase arr_MSAT
        Erase arr_SAT
        Erase arr_SSAT
        Erase file_byte
    End Sub
End Class

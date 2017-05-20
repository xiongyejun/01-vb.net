Imports System.IO
Imports System.Runtime.InteropServices

Public Class CXlsFile
    Inherits CCompdocFile

    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal Destination As IntPtr, ByVal Source As IntPtr, ByVal Length As Integer)

    Sub New(file_name As String)
        MyBase.New(file_name)

    End Sub

    Overrides Function GetFileByte() As Integer
        MFunc.read_file_to_byte(Me.path, file_byte)
        FileAddress = GCHandle.Alloc(file_byte, GCHandleType.Pinned).AddrOfPinnedObject()

        Return 1
    End Function

    Overrides Sub ReWriteFile(ByRef arr_byte_to_write() As Byte, arr_address(,) As Integer, step_address As Integer, n_size As Integer, arr_dir_address_j As Integer, ByRef stream_size As Integer)

        Dim tmp_byte(step_address - 1) As Byte '临时写入文件用的，中间桥梁
        Dim p1 As IntPtr = GCHandle.Alloc(arr_byte_to_write, GCHandleType.Pinned).AddrOfPinnedObject()
        Dim p2 As IntPtr = GCHandle.Alloc(tmp_byte, GCHandleType.Pinned).AddrOfPinnedObject()

        Dim fw As FileStream = New FileStream(Me.path, FileMode.Open)
        'If arr_byte_to_write.Length \ step_address > n_size Then n_size = arr_byte_to_write.Length \ step_address
        For i_address As Integer = 0 To n_size
            CopyMemory(p2, p1 + i_address * step_address, step_address)
            CopyMemory(FileAddress + arr_address(i_address, 1), p1 + i_address * step_address, step_address)
            fw.Seek(arr_address(i_address, 1), origin:=0)
            fw.Write(tmp_byte, 0, tmp_byte.Length)
        Next

        '重设PROJECT目录的长度

        stream_size = arr_byte_to_write.Length
        Dim tmp_i As Integer = arr_byte_to_write.Length
        Dim tmp_i_to_byte(3) As Byte
        p1 = GCHandle.Alloc(tmp_i, GCHandleType.Pinned).AddrOfPinnedObject()
        p2 = GCHandle.Alloc(tmp_i_to_byte, GCHandleType.Pinned).AddrOfPinnedObject()
        CopyMemory(p2, p1, 4)

        CopyMemory(arr_dir_address_j + DIR_SIZE - 4 + FileAddress, p1, 4）

        fw.Seek(arr_dir_address_j + DIR_SIZE - 4 * 2, origin:=0) '128dir的长度，stream_size是倒数第2个  
        fw.Write(tmp_i_to_byte, 0, tmp_i_to_byte.Length)


        fw.Close()
    End Sub

End Class

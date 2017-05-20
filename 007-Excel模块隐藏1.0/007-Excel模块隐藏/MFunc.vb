Imports System.Text
Imports System.IO

Module MFunc
    '-1     不知道什么文件
    '0      xls文件
    '1      zip文件
    Function IsCompdocFile(file_name As String) As Integer
        Dim head_byte() As Byte = {&HD0, &HCF, &H11, &HE0, &HA1, &HB1， &H1A, &HE1}
        Dim file_byte(head_byte.Length - 1) As Byte
        If read_file_to_byte(file_name, file_byte) = 0 Then Return -1

        For i As Integer = 0 To head_byte.Length - 1
            If head_byte(i) <> file_byte(i) Then
                If file_byte(0) = &H50 AndAlso file_byte(1) = &H4B Then
                    Return 1
                Else
                    Return False
                End If
            End If
        Next

        Return 0
    End Function

    Function write_byte_to_file(file_name As String, arr_byte() As Byte, start_address As Long) As Integer
        Dim fw As FileStream = New FileStream(file_name, FileMode.Open)
        Try
            fw.Seek(start_address, origin:=0)
            fw.Write(arr_byte, 0, arr_byte.Length)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            fw.Close()
        End Try
        Return 1
    End Function

    Function read_file_to_byte(ByVal file_name As String, ByRef arr_byte() As Byte) As Integer
        Try
            Dim fs As FileStream = New FileStream(file_name, FileMode.Open)
            Dim file_info As FileInfo = New FileInfo(file_name)
            Dim len As Long = file_info.Length
            If len = 0 Then
                MsgBox("空文件")
                Return -1
            End If

            ReDim arr_byte(len - 1)
            fs.Read(arr_byte, 0, len)
            fs.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            Return 0
        End Try

        Return 1
    End Function

    Function my_hex(num As Integer) As String

        my_hex = Hex(num)

        If my_hex.Length = 1 Then
            my_hex = "0" & my_hex
        End If

        my_hex = "&H" & my_hex
    End Function

    Function double_byte(str As String) As Integer
        Dim i As Integer
        Dim i_len As Integer

        For i = 1 To str.Length
            If Asc(Mid(str, i, 1)) < 0 Then
                i_len += 1
            End If
        Next i
        Return i_len
    End Function

End Module

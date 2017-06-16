Imports System.Text
Imports System.IO

Module MFunc
    ''' <summary>
    ''' 读取文件到string
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <returns></returns>
    Function ReadFileToString(FileName As String) As String
        Dim sr As StreamReader = New StreamReader(FileName)
        Dim str As String = ""

        Try
            str = sr.ReadToEnd()

        Catch ex As Exception

        Finally
            sr.Close()
        End Try

        Return str
    End Function

    Function write_byte_to_file(file_name As String, arr_byte() As Byte, start_address As Long) As Long
        Dim fw As FileStream = New FileStream(file_name, FileMode.Create)
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

    Function read_file_to_byte(ByVal file_name As String, ByRef arr_byte() As Byte) As Long
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
End Module

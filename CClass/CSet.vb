Imports System.IO

''' <summary>
''' 保存或读取一些设置
''' </summary>
''' <remarks></remarks>
Public Class CSet
    Private File_Name As String
    Const SPLIT_WORD As String = "※" '分割符号

    Sub New(ByVal FileName As String)
        File_Name = FileName
    End Sub

    Function Read() As Dictionary(Of String, String)
        Dim dic As New Dictionary(Of String, String)
        If Not File.Exists(File_Name) Then Return dic

        Dim sr As StreamReader

        Try
            sr = New StreamReader(File_Name)
            Do Until sr.EndOfStream()
                Dim str As String = sr.ReadLine()
                Dim tmp() As String = Split(str, SPLIT_WORD)
                dic(tmp(0)) = tmp(1)
            Loop
            sr.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return dic
    End Function

    Function Write(ByVal dic As Dictionary(Of String, String)) As Integer
        Dim sw As StreamWriter = New StreamWriter(File_Name, False)
        For i As Integer = 0 To dic.Count - 1
            sw.WriteLine(dic.Keys(i) & SPLIT_WORD & dic.Values(i))
        Next
        sw.Close()

        Return 1
    End Function

End Class

Imports System.IO

Module Func
    Public Const SPE_CHAR As String = " ㊣ ㊣,㊣，㊣.㊣。㊣;㊣；/㊣、㊣*㊣|㊣、㊣+㊣(㊣（㊣)㊣）㊣'㊣①㊣②㊣③㊣④㊣⑤㊣⑥㊣⑦㊣⑧㊣⑨㊣⑩㊣"

    Sub TxtWrite(ByVal fileName As String, ByVal strWrite As String)
        Dim tw As StreamWriter = New StreamWriter(fileName)
        tw.Write(strWrite)
        tw.Close()
    End Sub

    Function TxtRead(ByVal fileName As String) As String
        Dim tr As StreamReader = New StreamReader(fileName)
        Dim str = tr.ReadToEnd()
        tr.Close()
        Return str
    End Function

    Function ExcelData(ByVal fileName As String) As String '读取Excel数据库
        ExcelData = "Provider =Microsoft.jet.OLEDB.4.0;data source=" & fileName & ";Extended properties=""Excel 8.0;HDR=YES"";"
    End Function
    Function AccessData(ByVal fileName As String) As String '读取Aceess数据库
        AccessData = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fileName

    End Function
    'Microsoft ActiveX Data Objects 2.8 Library
    Function CreateAdoArr(ByVal SqlStr As String, ByVal fileName As String, ByRef Arr(,) As System.Object) As Long '0
        '表示出错， 1表示正确
        Dim AdoConn As Object = Nothing, rst As Object
        Try
            Erase Arr

            AdoConn = CreateObject("ADODB.Connection")
            rst = CreateObject("ADODB.Recordset")

            AdoConn.Open(AccessData(fileName))
            rst.Open(SqlStr, AdoConn)

            Dim i As Integer = 0
            Do Until rst.eof
                ReDim Preserve Arr(rst.fields.count - 1, i)

                For j As Integer = 0 To rst.fields.count - 1
                    Arr(j, i) = rst.fields(j).value
                Next
                i += 1
                rst.movenext
            Loop

            CreateAdoArr = 1
            rst.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            CreateAdoArr = 0
        Finally
            AdoConn.Close()
            rst = Nothing
            AdoConn = Nothing
        End Try
    End Function

    Function clearChar(ByVal str As String) As String
        Dim arr() As String = Split(SPE_CHAR, "㊣")
        For i As Integer = 0 To arr.Length - 1
            str = Replace(str, arr(i), "")
        Next

        Return str
    End Function
End Module

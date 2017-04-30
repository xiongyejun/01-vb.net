Imports System.Text

Module MFunc
    Private Declare Auto Function GetShortPathName Lib "kernel32" (
        ByVal longPath As String,
        ByVal shortPath As StringBuilder,
        ByVal shortBufferSize As Int32) As Int32

    '取得短路徑及檔名
    Function ConvertToShortPathName(fileName As String) As String
        Dim shortName As New StringBuilder(256)
        GetShortPathName(fileName, shortName, shortName.Capacity)
        Return shortName.ToString
    End Function

    Function ByteToFile(ByRef arrByte() As Byte, strFileName As String) As String

        With CreateObject("Adodb.Stream")
            .Type = 1 'adTypeBinary
            .Open
            .Write(arrByte)
            .SaveToFile(strFileName, 2) 'adSaveCreateOverWrite
            .Close
        End With

        Return strFileName
    End Function

    Function Json(str_html As String) As String
        Dim objJSON As Object
        Dim Cell '这里不能定义为object类型
        Dim tmp
        Dim str As String = ""

        With CreateObject("msscriptcontrol.scriptcontrol")
            .Language = "JavaScript"
            .AddCode("var mydata =" & str_html)
            objJSON = .CodeObject
        End With
        '    Stop '查看vba本地窗口里objJSON对象以了解JSON数据在vba里的形态
        For Each Cell In objJSON.mydata.translateResult
            For Each tmp In Cell
                str = str & tmp.tgt
            Next tmp
        Next

        Return str
    End Function
End Module

Imports System.Net
Imports System.IO
Imports System.Text

Module Module1

    Function Main() As Integer
        Console.Title = "Translate"
        For Each argument In My.Application.CommandLineArgs
            Console.WriteLine(argument.ToString)
        Next

        WriteTiShi()

        Dim StrReadLine As String = Console.ReadLine()

        Do Until StrReadLine = "exit"
            Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine(Translate(StrReadLine))
            Console.WriteLine(GetJsonData(Translate(StrReadLine)))
            Console.ResetColor()

            WriteTiShi()
            StrReadLine = Console.ReadLine()
        Loop

        Return 0
    End Function

    ''' <summary>
    ''' 网抓
    ''' </summary>
    ''' <param name="Str"></param>
    ''' <returns></returns>
    Function Translate(Str As String) As String
        Dim str_html As String = ""
        Dim httpReq As HttpWebRequest = Nothing
        Dim httpResp As HttpWebResponse = Nothing
        Dim sr As StreamReader = Nothing

        Const strURL As String = "http://fanyi.youdao.com/translate"

        Dim httpURL As New System.Uri(strURL)
        Dim str_post As String = "i=" & Str & "&doctype=json"
        Dim a() As Byte = Encoding.Default.GetBytes(str_post)

        Try
            httpReq = CType(WebRequest.Create(httpURL), HttpWebRequest)
            httpReq.Method = "POST"
            httpReq.ContentLength = a.Length
            httpReq.ContentType = "application/x-www-form-urlencoded"
            Dim s As Stream = httpReq.GetRequestStream()
            s.Write(a, 0, a.Length)

            httpResp = CType(httpReq.GetResponse(), HttpWebResponse)
            sr = New StreamReader(httpResp.GetResponseStream, Encoding.UTF8)
            str_html = sr.ReadToEnd()
            sr.Close()
            s.Close()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally
            httpResp = Nothing
            httpReq.Abort()
            httpReq = Nothing

        End Try

        Return str_html

    End Function

    Function GetJsonData(StrJson As String) As String
        Dim objJSON As Object
        Dim Cell '这里不能定义为object类型
        Dim tmp
        Dim str As String = ""

        With CreateObject("msscriptcontrol.scriptcontrol")
            .Language = "JavaScript"
            .AddCode("var mydata =" & StrJson)
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

    ''' <summary>
    ''' 提示信息
    ''' </summary>
    Sub WriteTiShi()
        Console.ForegroundColor = ConsoleColor.Green
        Console.WriteLine()
        Console.WriteLine("exit:退出      其他:翻译")
        Console.Write("Translate$ ")
        Console.ResetColor()
    End Sub

End Module

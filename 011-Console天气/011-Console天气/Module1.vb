Imports System.Net
Imports System.Text
Imports System.IO

Module Module1

    Sub Main()
        Dim str_html As String = GetHtml()
        Dim str As String = htmlfile(str_html)
        str = Split(str, "分时段预报")(0)
        str = str.Trim()
        str = Replace(str, " ", "")
        str = Replace(str, vbCrLf & vbCrLf, vbCrLf)

        Console.BackgroundColor = ConsoleColor.DarkMagenta
        Console.WriteLine(str)
        Console.ResetColor()

        'Console.Read()
    End Sub



    ''' <summary>
    ''' 网抓
    ''' </summary>
    ''' <returns></returns>
    Function GetHtml() As String
        Dim str_html As String = ""
        Dim httpReq As HttpWebRequest = Nothing
        Dim httpResp As HttpWebResponse = Nothing
        Dim sr As StreamReader = Nothing

        Const strURL As String = "http://www.weather.com.cn/weather/101240101.shtml"

        Dim httpURL As New System.Uri(strURL)
        'Dim str_post As String = "i=" & Str & "&doctype=json"
        'Dim a() As Byte = Encoding.ASCII.GetBytes(str_post)

        Try
            httpReq = CType(WebRequest.Create(httpURL), HttpWebRequest)
            httpReq.Method = "GET"
            'httpReq.ContentLength = a.Length
            httpReq.ContentType = "application/x-www-form-urlencoded"
            'Dim s As Stream = httpReq.GetRequestStream()
            's.Write(a, 0, a.Length)

            httpResp = CType(httpReq.GetResponse(), HttpWebResponse)
            sr = New StreamReader(httpResp.GetResponseStream, Encoding.UTF8)
            str_html = sr.ReadToEnd()
            sr.Close()
            's.Close()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally
            httpResp = Nothing
            httpReq.Abort()
            httpReq = Nothing

        End Try

        Return str_html

    End Function

    Function htmlfile(strHtml As String) As String
        Dim HTML As Object
        'Dim post_list As Object
        Dim el As Object

        HTML = CreateObject("htmlfile")


        HTML.write(strHtml) ' 写入数据
        el = HTML.getElementById("7d")

        Return el.innerText
        'For Each el In post_list.Children
        '    Debug.Print el.getElementsByTagName("a")(0).innerText
        'Next
    End Function
End Module

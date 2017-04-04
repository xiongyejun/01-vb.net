Imports System.Net
Imports System.IO

Module Func

    Sub btnAdd(ByVal form As Object, ByVal btn As System.Windows.Forms.Button, ByVal btnText As String, Optional ByVal iLeft As Integer = 5, Optional ByVal iTop As Integer = 0, Optional ByVal iWidth As Integer = 100, Optional ByVal iHeight As Integer = 30)
        With btn
            .Left = iLeft
            .Top = iTop
            .Width = iWidth
            .Height = iHeight
            .Text = btnText
        End With
        form.Controls.Add(btn)
    End Sub

    Sub comboBoxAdd(ByVal form As Object, ByVal cb As System.Windows.Forms.ComboBox, ByVal text As String, Optional ByVal iLeft As Integer = 5, Optional ByVal iTop As Integer = 0, Optional ByVal iWidth As Integer = 100, Optional ByVal iHeight As Integer = 30)
        With cb
            .Left = iLeft
            .Top = iTop
            .Width = iWidth
            .Height = iHeight
            .Text = text
        End With
        form.Controls.Add(cb)
    End Sub
    Sub groupBoxAdd(ByVal form As Object, ByVal gp As System.Windows.Forms.GroupBox, ByVal text As String, Optional ByVal iLeft As Integer = 5, Optional ByVal iTop As Integer = 0, Optional ByVal iWidth As Integer = 100, Optional ByVal iHeight As Integer = 30)
        With gp
            .Left = iLeft
            .Top = iTop + 5
            .Width = iWidth
            .Height = iHeight
            .Text = text
        End With
        form.Controls.Add(gp)
    End Sub

    Sub labelAdd(ByVal form As Object, ByVal label As System.Windows.Forms.Label, ByVal labelText As String, Optional ByVal iLeft As Integer = 5, Optional ByVal iTop As Integer = 0, Optional ByVal iWidth As Integer = 100, Optional ByVal iHeight As Integer = 30)
        With label
            .Left = iLeft
            .Top = iTop + 5
            .Width = iWidth
            .Height = iHeight
            .Text = labelText
            .AutoSize = True
        End With
        form.Controls.Add(label)
    End Sub

    Sub textBoxAdd(ByVal form As Object, ByVal textBox As System.Windows.Forms.TextBox, ByVal textBoxText As String, Optional ByVal iLeft As Integer = 5, Optional ByVal iTop As Integer = 0, Optional ByVal iWidth As Integer = 100, Optional ByVal iHeight As Integer = 30)
        With textBox
            .Left = iLeft
            .Top = iTop
            .Width = iWidth
            .Height = iHeight
            .Text = textBoxText
            .AutoSize = True
        End With
        form.Controls.Add(textBox)
    End Sub

    Sub richTextBoxAdd(ByVal form As Object, ByVal rtb As System.Windows.Forms.RichTextBox, ByVal text As String, Optional ByVal iLeft As Integer = 5, Optional ByVal iTop As Integer = 0, Optional ByVal iWidth As Integer = 100, Optional ByVal iHeight As Integer = 30)
        With rtb
            .Left = iLeft
            .Top = iTop
            .Width = iWidth
            .Height = iHeight
            .Text = text
        End With
        form.Controls.Add(rtb)
    End Sub

    Function readHtml(ByVal Url As String) As String
        Dim httpReq As System.Net.HttpWebRequest
        Dim httpResp As System.Net.HttpWebResponse
        Dim httpURL As New System.Uri(Url)
        httpReq = CType(WebRequest.Create(httpURL), HttpWebRequest)
        httpReq.Method = "GET"
        httpResp = CType(httpReq.GetResponse(), HttpWebResponse)
        httpReq.KeepAlive = False ' 获取或设置一个值，该值指示是否与 Internet资源建立持久连接。
        Dim reader As StreamReader = New StreamReader(httpResp.GetResponseStream, System.Text.Encoding.GetEncoding(-0))
        Dim respHTML As String = reader.ReadToEnd()

        Return respHTML
    End Function

    Function downFile(ByVal file As String, ByVal saveFile As String)
        Dim DownApp As New Microsoft.VisualBasic.Devices.Network ' System.Net.WebClient 'Microsoft.VisualBasic.Devices.Network
        DownApp.DownloadFile(file, saveFile, "", "", False, 5000, True)
        Return 0
    End Function

    Function ReadText(ByVal file As String) As String
        Dim sr As StreamReader = New StreamReader(file)
        Dim str As String = sr.ReadToEnd
        sr.Close()
        Return str
    End Function

    Sub WriteText(ByVal file As String, ByVal str As String)
        Dim sw As StreamWriter = New StreamWriter(file)
        sw.Write(str)
        sw.Close()
    End Sub

End Module

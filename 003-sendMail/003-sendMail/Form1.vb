Public Class Form1
    Private WithEvents btnSend As System.Windows.Forms.Button


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim iLeft As Integer = 5
        Dim iTop As Integer = 5

        btnSend = New Button
        Func.btnAdd(Me, btnSend, "Send", iLeft, iTop)


    End Sub

    Private Sub btnSend_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSend.Click
        'Dim receiveAddressList As List(Of String) = New List(Of String)
        'receiveAddressList.Add("244114746@qq.com")

        'SendMail(receiveAddressList, "test主题", "test内容")
        SendMail()
    End Sub

    ''' <summary>  
    ''' 通过SmtpClient类发送电子邮件  
    ''' </summary>  
    ''' <param name="ReceiveAddressList">收件人地址列表</param>  
    ''' <param name="Subject">邮件主题</param>  
    ''' <param name="Content">邮件内容</param>  
    ''' <param name="AttachFile">附件列表Hastable。KEY=文件名,Value文件路径</param>  
    Private Function SendMail(ByVal ReceiveAddressList As List(Of String), ByVal Subject As String, ByVal Content As String, _
                              Optional ByVal AttachFile As Hashtable = Nothing) As Boolean
        Dim i As Integer
        'SMTP客户端  
        Dim smtp As New System.Net.Mail.SmtpClient("SMTP.qq.com", 25)
        'smtp.Host = "smtp.163.com"       'SMTP服务器名称  
        '发件人邮箱身份验证凭证。 参数分别为 发件邮箱登录名和密码  
        smtp.Credentials = New System.Net.NetworkCredential("244114746@qq.com", "xiongye001")
        
        'smtp.DeliveryMethod = Net.Mail.SmtpDeliveryMethod.Network
        'smtp.UseDefaultCredentials = False ' 表示以当前登录用户的默认凭据进行身份验证

        '创建邮件  
        Dim mail As New System.Net.Mail.MailMessage()
        '主题编码  
        mail.SubjectEncoding = System.Text.Encoding.GetEncoding("GB2312")
        '正文编码  
        mail.BodyEncoding = System.Text.Encoding.GetEncoding("GB2312")
        '邮件优先级  
        mail.Priority = System.Net.Mail.MailPriority.Normal
        '以HTML格式发送邮件,为false则发送纯文本邮箱  
        mail.IsBodyHtml = True
        '发件人邮箱  
        mail.From = New System.Net.Mail.MailAddress("244114746@qq.com.cn")

        '添加收件人,如果有多个,可以多次添加  
        If ReceiveAddressList.Count = 0 Then Return False
        For i = 0 To ReceiveAddressList.Count - 1
            mail.To.Add(ReceiveAddressList.Item(i))
        Next

        '邮件主题和内容  
        mail.Subject = Subject
        mail.Body = Content

        '定义附件,参数为附件文件名,包含路径,推荐使用绝对路径  
        If Not AttachFile Is Nothing AndAlso AttachFile.Count <> 0 Then
            For Each sKey As String In AttachFile.Keys
                Dim objFile As New System.Net.Mail.Attachment(AttachFile.Item(sKey))
                '附件文件名,用于收件人收到附件时显示的名称  
                objFile.Name = sKey
                '加入附件,可以多次添加  
                mail.Attachments.Add(objFile)
            Next
        End If

        '发送邮件  
        
        Try
            smtp.Send(mail)
            MessageBox.Show("邮件发送成功！")
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            'MessageBox.Show("邮件发送失败！")
            Return False
        Finally
            mail.Dispose()
        End Try
    End Function

    Function sendMail()
        Dim outObj As Object
        Dim item As Object

        outObj = CreateObject("OutLook.Application")
        item = outObj.CreateItem(0)
        '设定收件人地址
        item.To = "244114746@qq.com"
        '设定邮件主题
        item.Subject = "test"
        '设定邮件内容
        item.Body = "dasfdasfsda"
        '设定添附文件
        'item.Attachments.Add("c:\tree.txt")
        '发送
        item.Send()

        Return 0
    End Function
End Class

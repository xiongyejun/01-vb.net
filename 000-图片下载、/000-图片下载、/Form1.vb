Imports System.Threading
Imports System.IO
Imports System.Text

Public Class Form1
    Private gb As System.Windows.Forms.GroupBox
    Private lbKey As System.Windows.Forms.Label
    Private WithEvents cbKey As System.Windows.Forms.ComboBox
    Private WithEvents btnDeleteKey As System.Windows.Forms.Button

    Private lbPages As System.Windows.Forms.Label
    Private WithEvents numPage As System.Windows.Forms.NumericUpDown

    Private lbSave As System.Windows.Forms.Label
    Private WithEvents tbSavePath As System.Windows.Forms.TextBox
    Private WithEvents btnSavePath As System.Windows.Forms.Button
    Private WithEvents fbdSavePath As System.Windows.Forms.FolderBrowserDialog
    Private WithEvents pb As System.Windows.Forms.ProgressBar

    Private WithEvents btnStart As System.Windows.Forms.Button
    Private WithEvents btnEnd As System.Windows.Forms.Button

    Private WithEvents btnOpenStart As System.Windows.Forms.Button

    Private WithEvents rtb As System.Windows.Forms.RichTextBox
    Private WithEvents pic As System.Windows.Forms.PictureBox

    Dim t As Thread
    Dim txtFile As String = Application.StartupPath & "\data.txt"
    Dim cbKeyDic As New Dictionary(Of String, String)
    Dim picArr() As String, iPic As Integer = 0

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ileft As Integer = 5
        Dim iTop As Integer = 0
        Dim iBtnWidth As Integer = 30
        Dim iWidth As Integer = Screen.PrimaryScreen.Bounds.Width

        gb = New GroupBox
        Func.groupBoxAdd(Me, gb, "", ileft, iTop, iWidth - 10)

        iTop = 5
        lbKey = New Label
        Func.labelAdd(gb, lbKey, "关键字", ileft, iTop, 50)

        ileft += lbKey.Width
        cbKey = New ComboBox
        Func.comboBoxAdd(gb, cbKey, "美女", ileft, 7, 200)

        ileft += cbKey.Width
        btnDeleteKey = New Button
        Func.btnAdd(gb, btnDeleteKey, "deleteKey", ileft, iTop, 100, 25)

        ileft += btnDeleteKey.Width
        lbPages = New Label
        Func.labelAdd(gb, lbPages, "下载多少页", ileft, iTop, 80)

        ileft += lbPages.Width
        numPage = New NumericUpDown
        With numPage
            .Top = iTop + 2
            .Left = ileft
            .Value = 1
        End With
        gb.Controls.Add(numPage)

        iTop += cbKey.Height
        iTop += 5
        ileft = 5
        lbSave = New Label
        Func.labelAdd(gb, lbSave, "保存路径", ileft, iTop, 60)

        ileft += lbSave.Width
        tbSavePath = New TextBox
        Func.textBoxAdd(gb, tbSavePath, Application.StartupPath & "\Pic\", ileft, iTop, 500)

        ileft = 5
        iTop += tbSavePath.Height
        iTop += 5

        pb = New ProgressBar
        With pb
            .Left = ileft
            .Width = lbSave.Width + tbSavePath.Width
            .Top = iTop
        End With
        gb.Controls.Add(pb)

        iTop += pb.Height
        iTop += 5

        btnStart = New Button
        Func.btnAdd(gb, btnStart, "开始", ileft, iTop)
        ileft += btnStart.Width
        btnEnd = New Button
        Func.btnAdd(gb, btnEnd, "结束", ileft, iTop)
        ileft += btnEnd.Width
        btnOpenStart = New Button
        Func.btnAdd(gb, btnOpenStart, "打开文件夹", ileft, iTop)

        ileft = 5
        iTop += btnOpenStart.Height

        rtb = New RichTextBox
        Func.richTextBoxAdd(gb, rtb, "", lbSave.Width + tbSavePath.Width + 10, 0, iWidth - lbSave.Width - tbSavePath.Width - 20, iTop)
        gb.Height = iTop + 5

        ileft = 5
        iTop += 10

        pic = New PictureBox
        With pic
            .Left = ileft
            .Top = iTop
            .Width = iWidth - 10
            .Height = Screen.PrimaryScreen.Bounds.Height - iTop - 60
            .SizeMode = PictureBoxSizeMode.Zoom
            .Enabled = False
            .BorderStyle = BorderStyle.FixedSingle
        End With

        With Me.Controls
            .Add(pic)
        End With

        With Me
            .Width = iWidth
            .Height = Screen.PrimaryScreen.Bounds.Height - 20
            .Text = "文件下载"
            .Location = New Point(0, 0)
            .WindowState = 2
        End With
        Me.readData()

        Control.CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Function readData()
        If File.Exists(txtFile) Then
            Dim str As String = Func.ReadText(txtFile)
            Dim arr() As String = Split(str, vbNewLine)

            Me.numPage.Value = arr(0)
            Me.tbSavePath.Text = arr(1)

            For i As Integer = 2 To arr.Length - 1
                Me.cbKey.Items.Add(arr(i))
                cbKeyDic(arr(i)) = i.ToString
            Next
        End If
        Return 0
    End Function

    Private Function writeData()
        Dim str As String = Me.numPage.Value.ToString & vbNewLine
        str &= Me.tbSavePath.Text
        str &= vbNewLine

        For i As Integer = 0 To Me.cbKey.Items.Count - 1
            If Me.cbKey.Items(i).ToString <> "" Then
                str &= Me.cbKey.Items(i).ToString
                str &= vbNewLine
            End If
        Next

        Func.WriteText(txtFile, str)
        Return 0
    End Function

    Private Sub btnStart_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStart.Click
        If Not cbKeyDic.ContainsKey(Me.cbKey.Text) Then
            cbKeyDic(Me.cbKey.Text) = Me.cbKey.Text
            Me.cbKey.Items.Add(Me.cbKey.Text)
        End If

        t = New Thread(AddressOf savePic)
        t.Start()
    End Sub

    Sub savePic()
        Dim savePath As String = Me.tbSavePath.Text

        If Directory.Exists(savePath) Then
            Directory.Delete(savePath, True)
        End If
        Directory.CreateDirectory(savePath)
        Dim uriStr As String = "http://image.baidu.com/search/avatarjson?tn=resultjsonavatarnew&ie=utf-8&"
        uriStr = uriStr & "word=" & Uri.EscapeDataString(Me.cbKey.Text) & "&cg=girl&"

        Me.pb.Maximum = Me.numPage.Value * 60
        Me.pb.Value = 0
        Me.rtb.Text = ""

        pic.Enabled = False

        For i = 1 To Me.numPage.Value
            Dim str As String = ""
            str = uriStr & "pn=" & i * 60 & "&rn=60&itg=1&z=0&fr=&lm=-1&ic=0&s=0&st=-1&gsm=" & GenerationRandom(10)
            str = Func.readHtml(str)
            Dim arr() As String = Split(str, """objURL"":""")

            For j = 1 To arr.Length - 1
                Dim iLen As Integer = arr(j).IndexOf(".jpg")
                Dim subStr As String = arr(j).Substring(0, iLen + 4)
                Me.rtb.AppendText(subStr & vbNewLine)
                Dim saveFile As String = savePath & Format(i, "00-") & Format(j, "00") & ".jpg"

                Try
                    Func.downFile(subStr, saveFile)
                    Me.rtb.AppendText(saveFile & vbNewLine)
                    rtb.ScrollToCaret()
                    Me.pic.ImageLocation = saveFile
                Catch ex As Exception
                End Try

                Try
                    Dim f As FileInfo = New FileInfo(saveFile)
                    If f.Length Then
                        ReDim Preserve picArr(iPic)
                        picArr(iPic) = saveFile
                        iPic += 1
                    Else
                        File.Delete(saveFile)
                    End If
                Catch ex As Exception

                End Try

                Me.pb.Value += 1
            Next
        Next

        pic.Enabled = True
        iPic -= 1
        MsgBox("完成")
    End Sub

    Private Sub pic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pic.Click
        Dim x As Integer = Windows.Forms.Cursor.Position.X
        Dim maxX As Integer = Screen.PrimaryScreen.Bounds.Width
        If x < (maxX >> 2) Then
            iPic -= 1
            If iPic < 0 Then iPic = picArr.Length - 1
        Else
            iPic += 1
            If iPic >= picArr.Length Then iPic = 0
        End If
        Me.pic.ImageLocation = picArr(iPic)
        Me.Text = "图片下载-" & iPic
    End Sub

    Private Function GenerationRandom(ByVal length As Integer) As String

        Dim str() As String = New String() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"}

        Dim sb As StringBuilder = New StringBuilder()
        Dim random As Random = New Random()
        For i = 0 To length
            sb.Append(random.Next(36))
        Next

        Return sb.ToString()
    End Function

    Private Sub btnEnd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEnd.Click, Me.FormClosed
        Me.pic.Enabled = True
        Try
            t.Abort()
        Catch ex As Exception

        End Try
        Me.writeData()
    End Sub

    Private Sub btnOpenStart_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOpenStart.Click
        Process.Start(Application.StartupPath)
    End Sub

    Private Sub btnDeleteKey_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDeleteKey.Click
        Me.cbKey.Items.Remove(Me.cbKey.Text)
    End Sub
End Class

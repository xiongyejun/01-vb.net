

Public Class Form1

    Private WithEvents rtb_original As System.Windows.Forms.RichTextBox
    Private WithEvents rtb_result As System.Windows.Forms.RichTextBox

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i_left As Integer = 0
        Dim i_top As Integer = 0
        Const TB_WIDTH As Integer = 300
        Const TB_HEIGHT As Integer = 300

        rtb_original = New RichTextBox
        With rtb_original
            .Left = i_left
            .Top = i_top
            .Width = TB_WIDTH
            .Height = TB_HEIGHT

        End With

        i_left += rtb_original.Width
        rtb_result = New RichTextBox
        With rtb_result
            .Left = i_left
            .Top = i_top
            .Width = TB_WIDTH
            .Height = TB_HEIGHT
            .ReadOnly = True
        End With

        With Me.Controls
            .Add(rtb_original)
            .Add(rtb_result)
        End With

        Me.Width = 2 * TB_WIDTH + 15
        Me.Height = TB_HEIGHT + 40
        Me.TopMost = True
        Me.Text = "有道翻译——Me.TopMost=" & Me.TopMost.ToString
        '初始到左上角
        Me.Location = New Point(Screen.PrimaryScreen.Bounds.Width - Me.Width, 0)
    End Sub
    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Me.rtb_original.Width = (Me.Width - 15) / 2
        Me.rtb_original.Height = Me.Height - 40

        Me.rtb_result.Left = Me.rtb_original.Width
        Me.rtb_result.Width = Me.rtb_original.Width
        Me.rtb_result.Height = Me.rtb_original.Height
    End Sub
    Private Sub rtb_result_DoubleClick(sender As Object, e As EventArgs) Handles rtb_result.DoubleClick
        '下载MP3
        Dim str_word As String = Me.rtb_original.Text
        str_word = str_word.Replace(" ", "%20")
        Dim strFileName As String = Application.StartupPath & "\MP3\" & str_word & ".mp3"

        If Not System.IO.File.Exists(strFileName) Then
            strFileName = GetMp3(str_word, strFileName)
        End If

        If strFileName.Length > 0 Then
            PlayMidiFile(strFileName)
        End If

    End Sub

    Private Sub rtb_original_DoubleClick(sender As Object, e As EventArgs) Handles rtb_original.DoubleClick
        Me.TopMost = Not Me.TopMost
        Me.Text = "有道翻译——Me.TopMost=" & Me.TopMost.ToString
    End Sub
    Private Sub rtb_original_TextChanged(sender As Object, e As EventArgs) Handles rtb_original.TextChanged
        Dim str_word As String = Me.rtb_original.Text
        If str_word.Length = 0 Then
            Me.rtb_result.Text = ""
        Else
            Me.rtb_result.Text = translate(Me.rtb_original.Text)
        End If

    End Sub

    Function translate(str_word As String) As String
        Dim xml As Object = CreateObject("Microsoft.XMLHTTP")

        Try
            With xml
                .Open("POST", "http://fanyi.youdao.com/translate", False)
                .setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
                .send("i=" & str_word & "&doctype=json")
                Return json(.responsetext)
                '        translate = Split(translate, """tgt"":""")(1)
                '        translate = Split(translate, """}]]}")(0)
            End With
        Catch ex As Exception
            Return ex.ToString
        End Try

    End Function
    Function GetMp3(str_word As String, strFileName As String) As String
        Dim xml As Object = CreateObject("Microsoft.XMLHTTP")
        Dim arr_byte() As Byte = Nothing

        Try
            With xml
                .Open("GET", "http://dict.youdao.com/dictvoice?audio=" & str_word & "&type=2", False)
                .setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
                .send
                arr_byte = .responseBody

                Return ByteToFile(arr_byte, strFileName)
            End With

        Catch ex As Exception
            Return ""
        End Try


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
    Function json(str_html As String) As String
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


End Class

Imports System.ComponentModel
Imports System.Threading

Public Class Form1

    Private WithEvents rtb_original As System.Windows.Forms.RichTextBox
    Private WithEvents rtb_result As System.Windows.Forms.RichTextBox

    Private WithEvents cms As ContextMenuStrip
    Private WithEvents cms_cls As New System.Windows.Forms.ToolStripMenuItem
    Private WithEvents cms_playMP3 As New System.Windows.Forms.ToolStripMenuItem
    Private WithEvents cms_OpenMP3Dir As New System.Windows.Forms.ToolStripMenuItem
    Private WithEvents cms_TopMost As New System.Windows.Forms.ToolStripMenuItem

#Region "Form"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i_left As Integer = 0
        Dim i_top As Integer = 0
        Const TB_WIDTH As Integer = 200
        Const TB_HEIGHT As Integer = 120

        cms = New System.Windows.Forms.ContextMenuStrip
        cms_cls.Text = "Clear"
        cms_playMP3.Text = "PlayMP3"
        cms_OpenMP3Dir.Text = "Open MP3 Dir"
        cms_TopMost.Text = "TopMost"
        With cms.Items
            .Add(cms_playMP3)
            .Add(cms_cls)
            .Add(cms_OpenMP3Dir)
            .Add(cms_TopMost)
        End With

        rtb_original = New RichTextBox
        With rtb_original
            .Left = i_left
            .Top = i_top
            .Width = TB_WIDTH
            .Height = TB_HEIGHT
            .ForeColor = Color.Red
            .ContextMenuStrip = cms
        End With

        i_left += rtb_original.Width
        rtb_result = New RichTextBox
        With rtb_result
            .Left = i_left
            .Top = i_top
            .Width = TB_WIDTH
            .Height = TB_HEIGHT
            .ReadOnly = True
            .ContextMenuStrip = cms
            .ForeColor = Color.Red

        End With

        With Me.Controls
            .Add(rtb_original)
            .Add(rtb_result)
        End With

        Control.CheckForIllegalCrossThreadCalls = False '允许跨线程访问控件

        Me.Opacity = 0.8
        'Me.TransparencyKey = Me.rtb_original.BackColor
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

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Dim folder As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(Application.StartupPath & "\MP3"）
        Try
            For Each fileinfo As System.IO.FileInfo In folder.GetFiles
                fileinfo.Delete()
            Next
        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "RTB"
    Private Sub rtb_result_DoubleClick(sender As Object, e As EventArgs) Handles rtb_result.DoubleClick
        GetMp3()
    End Sub

    Private Sub rtb_original_DoubleClick(sender As Object, e As EventArgs) Handles rtb_original.DoubleClick
        Me.rtb_original.Text = ""
        Me.rtb_original.Paste()
    End Sub

    Private Sub rtb_original_TextChanged(sender As Object, e As EventArgs) Handles rtb_original.TextChanged
        Dim str_word As String = Trim(Me.rtb_original.Text)
        Me.rtb_original.Text = str_word

        If str_word.Length = 0 Then
            Me.rtb_result.Text = ""
        Else
            Dim t As Thread = New Thread(AddressOf Translate)
            t.Start()
        End If
    End Sub

    Sub Translate()
        Dim str_word As String = Me.rtb_original.Text

        Dim xml As Object = CreateObject("Microsoft.XMLHTTP")

        Try
            With xml
                .Open("POST", "http://fanyi.youdao.com/translate", False)
                .setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
                .send("i=" & str_word & "&doctype=json")
                Me.rtb_result.Text = Json(.responsetext)
                '        translate = Split(translate, """tgt"":""")(1)
                '        translate = Split(translate, """}]]}")(0)
            End With
        Catch ex As Exception
            Me.rtb_result.Text = ex.ToString
        End Try

    End Sub
#End Region

#Region "cms"
    Private Sub cms_cls_Click(sender As Object, e As EventArgs) Handles cms_cls.Click
        Me.rtb_original.Text = ""
    End Sub

    Private Sub cms_playMP3_Click(sender As Object, e As EventArgs) Handles cms_playMP3.Click
        '下载MP3
        'Dim t As Thread = New Thread(AddressOf GetMp3)
        't.Start()
        GetMp3() '不能用thread，不知道为什么！
    End Sub

    Private Sub cms_TopMost_Click(sender As Object, e As EventArgs) Handles cms_TopMost.Click
        Me.TopMost = Not Me.TopMost
        Me.Text = "有道翻译——Me.TopMost=" & Me.TopMost.ToString
    End Sub

    Sub GetMp3()
        Dim xml As Object = CreateObject("Microsoft.XMLHTTP")
        Dim arr_byte() As Byte = Nothing

        Dim str_word As String = Me.rtb_original.Text
        str_word = str_word.Replace(" ", "%20")
        Dim strFileName As String = str_word
        If strFileName.Length > 20 Then strFileName = strFileName.Substring(0, 20)
        strFileName = Application.StartupPath & "\MP3\" & strFileName & ".mp3"

        If Not System.IO.File.Exists(strFileName) Then
            Try
                With xml
                    .Open("GET", "http://dict.youdao.com/dictvoice?audio=" & str_word & "&type=2", False)
                    .setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
                    .send
                    arr_byte = .responseBody

                    strFileName = ByteToFile(arr_byte, strFileName)
                End With

            Catch ex As Exception
                strFileName = ""
            End Try
        End If

        If strFileName.Length > 0 Then
            PlayMidiFile(ConvertToShortPathName(strFileName))
            'Dim player As System.Media.SoundPlayer = New System.Media.SoundPlayer
            'player.SoundLocation = strFileName
            'player.Load()
            'player.Play()

        End If

    End Sub

    Private Sub cms_OpenMP3Dir_Click(sender As Object, e As EventArgs) Handles cms_OpenMP3Dir.Click
        System.Diagnostics.Process.Start(Application.StartupPath & "\MP3")
    End Sub





#End Region

End Class

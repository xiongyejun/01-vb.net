Public Class FCode

    Private WithEvents rtbCode As System.Windows.Forms.RichTextBox

    Public Sub New()

        ' 此调用是设计器所必需的。
        InitializeComponent()

        ' 在 InitializeComponent() 调用之后添加任何初始化。
        rtbCode = New RichTextBox
    End Sub

    WriteOnly Property RTBText() As String
        Set(value As String)
            rtbCode.Text = value
        End Set
    End Property


    Private Sub FCode_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i_left As Integer = 0
        Dim i_top As Integer = 0
        Const RTB_WIDTH As Integer = 500
        Const RTB_HEIGHT As Integer = 500


        With rtbCode
            .Top = i_top
            .Left = i_left
            .Width = RTB_WIDTH
            .Height = RTB_HEIGHT
        End With

        Me.Width = RTB_WIDTH + 10
        Me.Height = RTB_HEIGHT + 10

        Me.Controls.Add(rtbCode)
    End Sub

    Private Sub FCode_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Me.rtbCode.Width = Me.Width - 15
        Me.rtbCode.Height = Me.Height - 40
    End Sub

    Private Sub rtbCode_DoubleClick(sender As Object, e As EventArgs) Handles rtbCode.DoubleClick
        Dim str As String = rtbCode.Text
        Dim k As Integer = InStr(LCase(str), "_open")

        If k Then
            rtbCode.Select(k - 1, "_Open".Length)
        End If



    End Sub

End Class
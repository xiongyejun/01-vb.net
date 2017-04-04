

Public Class Form1
    Private WithEvents pic As System.Windows.Forms.PictureBox
    Private WithEvents cd As System.Windows.Forms.ColorDialog
    Private WithEvents btnPen As System.Windows.Forms.Button
    Private WithEvents btnSave As System.Windows.Forms.Button
    Private WithEvents btnClear As System.Windows.Forms.Button

    Dim newbit As System.Drawing.Bitmap
    Dim imagepen As Pen

    Dim x1 As Integer, y1 As Integer

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim iLeft As Integer = 5

        pic = New PictureBox
        With pic
            .Size = New Size(32, 32)
            .Location = New Point(10, 50)
            .BorderStyle = BorderStyle.FixedSingle
            .Image = Nothing
            .Cursor = System.Windows.Forms.Cursors.Cross
        End With

        cd = New ColorDialog

        btnPen = New Button
        Func.btnAdd(Me, btnPen, "选择画笔颜色", iLeft, 5)
        iLeft += btnPen.Width

        btnSave = New Button
        Func.btnAdd(Me, btnSave, "另存Ico", iLeft, 5)
        iLeft += btnSave.Width

        btnClear = New Button
        Func.btnAdd(Me, btnClear, "清除", iLeft, 5)

        Me.Controls.Add(pic)
        Me.Width = btnClear.Width * 3 + 15

        Dim bitnew As New System.Drawing.Bitmap(pic.Width, pic.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        For i As Integer = 0 To 31
            For j As Integer = 0 To 31
                bitnew.SetPixel(i, j, Color.Transparent)
            Next
        Next

        'Me.Icon = New Icon("C:\Documents and Settings\Administrator\桌面\xx.ico")
        newbit = bitnew
    End Sub

    Private Sub pic_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pic.MouseDown
        If e.Button = MouseButtons.Left Then
            x1 = e.X
            y1 = e.Y
        End If
    End Sub

    Private Sub pic_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pic.MouseMove
        Me.Text = Format(e.X, "X=0") & Format(e.Y, "  Y=0")
    End Sub

    Private Sub pic_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pic.MouseUp
        Me.k(Me.imagepen, x1, y1, e.X, e.Y)
    End Sub

    Public Sub k(ByVal drawtool As Object, ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer)

        pic.Image = newbit
        Dim Graphic As Graphics
        Graphic = Graphics.FromImage(Me.pic.Image) '在pic上画图 
        Graphic.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias '锯齿削边 
        'Graphic.DrawLine(drawtool, x1, y1, x2, y2) '画线 
        Graphic.DrawString("VB.net", New Font("宋体", 10), New SolidBrush(Color.Red), x1, y1)

    End Sub

    Private Sub btnPen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPen.Click
        Me.cd.ShowDialog()
        Dim pen As New Pen(cd.Color, 10)
        imagepen = pen
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim bmp As New System.Drawing.Bitmap(pic.Image, 32, 32)
        Dim ico As System.Drawing.Icon = Icon.FromHandle(bmp.GetHicon())
        Dim file As New System.IO.FileStream("C:\Documents and Settings\Administrator\桌面\xx.ico", IO.FileMode.Create)
        ico.Save(file)
        file.Close()
    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Dim Graphic As Graphics
        Graphic = Graphics.FromImage(Me.pic.Image) '在pic上画图 
        Graphic.Clear(Color.Transparent)
        pic.Refresh()
    End Sub
End Class

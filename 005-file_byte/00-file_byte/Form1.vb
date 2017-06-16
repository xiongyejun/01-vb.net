Public Class Form1
#Region "定义"
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents 菜单ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 选择文件ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

    Private WithEvents lv_hex As System.Windows.Forms.ListView
    Private WithEvents lb_file_path As System.Windows.Forms.Label

    '进制转换
    Private fr_trans As System.Windows.Forms.GroupBox
    Private lb_input_num As System.Windows.Forms.Label
    Private WithEvents tb_input_num As System.Windows.Forms.TextBox
    Private lb_show As System.Windows.Forms.Label

    '改写文件
    Private gb_write As System.Windows.Forms.GroupBox
    Private lb_wr As System.Windows.Forms.Label
    Private WithEvents tb_wr As System.Windows.Forms.TextBox
    Private lb_start As System.Windows.Forms.Label
    Private WithEvents tb_start As System.Windows.Forms.TextBox
    Private WithEvents btn_wr As System.Windows.Forms.Button

    Private WithEvents btn_read_hex_to_file As System.Windows.Forms.Button
#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i_left As Integer = 5
        Dim i_top As Integer = 5

        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.MenuStrip1.SuspendLayout()

        Me.菜单ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.选择文件ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        '
        'MenuStrip1
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.菜单ToolStripMenuItem})
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        '菜单ToolStripMenuItem
        Me.菜单ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.选择文件ToolStripMenuItem})
        Me.菜单ToolStripMenuItem.Text = "菜单"
        Me.菜单ToolStripMenuItem.Image = System.Drawing.SystemIcons.Information.ToBitmap
        '
        '选择文件ToolStripMenuItem        '
        Me.选择文件ToolStripMenuItem.Text = "选择文件"        '
        Me.选择文件ToolStripMenuItem.Image = System.Drawing.SystemIcons.Shield.ToBitmap

        i_top = 30
        lb_file_path = New Label
        With lb_file_path
            .Text = ""
            .Left = i_left
            .Top = i_top
            .AutoSize = True
        End With

        i_top += lb_file_path.Height
        lv_hex = New ListView
        With Me.lv_hex
            .Columns.Add("Address", 70, HorizontalAlignment.Right)
            .Columns.Add("00 01 02 03 04 05 06 07  08 09 0A 0B 0C 0D 0E 0F", 16 * 20, HorizontalAlignment.Left)
            .Columns.Add("asc", CInt(16 * 20 / 3) + 1)
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True
            .Sorting = SortOrder.None
            .Width = .Columns(0).Width + .Columns(1).Width + .Columns(2).Width + 25
            .Left = 5
            .Top = i_top - 5
            '.OwnerDraw = True
        End With

        fr_trans = New GroupBox
        With fr_trans
            .Text = "进制转换"
            .Left = 5
            .Top = i_top
            .Width = lv_hex.Width

        End With

        lb_input_num = New Label
        With lb_input_num
            .Text = "输入数字"
            .AutoSize = True
            .Top = 20
            .Left = 5
        End With

        tb_input_num = New TextBox
        With tb_input_num
            .Top = 20
            .Left = Me.lb_input_num.Width
        End With

        lb_show = New Label
        With lb_show
            .AutoSize = True
            .Top = 25
            .Left = tb_input_num.Left + tb_input_num.Width
        End With
        fr_trans.Height = tb_input_num.Height + 30

        i_top = Me.lv_hex.Top
        i_left = Me.lv_hex.Width + Me.lv_hex.Left + 5
        gb_write = New GroupBox
        With gb_write
            .Left = i_left
            .Top = i_top
            .Text = "改写文件"
        End With

        i_top = 25
        i_left = 5
        lb_wr = New Label
        With lb_wr
            .Text = "写入的文本:"
            .AutoSize = True
            .Left = i_left
            .Top = i_top + 5
        End With
        tb_wr = New TextBox
        With tb_wr
            .Left = 80 + lb_wr.Left
            .Top = i_top
        End With

        i_top += tb_wr.Height
        i_top += 5
        lb_start = New Label
        With lb_start
            .Text = "开始的地址:"
            .AutoSize = True
            .Left = i_left
            .Top = i_top + 5
        End With
        tb_start = New TextBox
        With tb_start
            .Left = 80 + lb_start.Left
            .Top = i_top
        End With
        btn_wr = New Button
        With btn_wr
            .Left = 5

            .Top = i_top + tb_start.Height
            .Text = "确定"
        End With

        btn_read_hex_to_file = New Button
        With btn_read_hex_to_file
            .Text = "读取HEX 另存文件"
            .Left = Me.lv_hex.Width + Me.lv_hex.Left + 5
            .Top = i_top + gb_write.Height
            .Width = gb_write.Width - 15
            .Height = 30
        End With
        Me.Controls.Add(btn_read_hex_to_file)


        Me.gb_write.Controls.Add(lb_wr)
        Me.gb_write.Controls.Add(tb_wr)
        Me.gb_write.Controls.Add(lb_start)
        Me.gb_write.Controls.Add(tb_start)
        Me.gb_write.Controls.Add(btn_wr)

        Me.gb_write.Width = tb_start.Left + tb_start.Width + 10
        Me.Controls.Add(gb_write)


        fr_trans.Controls.Add(lb_input_num)
        fr_trans.Controls.Add(tb_input_num)
        fr_trans.Controls.Add(lb_show)

        i_top = i_top + fr_trans.Height

        Me.Controls.Add(Me.lv_hex)
        Me.Controls.Add(lb_file_path)
        Me.Controls.Add(fr_trans)

        Me.Width = Me.lv_hex.Width + 20 + gb_write.Width
        Me.Height = 600

        Me.Text = "查看文件字节"
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()

    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Try
            Me.lv_hex.Height = Me.Height - Me.fr_trans.Height - 100
            Me.fr_trans.Top = Me.lv_hex.Height + Me.lv_hex.Top + 10
        Catch ex As Exception

        End Try

    End Sub


    Private Sub 选择文件ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 选择文件ToolStripMenuItem.Click
        Dim pfd As OpenFileDialog = New OpenFileDialog
        Dim file_name As String = ""
        If pfd.ShowDialog = DialogResult.OK Then
            file_name = pfd.FileName
            Me.lb_file_path.Text = file_name

            Dim file_byte() As Byte = Nothing
            If MFunc.read_file_to_byte(file_name, file_byte) = 1 Then
                add_arr_to_lv(file_byte)
            End If
        End If
    End Sub

    Function add_arr_to_lv(ByVal file_byte() As Byte)
        Dim n As Integer = CInt(file_byte.Length \ 16)
        If (file_byte.Length Mod 16) > 0 Then
            n += 1
        End If

        Dim Items(n - 1) As ListViewItem

        Dim str_add As String = ""

        For i As Integer = 0 To file_byte.Length - 1 Step 16
            Dim str As String = " "
            Dim str_asc As String = ""

            For j As Integer = 0 To 15
                If j = 8 Then
                    str = str & " "
                End If

                If (i + j) > (file_byte.Length - 1) Then
                    Exit For
                End If


                If file_byte(i + j) < 16 Then
                    str = str & "0" & Hex(file_byte(i + j)) & " "
                Else
                    str = str & Hex(file_byte(i + j)) & " "
                End If

                If file_byte(i + j) = 0 Then
                    str_asc &= " "
                Else
                    str_asc &= Chr(file_byte(i + j))
                End If

            Next

            str_add = Hex(i)
            If str_add.Length < 8 Then
                str_add = "00000000".Substring(str_add.Length, 8 - str_add.Length) & str_add & "H"
            End If
            Items(i / 16) = New ListViewItem((New String() {str_add, str, str_asc}))
        Next
        Me.lv_hex.Items.Clear()
        Me.lv_hex.Items.AddRange(Items)

        Return 0
    End Function

    Private Sub tb_input_num_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tb_input_num.KeyUp
        Dim str As String = Me.tb_input_num.Text

        If str = "" Then Exit Sub

        Try
            Dim n As Integer
            If str.Substring(0, 2) = "&H" Then
                n = Convert.ToInt16(str.Substring(2, str.Length - 2), 16)
            Else
                n = CInt(Me.tb_input_num.Text)
            End If

            Me.lb_show.Text = CInt(n).ToString & " = " & Hex(n) & "&H" & " = " & Oct(n) & "&O"
        Catch ex As Exception

        End Try

    End Sub

    Private Sub lb_file_path_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lb_file_path.DoubleClick
        Try
            System.Diagnostics.Process.Start(Me.lb_file_path.Text)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btn_wr_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_wr.Click
        Dim arr_byte() As Byte = System.Text.Encoding.Default.GetBytes(Me.tb_wr.Text)
        Dim start_address As Long = Convert.ToInt32(Me.tb_start.Text, 16)
        'MsgBox(start_address)
        If MFunc.write_byte_to_file(Me.lb_file_path.Text, arr_byte, start_address) = 1 Then
            MsgBox("OK")
        End If
    End Sub

    Private Sub lv_hex_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lv_hex.DoubleClick
        Dim item As ListViewItem = Me.lv_hex.SelectedItems(0)
        Dim str As String = item.SubItems(0).Text
        str = str.Substring(0, str.Length - 1)
        Me.tb_start.Text = str
    End Sub

    Private Sub btn_read_hex_to_file_Click(sender As Object, e As EventArgs) Handles btn_read_hex_to_file.Click
        Dim f As OpenFileDialog = New OpenFileDialog
        Dim read_file As String = ""

        f.Title = "选择内容为HEX的txt文件。"
        If f.ShowDialog = DialogResult.OK Then
            read_file = f.FileName
        End If
        If read_file <> "" Then
            Dim str_hex As String = MFunc.ReadFileToString(read_file)
            Dim file_byte(str_hex.Length / 2 - 1) As Byte
            For i As Integer = 0 To str_hex.Length - 1 Step 2
                file_byte(i / 2) = CByte("&H" & str_hex.Substring(i, 2))
            Next

            Dim sf As SaveFileDialog = New SaveFileDialog
            sf.Title = "请输入要保存的文件名。"
            If sf.ShowDialog = DialogResult.OK Then
                Dim save_file As String = sf.FileName
                If save_file <> "" Then
                    MFunc.write_byte_to_file(save_file, file_byte, 0)
                End If
            End If

        End If

    End Sub
End Class
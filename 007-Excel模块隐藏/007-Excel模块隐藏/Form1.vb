Public Class Form1

#Region "定义"
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents 菜单ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 选择文件ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsm_VBA As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsm_HideModule As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsm_UnHideModule As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsm_ReWritePROJECT As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsm_ReWritePROJECT_UnHide As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsm_UnProtectProject As System.Windows.Forms.ToolStripMenuItem

    Private WithEvents lv_hex As System.Windows.Forms.ListView
    Private WithEvents tree_dir As System.Windows.Forms.TreeView

    Private WithEvents btnVBA As System.Windows.Forms.Button

    Private gb_vba As System.Windows.Forms.GroupBox
    Private gb_workspace As System.Windows.Forms.GroupBox
    Private gb_vba_dir As System.Windows.Forms.Panel
    Private lb_tishi As System.Windows.Forms.Label

    Dim cls_cf As CCompdocFile
#End Region

#Region "From"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i_left As Integer = 5
        Dim i_top As Integer = 5
        Const i_HEIGHT As Integer = 300

        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.MenuStrip1.SuspendLayout()

        Me.菜单ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.选择文件ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsm_VBA = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsm_HideModule = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsm_UnHideModule = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsm_ReWritePROJECT = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsm_ReWritePROJECT_UnHide = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsm_UnProtectProject = New System.Windows.Forms.ToolStripMenuItem()
        '
        'MenuStrip1
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.菜单ToolStripMenuItem})
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        '菜单ToolStripMenuItem
        Me.菜单ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.选择文件ToolStripMenuItem, Me.tsm_VBA, Me.tsm_HideModule, Me.tsm_UnHideModule, Me.tsm_ReWritePROJECT, tsm_ReWritePROJECT_UnHide, tsm_UnProtectProject})
        Me.菜单ToolStripMenuItem.Text = "菜单"
        Me.菜单ToolStripMenuItem.Image = System.Drawing.SystemIcons.Information.ToBitmap
        '
        '选择文件ToolStripMenuItem        '
        Me.选择文件ToolStripMenuItem.Text = "选择文件"        '
        Me.选择文件ToolStripMenuItem.Image = System.Drawing.SystemIcons.Shield.ToBitmap

        Me.tsm_VBA.Text = "查找模块"        '
        tsm_VBA.Image = System.Drawing.SystemIcons.Question.ToBitmap

        tsm_HideModule.Text = "隐藏模块"        '
        Me.tsm_HideModule.Image = System.Drawing.SystemIcons.Hand.ToBitmap

        Me.tsm_UnHideModule.Text = "取消隐藏"        '
        Me.tsm_UnHideModule.Image = System.Drawing.SystemIcons.Asterisk.ToBitmap

        Me.tsm_ReWritePROJECT.Text = "隐藏ReWritePROJECT"        '
        Me.tsm_ReWritePROJECT.Image = System.Drawing.SystemIcons.Error.ToBitmap

        Me.tsm_ReWritePROJECT_UnHide.Text = "取消隐藏ReWritePROJECT"        '
        Me.tsm_ReWritePROJECT_UnHide.Image = System.Drawing.SystemIcons.Question.ToBitmap

        Me.tsm_UnProtectProject.Text = "VBA工程密码破解"        '
        Me.tsm_UnProtectProject.Image = System.Drawing.SystemIcons.Question.ToBitmap

        lb_tishi = New Label
        With lb_tishi
            .Top = 20
            .Left = 30
            .AutoSize = True
            .Text = "提示"
        End With

        i_top = 30
        tree_dir = New TreeView
        With tree_dir
            .Top = i_top
            .Left = 5
            .Height = i_HEIGHT
            .Width = 200
        End With

        i_left += tree_dir.Width

        lv_hex = New ListView
        With Me.lv_hex
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True
            .Sorting = SortOrder.None
            .Left = i_left
            .Top = i_top
            .Height = i_HEIGHT
            '.OwnerDraw = True
        End With

        gb_vba = New GroupBox
        gb_vba.Text = "Module"

        gb_workspace = New GroupBox
        gb_workspace.Text = "Workspace"

        gb_vba_dir = New Panel
        gb_vba_dir.Text = "vba_dir"

        Me.Width = tree_dir.Width + lv_hex.Width + 50
        Me.Height = i_HEIGHT + 100

        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1

        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()

        Me.Controls.Add(tree_dir)
        Me.Controls.Add(lv_hex)
        Me.Controls.Add(gb_vba)
        Me.Controls.Add(gb_vba_dir)
        Me.Controls.Add(gb_workspace)
        Me.Controls.Add(lb_tishi)

        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Try
            Me.lv_hex.Height = Me.Height - 400
            Me.lv_hex.Width = Me.Width - Me.tree_dir.Width - 50
            Me.tree_dir.Height = Me.lv_hex.Height
            Me.lb_tishi.Top = Me.lv_hex.Top + Me.lv_hex.Height
            Me.lb_tishi.Left = Me.lv_hex.Left
        Catch ex As Exception

        End Try

    End Sub

#End Region

#Region "Tree"
    Private Sub tree_dir_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles tree_dir.NodeMouseDoubleClick
        Dim str_node As String = ""
        Dim n As TreeNode = Nothing

        For Each n In Me.tree_dir.Nodes
            n = EachSubNode(n)
            If Not n Is Nothing Then Exit For
        Next

        Try
            str_node = n.Text
            If n.Parent.Text = "Dir" Then
                StreamToLV(str_node)
                Exit Sub
            End If
        Catch ex As Exception

        End Try


        Select Case str_node
            Case "HEAD"
                CFHeadToLV()
            Case "Dir"
                DirToLV()
        End Select
    End Sub

    Function EachSubNode(n As TreeNode) As TreeNode
        Dim sub_n As TreeNode = Nothing

        If n.IsSelected Then Return n
        For Each sub_n In n.Nodes
            If sub_n.IsSelected Then
                Return sub_n
            End If
        Next
        Return Nothing
    End Function
    Function TreeNode()
        Dim n As TreeNode = Nothing
        Dim str() As String = New String() {"HEAD", "MSAT", "SAT", "SSAT", "Dir"}

        Me.tree_dir.Nodes.Clear()

        For i As Integer = 0 To str.Length - 1
            n = New TreeNode
            n.Text = str(i)
            Me.tree_dir.Nodes.Add(n)
        Next

        Dim arr(,) As String = Nothing
        Dim i_row As Long = 0
        i_row = cls_cf.DirToArr(arr)
        For i As Integer = 0 To i_row
            Dim sub_n As TreeNode = New TreeNode
            If arr(i, 2) <> "0" Then
                sub_n.Text = arr(i, 1)
                n.Nodes.Add(sub_n)
            End If
        Next

        Return 0
    End Function

    Function StreamToLV(dir_name As String)
        Dim arr_byte() As Byte = Nothing
        Dim arr_address(,) As Integer = Nothing
        Dim if_short As Boolean
        Dim step_stream As Integer = 0
        Dim stream_len As Integer = 0
        Dim arr_font() As Integer = Nothing '如果是不连续的地址，标红
        Dim k_font As Integer = 0
        Dim pre_address As Integer = 0  '记录上一个地址

        If Not CheckCls() Then Return 0

        Dim i_row As Integer = cls_cf.GetStream(dir_name, arr_byte, stream_len, arr_address, if_short)

        With Me.lv_hex
            .Columns.Clear()
            .Columns.Add("address", 70, HorizontalAlignment.Right)
            .Columns.Add("00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F", 16 * 20, HorizontalAlignment.Left)
            .Columns.Add("asc", CInt(16 * 20 / 3) + 1)
        End With

        If if_short Then
            step_stream = 64
        Else
            step_stream = 512
        End If

        Dim Items((i_row + 1) * step_stream / 16 - 1) As ListViewItem
        Dim str(2) As String

        For i As Integer = 0 To i_row
            For j As Integer = 1 To step_stream / 16
                Dim start_address As Integer = arr_address(i, 1) + (j - 1) * 16

                If pre_address > 0 AndAlso (pre_address + 16) <> start_address Then
                    ReDim Preserve arr_font(k_font)
                    arr_font(k_font) = i * step_stream / 16 + j - 1
                    k_font += 1
                End If
                pre_address = start_address

                str(0) = Hex(start_address) '.PadLeft(8) & "H"
                str(1) = ""
                str(2) = ""
                For k As Integer = i * step_stream + (j - 1) * 16 To i * step_stream + (j - 1) * 16 + 16 - 1
                    If k >= stream_len Then Exit For
                    str(1) &= my_hex(arr_byte(k)).Substring(2, 2)
                    str(1) &= " "

                    If arr_byte(k) = 0 Then
                        str(2) &= "."
                    ElseIf arr_byte(k) = &HD OrElse arr_byte(k) = &HA Then
                        str(2) &= "-"
                    ElseIf arr_byte(k) < 128 Then
                        str(2) &= Chr(arr_byte(k))
                    Else
                        str(2) &= " "
                    End If
                Next
                Items(i * step_stream / 16 + j - 1) = New ListViewItem(str)
            Next
        Next

        LVAddItems(Items)

        For i As Integer = 0 To k_font - 1
            Me.lv_hex.Items(arr_font(i)).ForeColor = Color.Red
        Next
        lb_tishi.Text = Split(dir_name, vbNullChar)(0) & ":ITEMS " & Items.Length & "  RED " & k_font & "  if_short=" & if_short.ToString

        Return 0
    End Function

    Function DirToLV()
        Dim arr(,) As String = Nothing
        Dim i_row As Long = 0
        Dim arr_field() As String = New String() {"No.", "name", "len_name", "type", "color", "left_child", "right_child", "sub_dir", "time_create", "time_modify", "first_SID", "stream_size"}

        If cls_cf Is Nothing Then
            MsgBox("请选择文件。")
            Return -1
        End If

        i_row = cls_cf.DirToArr(arr)

        With Me.lv_hex
            .Columns.Clear()
            For i As Integer = 0 To arr_field.Length - 1
                .Columns.Add(arr_field(i), 50, HorizontalAlignment.Left)
            Next i
        End With

        Dim Items(i_row) As ListViewItem
        Dim i_col As Long = arr_field.Length - 1
        For i As Integer = 0 To i_row
            Dim str_item(i_col) As String
            For j As Integer = 0 To i_col
                str_item(j) = arr(i, j)
            Next

            Items(i) = New ListViewItem(str_item)
        Next


        LVAddItems(Items)

        Return 0
    End Function

    Function CFHeadToLV()
        Dim Items(16) As ListViewItem
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim str As String = "", str_asc As String = ""

        With Me.lv_hex
            .Columns.Clear()
            .Columns.Add("name", 70, HorizontalAlignment.Right)
            .Columns.Add("00 01 02 03 04 05 06 07  08 09 0A 0B 0C 0D 0E 0F", 16 * 20, HorizontalAlignment.Left)
            .Columns.Add("asc", CInt(16 * 20 / 3) + 1)
        End With

        If cls_cf Is Nothing Then
            MsgBox("请选择文件。")
            Return -1
        End If
        For j = 0 To cls_cf.cf_header.id.Length - 1
            str &= Hex(cls_cf.cf_header.id(j)) & " "
            str_asc &= Chr(cls_cf.cf_header.id(j))
        Next j
        Items(i) = New ListViewItem((New String() {"id", str, str_asc}))
        str = ""
        str_asc = ""
        i = i + 1

        'Dim file_id() As Byte             '文件唯一标识 
        For j = 0 To cls_cf.cf_header.file_id.Length - 1
            str &= Hex(cls_cf.cf_header.file_id(j)) & " "
            str_asc &= Chr(cls_cf.cf_header.file_id(j))
        Next j
        Items(i) = New ListViewItem((New String() {"file_id", str, str_asc}))
        str = ""
        str_asc = ""
        i = i + 1

        'Dim file_format_revision As Short '文件格式修订号
        Items(i) = New ListViewItem((New String() {"file_format_revision", Hex(cls_cf.cf_header.file_format_revision), cls_cf.cf_header.file_format_revision.ToString}))
        i = i + 1

        'Dim file_format_version As Short  '文件格式版本号
        Items(i) = New ListViewItem((New String() {"file_format_version", Hex(cls_cf.cf_header.file_format_version), cls_cf.cf_header.file_format_version.ToString}))
        i = i + 1
        'Dim memory_endian As Short        'FFFE表示 Little-Endian
        Items(i) = New ListViewItem((New String() {"memory_endian", Hex(cls_cf.cf_header.memory_endian), cls_cf.cf_header.memory_endian.ToString}))
        i = i + 1

        'Dim sector_size As Short          '扇区的大小 2的幂 通常为2^9=512
        Items(i) = New ListViewItem((New String() {"sector_size", Hex(cls_cf.cf_header.sector_size), cls_cf.cf_header.sector_size.ToString}))
        i = i + 1

        'Dim short_sector_size As Short    '短扇区大小，2的幂,通常为2^6
        Items(i) = New ListViewItem((New String() {"short_sector_size", Hex(cls_cf.cf_header.short_sector_size), cls_cf.cf_header.short_sector_size.ToString}))
        i = i + 1

        'Dim not_used_1() As Byte           '
        For j = 0 To cls_cf.cf_header.not_used_1.Length - 1
            str &= Hex(cls_cf.cf_header.not_used_1(j)) & " "
            str_asc &= Chr(cls_cf.cf_header.not_used_1(j))
        Next j
        Items(i) = New ListViewItem((New String() {"not_used_1", str, str_asc}))
        str = ""
        str_asc = ""
        i = i + 1

        'Dim SAT_count As Integer               '分区表扇区的总数
        Items(i) = New ListViewItem((New String() {"SAT_count", Hex(cls_cf.cf_header.SAT_count), cls_cf.cf_header.SAT_count.ToString}))
        i = i + 1

        'Dim dir_first_SID As Integer           '目录流第一个扇区的ID
        Items(i) = New ListViewItem((New String() {"dir_first_SID", Hex(cls_cf.cf_header.dir_first_SID), cls_cf.cf_header.dir_first_SID.ToString}))
        i = i + 1

        'Dim not_used_2() As Byte                '
        For j = 0 To cls_cf.cf_header.not_used_2.Length - 1
            str &= Hex(cls_cf.cf_header.not_used_2(j)) & " "
            str_asc &= Chr(cls_cf.cf_header.not_used_2(j))
        Next j
        Items(i) = New ListViewItem((New String() {"not_used_2", str, str_asc}))
        str = ""
        str_asc = ""
        i = i + 1

        'Dim min_stream_size As Integer         '最小标准流
        Items(i) = New ListViewItem((New String() {"min_stream_size", Hex(cls_cf.cf_header.min_stream_size), cls_cf.cf_header.min_stream_size.ToString}))
        i = i + 1

        'Dim SSAT_first_SID As Integer          '短分区表的第一个扇区ID
        Items(i) = New ListViewItem((New String() {"SSAT_first_SID", Hex(cls_cf.cf_header.SSAT_first_SID), cls_cf.cf_header.SSAT_first_SID.ToString}))
        i = i + 1

        'Dim SSAT_count As Integer              '短分区表扇区总数
        Items(i) = New ListViewItem((New String() {"SSAT_count", Hex(cls_cf.cf_header.SSAT_count), cls_cf.cf_header.SSAT_count.ToString}))
        i = i + 1

        'Dim MSAT_first_SID As Integer          '主分区表的第一个扇区ID
        Items(i) = New ListViewItem((New String() {"MSAT_first_SID", Hex(cls_cf.cf_header.MSAT_first_SID), cls_cf.cf_header.MSAT_first_SID.ToString}))
        i = i + 1

        'Dim MSAT_count As Integer              '分区表的扇区总数
        Items(i) = New ListViewItem((New String() {"MSAT_count", Hex(cls_cf.cf_header.MSAT_count), cls_cf.cf_header.MSAT_count.ToString}))
        i = i + 1

        'Dim arr_SID() As Integer            '主分区表前109个记录  108字节
        'For j = 0 To cls_cf.cf_Header.arr_SID.Length - 1
        '    str &= Hex(cls_cf.cf_Header.arr_SID(j)) & " "
        '    str_asc &= Chr(cls_cf.cf_Header.arr_SID(j))
        'Next j
        Items(i) = New ListViewItem((New String() {"arr_SID", str, cls_cf.cf_header.arr_SID.Length.ToString}))
        str = ""
        str_asc = ""
        i = i + 1

        LVAddItems(Items)
        Return 0
    End Function

    Function LVAddItems(Items() As ListViewItem)
        Me.lv_hex.Items.Clear()
        Me.lv_hex.Items.AddRange(Items)
        Me.lv_hex.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
        Me.lv_hex.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
        Return 0
    End Function
#End Region

#Region "菜单项"
    Private Sub 选择文件ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 选择文件ToolStripMenuItem.Click
        Dim fd As OpenFileDialog = New OpenFileDialog
        Dim file_name As String

        If fd.ShowDialog() = DialogResult.OK Then
            file_name = fd.FileName
        Else
            Exit Sub
        End If
        Me.Text = file_name

        cls_cf = New CCompdocFile(file_name)
        If CheckCls() Then TreeNode()
    End Sub

    Private Sub tsm_VBA_Click(sender As Object, e As EventArgs) Handles tsm_VBA.Click
        If Not CheckCls() Then Exit Sub

        Dim i_top As Integer = Me.tree_dir.Height + Me.tree_dir.Top
        gb_vba.Top = i_top + 15 : gb_workspace.Top = gb_vba.Top : gb_vba_dir.Top = gb_vba.Top
        gb_vba.Left = 5
        Me.gb_vba.Width = 300 : gb_workspace.Left = Me.gb_vba.Width + Me.gb_vba.Left + 5
        gb_workspace.Width = Me.gb_vba.Width : gb_vba_dir.Left = gb_workspace.Left + gb_workspace.Width + 5
        gb_vba_dir.Width = Me.gb_vba.Width

        Dim k_module As Integer = cls_cf.GetModule()
        Me.gb_vba.Controls.Clear()
        gb_workspace.Controls.Clear()

        If k_module > 0 Then
            i_top = 15
            For i = 0 To k_module - 1
                Dim cb As System.Windows.Forms.CheckBox = New CheckBox
                cb.Text = cls_cf.arr_Module(i).ModuleName
                cb.Width = gb_vba.Width - 10
                cb.Top = i_top
                gb_vba.Controls.Add(cb)

                i_top = i_top + cb.Height + 5
            Next i
            gb_vba.Height = i_top + 10

            i_top = 15
            For i = 0 To cls_cf.arr_Workspace.Length - 1
                Dim cb As System.Windows.Forms.CheckBox = New CheckBox
                cb.Text = cls_cf.arr_Workspace(i).Str
                cb.Width = gb_workspace.Width - 10
                cb.Top = i_top
                gb_workspace.Controls.Add(cb)

                i_top = i_top + cb.Height + 5
            Next
            gb_workspace.Height = i_top + 10


        End If
        'VBA DIR下的目录
        i_top = 15
        gb_vba_dir.Controls.Clear()

        For i = 0 To cls_cf.arr_VBA.Length - 1
            Dim str As String = cls_cf.arr_VBA(i)

            If (Not str Like "__SRP_*") AndAlso (Not cls_cf.dic_sheet.Contains(Split(str, vbNullChar)(0))) Then
                Dim TB As System.Windows.Forms.TextBox = New TextBox
                TB.Text = cls_cf.arr_VBA(i)
                TB.Width = gb_vba_dir.Width - 20
                TB.Top = i_top
                gb_vba_dir.Controls.Add(TB)

                i_top = i_top + TB.Height + 5
            End If


        Next
        gb_vba_dir.Height = i_top + 10
        If i_top > 150 Then
            gb_vba_dir.Height = 150
            gb_vba_dir.AutoScroll = True

        End If

    End Sub

    Private Sub tsm_HideModule_Click(sender As Object, e As EventArgs) Handles tsm_HideModule.Click
        If Not CheckCls() Then Exit Sub

        For Each ct As Control In Me.gb_vba.Controls
            If CType(ct, CheckBox).Checked Then
                If 1 = cls_cf.HideModule(CType(ct, CheckBox).Text) Then
                    tsm_VBA_Click(sender, e)
                End If
            End If
        Next
    End Sub

    Private Sub tsm_ReWritePROJECT_Click(sender As Object, e As EventArgs) Handles tsm_ReWritePROJECT.Click
        If Not CheckCls() Then Exit Sub

        For Each ct As Control In Me.gb_vba.Controls
            If CType(ct, CheckBox).Checked Then
                If 1 = cls_cf.ReWritePROJECT(CType(ct, CheckBox).Text) Then
                    tsm_VBA_Click(sender, e)
                End If
            End If
        Next
    End Sub

    Private Sub tsm_ReWritePROJECT_UnHide_Click(sender As Object, e As EventArgs) Handles tsm_ReWritePROJECT_UnHide.Click
        If Not CheckCls() Then Exit Sub

        Dim module_name As String = InputBox("输入模块的名称")
        If module_name = "" Then Exit Sub
        cls_cf.ReWritePROJECT(module_name, True)
        'tsm_VBA_Click(sender, e)
    End Sub

    Private Sub tsm_UnHideModule_Click(sender As Object, e As EventArgs) Handles tsm_UnHideModule.Click
        Dim ct_text As String
        Dim index_module As Integer = 0

        If Not CheckCls() Then Exit Sub

        For Each ct As Control In Me.gb_vba.Controls
            If CType(ct, CheckBox).Checked Then
                ct_text = CType(ct, CheckBox).Text
                If ct_text.Substring(0, 5) = "(隐藏的)" Then
                    Dim module_name As String = InputBox("输入模块的名称", "Module= 7个长度", "Module=")
                    If module_name = "" Then Exit Sub
                    cls_cf.UnHideModule(index_module, module_name)
                    tsm_VBA_Click(sender, e)
                End If
            End If
            index_module += 1
        Next


    End Sub

    Private Sub tsm_UnProtectProject_Click(sender As Object, e As EventArgs) Handles tsm_UnProtectProject.Click
        If Not CheckCls() Then Exit Sub
        cls_cf.ReWritePROJECT("xx", False, True)
    End Sub

#End Region

    Function CheckCls()
        If cls_cf Is Nothing Then
            Return False
        Else
            Return cls_cf.ready
        End If

    End Function
End Class

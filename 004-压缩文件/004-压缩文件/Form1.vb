Imports System.IO

Public Class Form1
    Private lbRar As System.Windows.Forms.Label
    Private WithEvents tbRarPath As System.Windows.Forms.TextBox '"E:\Program Files\WinRAR\WinRAR.exe"

    Private lbSavePath As System.Windows.Forms.Label
    Private WithEvents tbSavePath As System.Windows.Forms.TextBox

    Private lbRarList As System.Windows.Forms.Label
    Private WithEvents lvRarList As System.Windows.Forms.ListView

    Private WithEvents btnYaSuo As System.Windows.Forms.Button

    Private WithEvents cms As System.Windows.Forms.ContextMenuStrip
    Private WithEvents cmsAddPath As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents cmsAddFile As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents cmsDeletePath As System.Windows.Forms.ToolStripMenuItem

    Dim txtName As String = Application.StartupPath & "\data.txt"



    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim iLeft As Integer = 5
        Dim iTop As Integer = 5
        Dim iWidth As Integer = 0

        cms = New ContextMenuStrip
        cmsAddPath = New ToolStripMenuItem
        cmsAddPath.Text = "添加文件夹"
        cms.Items.Add(cmsAddPath)

        cmsAddFile = New ToolStripMenuItem
        cmsAddFile.Text = "添加文件"
        cms.Items.Add(cmsAddFile)

        cmsDeletePath = New ToolStripMenuItem
        cmsDeletePath.Text = "删除"
        cms.Items.Add(cmsDeletePath)


        lbRar = New Label
        Func.labelAdd(Me, lbRar, " RarPath", iLeft, iTop)
        iLeft += lbRar.Width

        tbRarPath = New TextBox
        Func.textBoxAdd(Me, tbRarPath, "", iLeft, iTop, 400)

        iLeft = 5
        iTop += tbRarPath.Height
        iTop += 5
        lbSavePath = New Label
        Func.labelAdd(Me, lbSavePath, "SavePath", iLeft, iTop)
        iLeft += lbSavePath.Width

        tbSavePath = New TextBox
        Func.textBoxAdd(Me, tbSavePath, "", iLeft, iTop, Me.tbRarPath.Width)

        iLeft = 5
        iTop += tbSavePath.Height

        lbRarList = New Label
        Func.labelAdd(Me, lbRarList, "RarList:", iLeft, iTop)

        iWidth = Me.lbRar.Width + Me.tbRarPath.Width
        iTop += lbRarList.Height
        iTop += 5
        lvRarList = New ListView
        With lvRarList
            .Columns.Add("序号", 50, HorizontalAlignment.Center)
            .Columns.Add("路径", iWidth - 50, HorizontalAlignment.Left)
            .View = View.Details
            .FullRowSelect = True
            .GridLines = True

            .ContextMenuStrip = cms
        End With
        Func.listViewAdd(Me, lvRarList, iLeft, iTop, iWidth, 300)

        iLeft = 5
        iTop += lvRarList.Height
        iTop += 10
        btnYaSuo = New Button
        Func.btnAdd(Me, btnYaSuo, "压缩", iLeft, iTop, iWidth)

        iTop += btnYaSuo.Height
        With Me
            .Width = iWidth + 20
            .Height = iTop + 50
            .StartPosition = FormStartPosition.CenterScreen
        End With

        readData()
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        writeData()
    End Sub

    Function readData()
        If File.Exists(txtName) Then
            Dim str As String = Func.ReadText(txtName)
            Dim Arr() As String = Split(str, vbNewLine)
            Try
                Me.tbRarPath.Text = Arr(0)
                Me.tbSavePath.Text = Arr(1)

                Dim items(Arr.Length - 1 - 2) As ListViewItem
                For i As Integer = 2 To Arr.Length - 1
                    items(i - 2) = New ListViewItem((New String() {i - 1, Arr(i)}))
                Next
                Me.lvRarList.Items.AddRange(items)
            Catch ex As Exception

            End Try

        End If
        Return 0
    End Function

    Function writeData()
        Dim str As String = ""
        str &= Me.tbRarPath.Text
        str &= vbNewLine
        str &= Me.tbSavePath.Text

        For i As Integer = 0 To Me.lvRarList.Items.Count - 1
            str &= vbNewLine
            str &= Me.lvRarList.Items(i).SubItems(1).Text
        Next

        Func.WriteText(txtName, str)
        Return 0
    End Function

    Private Sub btnYaSuo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnYaSuo.Click
        Dim sourePath(Me.lvRarList.Items.Count - 1) As String

        For i As Integer = 0 To Me.lvRarList.Items.Count - 1
            'yaSuo(Me.lvRarList.Items(i).SubItems(1).Text, Me.tbSavePath.Text)
            sourePath(i) = Me.lvRarList.Items(i).SubItems(1).Text
            sourePath(i) = Me.addFenHao(sourePath(i))
        Next
        yaSuo(Join(sourePath, " "), Me.tbSavePath.Text)

        MsgBox("OK")
    End Sub

    Function yaSuo(ByVal sourePath As String, ByVal saveName As String)
        saveName = Me.addFenHao(saveName)

        Dim strShell As String = Me.tbRarPath.Text & " a -ep1 " & saveName & " " & sourePath
        Dim result As Integer = Shell(strShell, vbHide)

        Do Until result > 0

        Loop

        Return result
    End Function

    Function addFenHao(ByVal str As String) As String
        If InStr(str, " ") > 0 Then
            Dim Arr() As String = Split(str, "\")
            For i As Integer = 0 To Arr.Length - 1
                If InStr(Arr(i), " ") > 0 Then
                    Arr(i) = """" & Arr(i) & """"
                End If
            Next
            str = Join(Arr, "\")
        End If

        Return str
    End Function

    Private Sub cmsDeletePath_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmsDeletePath.Click
        Me.lvRarList.Items.Remove(Me.lvRarList.SelectedItems(0))
    End Sub

    Private Sub cmsAddPath_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmsAddPath.Click
        Dim fbd As New System.Windows.Forms.FolderBrowserDialog
        Dim addPath As String = ""

        If fbd.ShowDialog = Windows.Forms.DialogResult.OK Then
            addPath = fbd.SelectedPath
        End If

        If addPath <> "" Then
            Dim item(0) As ListViewItem
            item(0) = New ListViewItem(New String() {Me.lvRarList.Items.Count + 1, addPath})
            Me.lvRarList.Items.AddRange(item)
        End If
    End Sub

    Private Sub lvRarList_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvRarList.DoubleClick
        System.Diagnostics.Process.Start(Me.lvRarList.SelectedItems(0).SubItems(1).Text)
    End Sub

    Private Sub cmsAddFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmsAddFile.Click
        Dim ofd As New System.Windows.Forms.OpenFileDialog

        Dim addFile As String = ""
        If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
            addFile = ofd.FileName
            Dim item(0) As ListViewItem
            item(0) = New ListViewItem(New String() {Me.lvRarList.Items.Count + 1, addFile})
            Me.lvRarList.Items.AddRange(item)
        End If
    End Sub
End Class

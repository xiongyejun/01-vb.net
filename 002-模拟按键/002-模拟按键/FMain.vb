Imports System.IO

Public Class FMain
    Dim iSleep As Integer = 0
    Const USER_NAME As String = "029013��033065��"  'system.Environment.UserName
    Const FU_HAO As String = "��"
    Const FILE_NAME As String = "ZhiLing.txt"

    '�����̹���
    Friend WithEvents MyHk As New CHook(False)


    Friend lbZhiLing As Label            'ָ���ǩ
    Friend WithEvents btnClearRTB As Button         '���ָ�ť
    Friend WithEvents rtbZhiLing As RichTextBox     'ָ���ı���
    Friend WithEvents lvKeys As ListView            '���ܼ��б��

    Friend WithEvents tbWenBen As TextBox           '���ָ����ı���
    Friend WithEvents btnAddWenBen As Button        '���ָ���ı��İ�ť
    Friend WithEvents btnZuoBiao As Button          '������깳�ӵİ�ť

    Friend lbSleep As Label              '�ӳ�ʱ���ǩ
    Friend nudSleep As NumericUpDown     '�ӳ�ʱ������
    Friend WithEvents btnStart As Button            '��ʼ��ť

    Friend lbLoop As Label              'ѭ������
    Friend nudLoop As NumericUpDown     'ѭ������

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MyHk.KeyHookEnabled = False
        MyHk.MouseHookEnabled = False

        Dim filename As String = Application.StartupPath & "\" & FILE_NAME
        If Me.rtbZhiLing.Text <> "" Then
            WriteText(filename, Me.rtbZhiLing.Text)
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim iTop As Integer = 10
        Dim iLeft As Integer = 10
        Const iStep As Integer = 30

        With Me
            .Size = New Size(400, 400)
            .FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
            .MaximizeBox = False
            .Text = "ģ�ⰴ��"
            '.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        End With

        lbZhiLing = New Label
        With lbZhiLing
            .Location = New Point(iLeft, iTop)
            .Text = "����ָ��:"
        End With
        btnClearRTB = New Button
        With btnClearRTB
            .Location = New Point(iLeft + lbZhiLing.Width + lbZhiLing.Left, iTop)
            .Text = "���ָ��"
        End With

        'ָ���ı���
        iTop += lbZhiLing.Height
        rtbZhiLing = New RichTextBox
        With rtbZhiLing
            .Location = New Point(iLeft, iTop)
            .Width = Me.Width - 20
            '.ReadOnly = True
        End With

        '���ܰ���ѡ���б�
        iTop += rtbZhiLing.Height
        lvKeys = New ListView
        With lvKeys
            .Location = New Point(iLeft, iTop)
            .Width = 230
            .Height = Me.Height - rtbZhiLing.Height - lbZhiLing.Height - 50

            .Columns.Add("���", 40, HorizontalAlignment.Left)
            .Columns.Add("��", 80, HorizontalAlignment.Left)
            .Columns.Add("����", 80)
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True

            .Items.AddRange(GetKeys)
        End With

        '�ı��༭
        iLeft += lvKeys.Width + 5

        tbWenBen = New TextBox
        With tbWenBen
            .Location = New Point(iLeft, iTop)
            .Width = Me.Height - lvKeys.Height - 30
        End With

        iTop += iStep
        btnAddWenBen = New Button
        With btnAddWenBen
            .Location = New Point(iLeft, iTop)
            .Text = "����ı�"
        End With

        '�������
        iTop += iStep
        btnZuoBiao = New Button
        With btnZuoBiao
            .Location = New Point(iLeft, iTop)
            .Text = "����MouseHook"
            .Width = 100
        End With

        '�����ӳ�ʱ��
        iTop += iStep
        lbSleep = New Label
        With lbSleep
            .Location = New Point(iLeft, iTop + 5)
            .Text = "�����ӳ�ʱ��(����)"
            .Width = 120
        End With

        iTop += iStep
        nudSleep = New NumericUpDown
        With nudSleep
            .Location = New Point(iLeft, iTop)
            .Width = 50
            .Minimum = 200
            .Maximum = 10000
            .Increment = 200
            .Value = 1000
        End With

        iTop += iStep
        lbLoop = New Label
        With lbLoop
            .Location = New Point(iLeft, iTop + 5)
            .Text = "����ѭ������"
            .Width = 120
        End With

        iTop += iStep
        nudLoop = New NumericUpDown
        With nudLoop
            .Location = New Point(iLeft, iTop)
            .Width = 50
            .Minimum = 1
            .Maximum = 1000
            .Increment = 1
            .Value = 1
        End With

        '��ʼ��ť
        iTop += iStep
        btnStart = New Button
        With btnStart
            .Location = New Point(iLeft, iTop)
            .Text = "Start"
        End With


        With Me.Controls
            .Add(lbZhiLing)
            .Add(btnClearRTB)
            .Add(rtbZhiLing)
            .Add(lvKeys)

            .Add(tbWenBen)
            .Add(btnAddWenBen)

            .Add(btnZuoBiao)

            .Add(btnStart)
            .Add(nudSleep)
            .Add(lbSleep)

            .Add(nudLoop)
            .Add(lbLoop)
        End With

        Dim filename As String = Application.StartupPath & "\" & FILE_NAME
        If File.Exists(filename) Then
            Me.rtbZhiLing.Text = ReadText(filename)
        End If
    End Sub

    Private Sub lvKeys_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvKeys.DoubleClick
        Dim str As String = Me.lvKeys.SelectedItems(0).SubItems(2).Text
        rtbTextAdd(str)
    End Sub

    Private Sub btnAddWenBen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddWenBen.Click
        Dim str As String = Me.tbWenBen.Text
        rtbTextAdd(str)
    End Sub

    Private Sub btnZuoBiao_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnZuoBiao.Click
        With Me.btnZuoBiao
            If .Text = "����MouseHook" Then
                .Text = "�ر�MouseHook"
                MyHk.KeyHookEnabled = True
                MyHk.MouseHookEnabled = True
            Else
                .Text = "����MouseHook"
                MyHk.KeyHookEnabled = False
                MyHk.MouseHookEnabled = False
            End If
        End With
    End Sub

    Private Sub btnClearRTB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClearRTB.Click
        Me.rtbZhiLing.Text = ""
    End Sub

    Private Sub btnStart_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Dim Arr() As String = Split(Me.rtbZhiLing.Text, FU_HAO)
        iSleep = Me.nudSleep.Value
        Me.WindowState = FormWindowState.Minimized
        System.Threading.Thread.Sleep(iSleep)

        For j As Integer = 0 To Val(nudLoop.Value) - 1
            For i As Integer = 0 To Arr.Length - 1
                If Strings.Left(Arr(i), 2) = "X(" Then
                    Dim strTmp As String = Strings.Mid(Arr(i), 3, Len(Arr(i)) - 2)
                    Dim temp() As String = Split(strTmp, ",")
                    Screen_Click(Val(temp(1)), Val(temp(2)))
                    System.Threading.Thread.Sleep(Val(temp(0)))
                Else
                    ShuRu(Arr(i))
                    System.Threading.Thread.Sleep(iSleep)
                End If

            Next
        Next

        Erase Arr
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub rtbTextAdd(ByVal str As String)
        Dim Str1 As String = ""
        Dim i As Integer = Me.rtbZhiLing.SelectionStart

        With Me.rtbZhiLing
            If .Text = "" Then
                .Text = str
            Else
                If i = 0 Then
                    str = str & FU_HAO
                Else
                    str = FU_HAO & str
                End If

                .Text = .Text.Insert(i, str)   '��ӷ���
            End If

            i = i + Len(str)
            .SelectionStart = i

        End With
    End Sub

    Private Sub MyHk_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyHk.KeyDown
        Dim iCode As Integer = e.KeyCode
        If iCode = 112 Then
            Me.btnZuoBiao.Text = "����MouseHook"
            MyHk.KeyHookEnabled = False
            MyHk.MouseHookEnabled = False
        End If
    End Sub



    Private Sub MyHk_MouseActivity(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyHk.MouseActivity

        Dim x As Integer = System.Windows.Forms.Control.MousePosition.X
        Dim y As Integer = System.Windows.Forms.Control.MousePosition.Y
        Me.tbWenBen.Text = "X(3000," & x & "," & y & ")Y"

    End Sub
End Class

Imports System.Threading

Public Class FWrite
#Region "声明"
    Enum baoMi
        shiJuan = 0
        answer = 1
        question = 2
        leiXing = 3
        sign = 4
        row = 5
    End Enum

    Enum tiXing
        tianKong = 1
        panDuan
        danXuan
        duoXuan
        jianDa
    End Enum

    Const fileName As String = "\题库.accdb"
    Const tableName As String = "题库"
    Const passWordTxt As String = "\passWord.txt"

    'Dim dataArr(,) As System.Object = Nothing
    Dim dt As DataTable
    Dim scrollValue As Integer = 0    '

    Private WithEvents ms As System.Windows.Forms.MenuStrip
    Private WithEvents msItem As System.Windows.Forms.ToolStripMenuItem

    Private WithEvents gbShiJuan As System.Windows.Forms.GroupBox
    Private WithEvents panelQuestions As System.Windows.Forms.Panel
    Private WithEvents btnTiJiao As System.Windows.Forms.Button
    Private lbScore As System.Windows.Forms.Label               '得分
    Private WithEvents tTimer As System.Windows.Forms.Timer     '计时器

    Dim questionsCount As Integer = 0    '记录问题的个数
    Dim iTimer As Integer = 0               '用时
#End Region

#Region "Form事件"

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False '允许跨线程访问控件

        Dim iWidth As Integer = Screen.PrimaryScreen.Bounds.Width
        Dim iTop As Integer = 5
        Control.CheckForIllegalCrossThreadCalls = False

        iTop += 20
        gbShiJuan = New GroupBox
        ControlAdd.groupBoxAdd(Me, gbShiJuan, "试卷", 5, iTop, iWidth - 15, 40)

        For i As Integer = 1 To 20
            Dim cb As New RadioButton
            ControlAdd.radioButtonAdd(gbShiJuan, cb, i.ToString, (i - 1) * 40 + 5, 12, 40, 30, False)
            AddHandler cb.Click, AddressOf addQuestions

        Next

        iTop += Me.gbShiJuan.Height

        ms = New MenuStrip
        ms.Text = "切换"
        Me.MainMenuStrip = ms

        msItem = New ToolStripMenuItem
        msItem.Text = "切换到读卷模式"
        msItem.Image = SystemIcons.Application.ToBitmap
        ms.Items.Add(msItem)

        Me.Controls.Add(ms)
        With Me
            .Width = iWidth
            .Height = Screen.PrimaryScreen.Bounds.Height - 100
            .Location = New Point((Screen.PrimaryScreen.Bounds.Width - Me.Width) / 2, (Screen.PrimaryScreen.Bounds.Height - Me.Height) / 2)
            .Text = "背题"
            .WindowState = FormWindowState.Maximized
        End With


        btnTiJiao = New Button
        ControlAdd.btnAdd(Me, btnTiJiao, "提交", 5, Me.Height - 70)
        btnTiJiao.Enabled = False
        lbScore = New Label
        ControlAdd.labelBoxAdd(Me, lbScore, "得分:", btnTiJiao.Left + btnTiJiao.Width, btnTiJiao.Top)

        tTimer = New System.Windows.Forms.Timer
        Me.tTimer.Enabled = False

        Dim c_ado = New CADO(Application.StartupPath & fileName)
        c_ado.StrSql = "Select * From " & tableName
        dt = c_ado.GetData()
        c_ado = Nothing
        'Func.CreateAdoArr("Select * From " & tableName, Application.StartupPath & fileName, dataArr)
    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Dim t As Thread = New Thread(AddressOf controlResize)
        t.Start()
    End Sub
#End Region


    Sub addQuestions(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim iTop As Integer = 5
        Const xuanXiangHeight As Integer = 60 '选择题的height

        Me.btnTiJiao.Enabled = True
        Me.lbScore.Text = "正在答卷……"


        questionsCount = 0
        Me.Controls.Remove(Me.panelQuestions)
        panelQuestions = New Panel
        With panelQuestions
            .AutoScroll = True
            .Left = 5
            .Width = Me.Width - 15
            .Height = Me.Height - Me.gbShiJuan.Height - Me.gbShiJuan.Top - 80
            .Top = Me.gbShiJuan.Top + Me.gbShiJuan.Height + 5
            .BorderStyle = BorderStyle.FixedSingle
            .Font = New System.Drawing.Font("宋体", 16, System.Drawing.FontStyle.Bold)

            AddHandler panelQuestions.Click, AddressOf Focus_Click
        End With
        Me.Controls.Add(panelQuestions)

        '在panelQuestions里为每个问题添加一个gb
        Dim iShiJuan As Integer = CInt(sender.text)
        For i As Integer = 0 To dt.Rows.Count - 1 ' dataArr.Length / 6 - 1
            If dt.Rows(i)(baoMi.shiJuan) = iShiJuan Then
                Dim gb As New GroupBox
                With gb
                    .Tag = i.ToString             '记录问题的下标
                    AddHandler .Click, AddressOf Focus_Click
                    AddHandler .MouseWheel, AddressOf panelQuestions_MouseWheel
                End With

                'ControlAdd.groupBoxAdd(Me.panelQuestions, gb, dataArr(baoMi.question, gb.Tag), 5, 0, Me.panelQuestions.Width - 15, 100)
                ControlAdd.groupBoxAdd(Me.panelQuestions, gb, dt.Rows(gb.Tag)(baoMi.question), 5, 0, Me.panelQuestions.Width - 15, 100)
                questionsCount += 1             '对应问题
            End If
        Next

        For i As Integer = 0 To questionsCount - 1
            Dim ct As System.Windows.Forms.Control = Me.panelQuestions.Controls(i)

            Select Case dt.Rows(CInt(ct.Tag))(baoMi.leiXing)'dataArr(baoMi.leiXing, CInt(ct.Tag))

                Case tiXing.tianKong '填空题——添加一个textbox作为回答用
                    Dim tb As TextBox = New TextBox
                    ControlAdd.textBoxAdd(ct, tb, 5, 50, 500)
                    tb.Tag = i.ToString             '记录Controls(i),响应enter到下一个控件
                    AddHandler tb.KeyDown, AddressOf tb_Enter

                    iTop = Me.setPanelControls(ct, 90, iTop)

                Case tiXing.panDuan                    '判断——添加2个radiobutton供选择
                    Dim rb1 As RadioButton = New RadioButton
                    ControlAdd.radioButtonAdd(ct, rb1, "√", 45, 55, 50, 30, False)
                    Dim rb2 As RadioButton = New RadioButton
                    ControlAdd.radioButtonAdd(ct, rb2, "×", 100, 55, 50, 30, False)
                    iTop = Me.setPanelControls(ct, 90, iTop)

                Case tiXing.danXuan '单选，将答案作为radiobutton供选择
                    With ct
                        Dim tempArr() As String = Split(.Text, Chr(10))
                        .Text = tempArr(0)
                        For j As Integer = 1 To tempArr.Length - 1
                            Dim rb As New RadioButton
                            ControlAdd.radioButtonAdd(ct, rb, tempArr(j), 20, (j - 1) * xuanXiangHeight + 45, 50, xuanXiangHeight)
                        Next
                        iTop = Me.setPanelControls(ct, tempArr.Length * xuanXiangHeight, iTop)
                    End With

                Case tiXing.duoXuan '多选，将答案作为checkbox供选择
                    With ct
                        Dim tempArr() As String = Split(.Text, Chr(10))
                        .Text = tempArr(0)
                        For j As Integer = 1 To tempArr.Length - 1
                            Dim rb As New CheckBox
                            ControlAdd.checkBoxAdd(ct, rb, tempArr(j), 20, (j - 1) * xuanXiangHeight + 45, 50, xuanXiangHeight)
                            rb.AutoSize = True
                        Next
                        iTop = Me.setPanelControls(ct, tempArr.Length * xuanXiangHeight, iTop)

                    End With
                Case tiXing.jianDa '简答题——添加一个richbox
                    Dim rtb As New RichTextBox
                    ControlAdd.richTextBoxAdd(ct, rtb, 5, 20, 500, 300)
                    iTop = Me.setPanelControls(ct, rtb.Height + 30, iTop)
                    AddHandler rtb.MouseWheel, AddressOf panelQuestions_MouseWheel
            End Select
        Next

    End Sub
    Private Sub tb_Enter(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            Dim i As Integer = CInt(sender.tag) + 1
            'If dataArr(baoMi.leiXing, CInt(Me.panelQuestions.Controls(i).Tag)) = tiXing.tianKong Then
            If dt.Rows(CInt(Me.panelQuestions.Controls(i).Tag))(baoMi.leiXing) = tiXing.tianKong Then
                Me.panelQuestions.Controls(i).Controls(0).Focus()
            End If

        End If
    End Sub

    Sub controlResize()
        Try
            Me.gbShiJuan.Width = Me.Width - 15
            btnTiJiao.Top = Me.Height - 70
            lbScore.Top = btnTiJiao.Top

            panelQuestions.Width = Me.Width - 15
            panelQuestions.Height = Me.Height - Me.gbShiJuan.Height - Me.gbShiJuan.Top - 80
            For Each ct As Control In Me.panelQuestions.Controls
                If TypeName(ct) = "GroupBox" Then
                    ct.Width = Me.panelQuestions.Width - 15
                End If
            Next
        Catch ex As Exception

        End Try

        'Me.Refresh()
    End Sub

    Private Sub panelQuestions_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles panelQuestions.MouseWheel
        Dim scrollValue As Integer = Me.panelQuestions.VerticalScroll.Value
        Dim iMax As Integer = Me.panelQuestions.VerticalScroll.Maximum
        Dim iMin As Integer = Me.panelQuestions.VerticalScroll.Minimum

        scrollValue -= (e.Delta / 2)
        If scrollValue > iMax Then scrollValue = iMax
        If scrollValue < iMin Then scrollValue = iMin
        Me.panelQuestions.VerticalScroll.Value = scrollValue

    End Sub

    Private Sub Focus_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Focus()
    End Sub

    Function setPanelControls(ByVal ctr As System.Windows.Forms.Control, ByVal iHeight As Integer, ByVal iTop As Integer) As Integer
        With ctr
            .Height = iHeight
            .Top = iTop
        End With

        Return iTop + ctr.Height + 10
    End Function

    Private Sub btnTiJiao_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTiJiao.Click
        Dim iData As Integer = 0
        Dim deFenArr(tiXing.jianDa - 1, 1) As Integer '0列正确数量、1列错误数量

        Const tempStr As String = "ABCDEFGHIJ"
        Me.btnTiJiao.Enabled = False
        Me.tTimer.Enabled = False

        For i As Integer = 0 To Me.questionsCount - 1
            Dim ct As System.Windows.Forms.Control = Me.panelQuestions.Controls(i)

            iData = CInt(ct.Tag)
            Dim answerStr As String = dt.Rows(iData)(baoMi.answer).ToString ' dataArr(baoMi.answer, iData)  '标准答案
            Dim str As String = ""                                  '回答的答案

            Select Case dt.Rows(iData)(baoMi.leiXing) 'dataArr(baoMi.leiXing, iData)
                Case tiXing.tianKong '填空题
                    str = Func.clearChar(ct.Controls(0).Text)
                    CType(ct.Controls(0), TextBox).ReadOnly = True
                    answerStr = Func.clearChar(answerStr)

                Case tiXing.panDuan  '判断

                    If CType(ct.Controls(0), RadioButton).Checked Then
                        str = "√"
                    ElseIf CType(ct.Controls(1), RadioButton).Checked Then
                        str = "×"
                    End If
                Case tiXing.danXuan  '单选

                    For j As Integer = 0 To ct.Controls.Count - 1
                        If CType(ct.Controls(j), RadioButton).Checked Then
                            str = tempStr.Substring(j, 1)
                            Exit For
                        End If
                    Next
                Case tiXing.duoXuan '多选

                    For j As Integer = 0 To ct.Controls.Count - 1
                        If CType(ct.Controls(j), CheckBox).Checked Then
                            str &= tempStr.Substring(j, 1)
                        End If
                    Next

                Case tiXing.jianDa '简单题
                    Dim rtb2 As New RichTextBox
                    ControlAdd.richTextBoxAdd(ct, rtb2, 10 + ct.Controls(0).Width, ct.Controls(0).Top, ct.Controls(0).Width, ct.Controls(0).Height)
                    rtb2.Text = answerStr
                    AddHandler rtb2.MouseWheel, AddressOf panelQuestions_MouseWheel
                    Dim LCD As Double = Me.compar_LCD(ct.Controls(0), rtb2)
                    CType(ct.Controls(0), RichTextBox).ReadOnly = True
                    rtb2.ReadOnly = True

                    If LCD > 0.5 Then
                        str = answerStr
                        rtb2.AppendText(Chr(10) & Format(LCD, "准确率0.00%，正确"))
                    Else
                        rtb2.AppendText(Chr(10) & Format(LCD, "准确率0.00%，错误"))
                    End If

            End Select

            If str = answerStr Then
                'deFenArr(dataArr(baoMi.leiXing, iData) - 1, 0) += 1
                deFenArr(dt.Rows(iData)(baoMi.leiXing) - 1, 0) += 1
            Else
                deFenArr(dt.Rows(iData)(baoMi.leiXing) - 1, 1) += 1
                ct.ForeColor = Color.Red
                If dt.Rows(iData)(baoMi.leiXing) <> tiXing.jianDa Then
                    Dim lb As New Label
                    'ControlAdd.labelBoxAdd(ct, lb, dataArr(baoMi.answer, iData), 500, 50)
                    ControlAdd.labelBoxAdd(ct, lb, dt.Rows(iData)(baoMi.answer), 500, 50)
                    lb.ForeColor = Color.Red
                    lb.BringToFront()

                End If


            End If
        Next

        Dim msgStr As String = getScore(deFenArr)
        MsgBox(msgStr)

        'getPic()
    End Sub

    Function getScore(ByVal deFenArr(,) As Integer) As String
        Dim sumDeFen As Integer = 0
        Dim msgStr As String = ""

        For i As Integer = 0 To 4
            Dim tiXing As String = ""
            Dim iFenShu As Integer = 0

            Select Case i
                Case 0
                    tiXing = "填空题"
                    iFenShu = 2
                Case 1
                    tiXing = "判断题"
                    iFenShu = 1
                Case 2
                    tiXing = "单选题"
                    iFenShu = 1
                Case 3
                    tiXing = "多选题"
                    iFenShu = 3
                Case 4
                    tiXing = "简单题"
                    iFenShu = 5
            End Select
            msgStr &= (tiXing & deFenArr(i, 0) + deFenArr(i, 1) & "个，正确" & deFenArr(i, 0) & "个，得分：" & deFenArr(i, 0) * iFenShu)
            sumDeFen += deFenArr(i, 0) * iFenShu
            msgStr &= Chr(10)
        Next
        lbScore.Text = "得分：" & sumDeFen & vbNewLine & "用时："
        msgStr &= "总得分：" & sumDeFen

        Return msgStr
    End Function

    '对比最长公共子序列并标识出来
    Function compar_LCD(ByVal rtb1 As RichTextBox, ByVal rtb2 As RichTextBox) As Double
        Dim str1() As String = {}, str2() As String = {}
        Dim index1() As Integer = {}, index2() As Integer = {}

        If rtb1.Text = "" Then
            Return 0
        End If

        LCD_Arr(rtb1.Text, str1, index1)
        LCD_Arr(rtb2.Text, str2, index2)

        Dim arrLength(,) As Integer, arrPointer(,) As String
        Dim m As Integer = str1.Length - 1
        Dim n As Integer = str2.Length - 1

        ReDim arrLength(m, n)
        ReDim arrPointer(m, n)

        For i As Integer = 1 To m
            arrLength(i, 0) = 0
        Next i

        For j As Integer = 0 To n
            arrLength(0, j) = 0
        Next j

        For i As Integer = 1 To m
            For j As Integer = 1 To n

                If str1(i) = str2(j) Then
                    arrLength(i, j) = arrLength(i - 1, j - 1) + 1
                    arrPointer(i, j) = "↖"
                ElseIf arrLength(i - 1, j) >= arrLength(i, j - 1) Then
                    arrLength(i, j) = arrLength(i - 1, j)
                    arrPointer(i, j) = "↑"
                Else
                    arrLength(i, j) = arrLength(i, j - 1)
                    arrPointer(i, j) = "←"
                End If
            Next j
        Next i

        rtb1.Font = New Font("宋体", 10, System.Drawing.FontStyle.Bold)
        rtb2.Font = New Font("宋体", 10, System.Drawing.FontStyle.Bold)
        print_LCS(rtb1, rtb2, arrPointer, index1, index2, m, n)

        Return arrLength(m, n) / n
    End Function

    Function LCD_Arr(ByVal str As String, ByRef arr() As String, ByRef index() As Integer)
        Dim k As Integer = 1

        Replace(str, Chr(10), "")

        For i As Integer = 0 To str.Length - 1
            Dim tempStr As String = str.Substring(i, 1)
            If InStr(SPE_CHAR, tempStr) = 0 Then
                ReDim Preserve arr(k)                  '字符
                ReDim Preserve index(k)                 '字符对应的原rtb中的位置

                arr(k) = tempStr
                index(k) = i
                k += 1
            End If
        Next

        Return 0
    End Function

    Function print_LCS(ByVal rtb1 As RichTextBox, ByVal rtb2 As RichTextBox, ByVal arrPointer(,) As String, ByVal index1() As Integer, ByVal index2() As Integer, ByVal i As Integer, ByVal j As Integer)
        If i = 0 Or j = 0 Then
            Return 0
            Exit Function
        ElseIf arrPointer(i, j) = "↖" Then
            rtb1.Select(index1(i), 1)
            setRtbSelectFont(rtb1)
            rtb2.Select(index2(j), 1)
            setRtbSelectFont(rtb2)

            print_LCS(rtb1, rtb2, arrPointer, index1, index2, i - 1, j - 1)

        ElseIf arrPointer(i, j) = "↑" Then
            print_LCS(rtb1, rtb2, arrPointer, index1, index2, i - 1, j)
        Else
            print_LCS(rtb1, rtb2, arrPointer, index1, index2, i, j - 1)
        End If

        Return 0
    End Function

    Function setRtbSelectFont(ByVal rtb1 As RichTextBox)
        rtb1.SelectionFont = New Font("宋体", 16, System.Drawing.FontStyle.Bold)
        rtb1.SelectionColor = Color.Red
        Return 0
    End Function

    Private Sub msItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles msItem.Click

        FRead.Show()
        Me.Hide()
    End Sub
End Class

Imports System.Drawing

Public Class FRead
    Enum baoMi
        shiJuan = 0
        answer = 1
        question = 2
        leiXing = 3
        sign = 4
        row = 5
    End Enum

    Const fileName As String = "\题库.accdb"
    Const shiZhi As String = "\设置.txt"
    Const tableName As String = "[题库]"
    Const LEI_XING = "填空题、判断题、单项选择题、多项选择题、简答题"

    Private WithEvents gbShiJuan As System.Windows.Forms.GroupBox
    Private WithEvents gbLeiXing As System.Windows.Forms.GroupBox
    Private WithEvents gbSign As System.Windows.Forms.GroupBox

    Private WithEvents lbQuestion As System.Windows.Forms.Label
    Private WithEvents lbAnswer As System.Windows.Forms.Label

    Private WithEvents gbBtn As System.Windows.Forms.GroupBox
    Private WithEvents btnPre As System.Windows.Forms.Button
    Private WithEvents btnNext As System.Windows.Forms.Button
    Private WithEvents btnAnswer As System.Windows.Forms.Button
    Private opRnd As System.Windows.Forms.RadioButton
    Private opSort As System.Windows.Forms.RadioButton
    Private WithEvents nudFont As System.Windows.Forms.NumericUpDown

    Private WithEvents btnAddSign As System.Windows.Forms.Button
    Private WithEvents btnRemoveSign As System.Windows.Forms.Button
    Private WithEvents btnRemoveAllSign As System.Windows.Forms.Button

    Private lbTiShi As System.Windows.Forms.Label

    Private WithEvents cms As ContextMenuStrip
    Private WithEvents cmsAddSign As New System.Windows.Forms.ToolStripMenuItem
    Private WithEvents cmsRemoveSign As New System.Windows.Forms.ToolStripMenuItem
    Private WithEvents cmsRemoveAllSign As New System.Windows.Forms.ToolStripMenuItem

    Private WithEvents ms As System.Windows.Forms.MenuStrip
    Private WithEvents msItem As System.Windows.Forms.ToolStripMenuItem

    Dim dt As DataTable
    'Dim dataArr(,) As System.Object = Nothing
    'Dim signArr() As Integer                '保存初始的sign，更新时只替换有变化的

    Dim iData As Integer = 0
    Dim maxSign As Integer = 1
    Dim ifSave As Boolean = False
    Dim adoResult As Integer = 100

    Private Sub Form2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click
        Dim g As Graphics = Me.gbShiJuan.CreateGraphics
        Dim drawFont As New Font("Arial", 16)
        Dim drawBrush As New SolidBrush(Color.Red)
        Dim drawString As String = "单击此处可全选，双击取消"

        g.DrawString(drawString, drawFont, drawBrush, New PointF(555, 12))
    End Sub

    Private Sub Form2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        saveExcel()
        saveSheZhi()

        FWrite.Close()
    End Sub


    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False '允许跨线程访问控件

        Const iWidth As Integer = 700
        Dim iLeft As Integer = 5
        Const iHeight As Integer = 200

        Dim iMax(,) As System.Object = Nothing, i As Integer
        Dim iTop As Integer = 0

        ms = New MenuStrip
        ms.Text = "切换"
        Me.MainMenuStrip = ms
        Me.Controls.Add(ms)

        msItem = New ToolStripMenuItem
        msItem.Text = "切换到答卷模式"
        msItem.Image = SystemIcons.Application.ToBitmap
        ms.Items.Add(msItem)

        cmsAddSign.Text = "添加标记"
        cmsRemoveSign.Text = "取消标记"
        cmsRemoveAllSign.Text = "取消当前所有标记"

        cms = New System.Windows.Forms.ContextMenuStrip
        With cms.Items
            .Add(cmsAddSign)
            .Add(cmsRemoveSign)
            .Add(cmsRemoveAllSign)
        End With

        iTop += 20

        gbShiJuan = New GroupBox
        ControlAdd.groupBoxAdd(Me, gbShiJuan, "试卷", iLeft, iTop, iWidth, 40)
        Func.CreateAdoArr("Select Max(试卷) From " & tableName, Application.StartupPath & fileName, iMax)

        For i = 1 To CInt(iMax(0, 0))
            cbAdd(gbShiJuan, (i - 1) * 50 + 5, 12, 50, i.ToString)
        Next

        iTop += Me.gbShiJuan.Height
        iTop += 5

        gbLeiXing = New GroupBox
        ControlAdd.groupBoxAdd(Me, gbLeiXing, "题型", iLeft, iTop, iWidth, 40)


        Dim tempArr() As String = Split(LEI_XING, "、")
        For i = 0 To tempArr.Length - 1
            cbAdd(gbLeiXing, i * 100 + 5, 12, 100, tempArr(i))
        Next

        iTop += Me.gbShiJuan.Height
        iTop += 5

        gbSign = New GroupBox
        ControlAdd.groupBoxAdd(Me, gbSign, "标记", iLeft, iTop, 180, 40)

        cbAdd(gbSign, 5, 12, 100, "未标记")
        cbAdd(gbSign, 105 + 5, 12, 100, "已标记")

        lbTiShi = New Label
        With lbTiShi
            .Font = New System.Drawing.Font("宋体", 20, System.Drawing.FontStyle.Bold)
            .ForeColor = Color.Red
            .AutoSize = True
        End With
        ControlAdd.labelBoxAdd(Me, lbTiShi, "", gbSign.Left + gbSign.Width + 5, gbSign.Top + 5)

        iTop += Me.gbSign.Height
        iTop += 5
        lbQuestion = New Label
        With lbQuestion
            .BackColor = Color.White
            .BorderStyle = BorderStyle.FixedSingle
            .Font = New System.Drawing.Font("宋体", 20, System.Drawing.FontStyle.Bold)
            '.Font = New System.Drawing.Font("宋体", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
            .ContextMenuStrip = cms
        End With
        ControlAdd.labelBoxAdd(Me, lbQuestion, "", iLeft, iTop, iWidth, iHeight)


        iTop += Me.lbQuestion.Height
        iTop += 5
        lbAnswer = New Label
        With lbAnswer
            .BackColor = Color.White
            .BorderStyle = BorderStyle.FixedSingle
            .Font = New System.Drawing.Font("宋体", 20, System.Drawing.FontStyle.Bold)
            .ForeColor = Color.Red
            .ContextMenuStrip = cms
        End With
        ControlAdd.labelBoxAdd(Me, lbAnswer, "", iLeft, iTop, iWidth, iHeight)

        iTop += Me.lbAnswer.Height
        iTop += 5
        gbBtn = New GroupBox
        ControlAdd.groupBoxAdd(Me, gbBtn, "", iLeft, iTop, iWidth, 45)

        iTop += Me.gbBtn.Height
        iTop += 5

        btnPre = New Button
        ControlAdd.btnAdd(gbBtn, btnPre, "上一个", iLeft, 10, 80, 30)
        iLeft += btnPre.Width

        btnNext = New Button
        ControlAdd.btnAdd(gbBtn, btnNext, "下一个", iLeft, btnPre.Top, btnPre.Width, btnPre.Height)
        iLeft += btnNext.Width

        btnAnswer = New Button
        ControlAdd.btnAdd(gbBtn, btnAnswer, "答案 ", iLeft, btnPre.Top, btnPre.Width, btnPre.Height)
        iLeft += btnAnswer.Width

        btnAddSign = New Button
        ControlAdd.btnAdd(gbBtn, btnAddSign, "标记 ", iLeft, btnPre.Top, btnPre.Width, btnPre.Height)
        iLeft += btnAddSign.Width

        btnRemoveSign = New Button
        ControlAdd.btnAdd(gbBtn, btnRemoveSign, "取消标记 ", iLeft, btnPre.Top, btnPre.Width, btnPre.Height)
        iLeft += btnRemoveSign.Width

        btnRemoveAllSign = New Button
        ControlAdd.btnAdd(gbBtn, btnRemoveAllSign, "取消所有 ", iLeft, btnPre.Top, btnPre.Width, btnPre.Height)
        iLeft += btnRemoveAllSign.Width


        opRnd = New RadioButton
        ControlAdd.radioButtonAdd(gbBtn, opRnd, "随机", iLeft, 15, 50)

        iLeft += opRnd.Width
        opSort = New RadioButton
        ControlAdd.radioButtonAdd(gbBtn, opSort, "顺序", iLeft, opRnd.Top, opRnd.Width)
        With opSort
            .Checked = True
            iLeft += .Width
        End With

        nudFont = New NumericUpDown
        With nudFont
            .Maximum = 50
            .Minimum = 12
            .Value = 20
            .Left = iLeft
            .Width = 50
            .Top = 15
        End With

        gbBtn.Controls.Add(nudFont)

        With Me

            .Width = iWidth + 15
            .Height = iTop + 50
            .Text = "背题"
            '.Size = New Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)
            .Location = New Point((Screen.PrimaryScreen.Bounds.Width - Me.Width) / 2, (Screen.PrimaryScreen.Bounds.Height - Me.Height) / 2)
            .WindowState = FormWindowState.Maximized
        End With

        readShiZhi()
        getData()

        Try
            If System.IO.File.Exists(Application.StartupPath & shiZhi) Then
                Dim str As String = Func.TxtRead(Application.StartupPath & shiZhi)
                Dim arr() As String = Split(str, vbNewLine)
                iData = arr(3)
                Me.lbQuestion.Text = dt.Rows(iData)(0) & "-" & dt.Rows(iData)(2) 'dataArr(0, iData) & "-" & dataArr(2, iData)
            End If
            tiShi()
        Catch ex As Exception

        End Try
    End Sub


    Private Sub Form2_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Try
            Me.gbShiJuan.Width = Me.Width - 15
            Me.gbLeiXing.Width = Me.Width - 15
            'Me.gbSign.Width = Me.Width - 15
            Me.lbQuestion.Width = Me.Width - 15
            Me.lbAnswer.Width = Me.Width - 15
            Me.gbBtn.Width = Me.Width - 15

            Me.gbBtn.Top = Me.Height - Me.gbBtn.Height * 2

            Dim y1 As Integer = Me.lbQuestion.Top
            Dim y2 As Integer = Me.gbBtn.Top

            Me.lbQuestion.Height = (y2 - y1) / 2
            Me.lbAnswer.Height = Me.lbQuestion.Height
            Me.lbAnswer.Top = Me.lbQuestion.Top + Me.lbQuestion.Height + 5
            Me.btnAnswer.Focus()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub cb_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        saveExcel()
        getData()
        'Debug.Print(getSql)
    End Sub

    Function cbAdd(ByVal gb As System.Windows.Forms.GroupBox, ByVal cbLeft As Integer, ByVal cbTop As Integer, ByVal cbWidth As Integer, ByVal cbText As String)
        Dim cb As New System.Windows.Forms.CheckBox
        AddHandler cb.Click, AddressOf cb_Click
        With cb
            .Top = cbTop
            .Text = cbText
            .Left = cbLeft
            .Width = cbWidth
            .Checked = True
        End With
        gb.Controls.Add(cb)
        Return 0
    End Function

    Function getData()
        Dim strSql As String = getSql()

        If strSql = "" Then
            MsgBox("有条件未选取")
            Return 0
        Else
            'adoResult = Func.CreateAdoArr(strSql, Application.StartupPath & fileName, dataArr)
            Dim c_ado = New CADO(Application.StartupPath & fileName)
            c_ado.StrSql = strSql
            dt = c_ado.GetData()
            c_ado = Nothing
        End If
        'If adoResult = 0 Then
        '    Return 0
        '    Exit Function
        'End If

        'ReDim signArr(dataArr.Length / 6 - 1)
        'For i As Integer = 0 To signArr.Length - 1
        '    signArr(i) = CInt(dataArr(baoMi.sign, i))
        'Next

        Me.lbAnswer.Text = ""
        iData = 0
        Me.lbQuestion.Text = dt.Rows(iData)(0) & "-" & dt.Rows(iData)(2) 'dataArr(0, iData) & "-" & dataArr(2, iData)
        Me.btnAnswer.Focus()
        tiShi()
        Return 0
    End Function

    Function getSql() As String
        Dim strSql As String = "Select * From " & tableName & " Where 题目<>'' And "

        'Return "Select * From " & tableName & " Where 题目<>''"

        Dim iShiJuan As Integer = 0, iLeiXing As Integer = 0, iSign As Integer = 0
        Dim arrShiJuan() As System.Object = Nothing, arrLeiXing() As System.Object = Nothing, arrSign() As System.Object = Nothing

        getCbText(Me.gbShiJuan, arrShiJuan, iShiJuan)
        If iShiJuan = 0 Then
            Return ""
            Exit Function
        End If

        getCbText(Me.gbLeiXing, arrLeiXing, iLeiXing)
        If iLeiXing = 0 Then
            Return ""
            Exit Function
        End If

        getCbText(Me.gbSign, arrSign, iSign)
        If iSign = 0 Then
            Return ""
            Exit Function
        End If

        Dim str1 As String = " (试卷 In (" & Join(arrShiJuan, ",") & "))"
        Dim str2 As String = " (类型 In (" & Join(arrLeiXing, ",") & "))"
        Dim str3 As String = " (标记 In (" & Join(arrSign, ",") & "))"


        strSql = strSql & str1 & " And" & str2 & " And" & str3
        strSql = strSql & " Order By [Row]"


        Return strSql

    End Function

    Function getCbText(ByVal gb As System.Windows.Forms.GroupBox, ByRef Arr() As System.Object, ByRef k As Integer)
        Dim iCount As Integer = 1

        For Each c As Control In gb.Controls
            If CType(c, System.Windows.Forms.CheckBox).Checked Then
                ReDim Preserve Arr(k)
                Arr(k) = iCount
                k += 1
            End If
            iCount += 1
        Next

        Return 0
    End Function

    Private Sub gbLeiXing_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles gbLeiXing.Click, gbShiJuan.Click, gbSign.Click
        For Each c As Control In sender.controls
            CType(c, System.Windows.Forms.CheckBox).Checked = True
        Next
        saveExcel()
        getData()
    End Sub

    Private Sub gbLeiXing_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles gbLeiXing.DoubleClick, gbShiJuan.DoubleClick, gbSign.DoubleClick
        For Each c As Control In sender.controls
            CType(c, System.Windows.Forms.CheckBox).Checked = False
        Next
        saveExcel()
        getData()
    End Sub

    Private Sub btnNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNext.Click
        Try
            If adoResult = 0 Then
                Exit Sub
            End If

            Me.lbAnswer.Text = ""

            If Me.opRnd.Checked Then
                iData = Int(Rnd() * dt.Rows.Count) ' dataArr.Length / 6)
            Else
                If iData < dt.Rows.Count - 1 Then
                    iData += 1
                Else
                    iData = 0
                End If
            End If

            If iData <= dt.Rows.Count - 1 Then
                Me.lbQuestion.Text = dt.Rows(iData)(0) & "-" & dt.Rows(iData)(2) 'dataArr(0, iData) & "-" & dataArr(2, iData)
            End If

            Me.btnAnswer.Focus()
            tiShi()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnAnswer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAnswer.Click
        Try
            If adoResult = 0 Then
                Exit Sub
            End If

            Me.lbAnswer.Text = dt.Rows(iData)(1).ToString 'dataArr(1, iData).ToString
            Me.btnNext.Focus()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnPre_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPre.Click
        Try
            If adoResult = 0 Then
                Exit Sub
            End If

            Me.lbAnswer.Text = ""

            If Me.opRnd.Checked Then
                iData = Int(Rnd() * dt.Rows.Count)
            Else
                If iData > 0 Then
                    iData -= 1
                Else
                    iData = dt.Rows.Count - 1
                End If
            End If

            If iData > 1 Then
                Me.lbQuestion.Text = dt.Rows(iData)(0) & "-" & dt.Rows(iData)(2) ' dataArr(0, iData) & "-" & dataArr(2, iData)
            End If

            Me.btnAnswer.Focus()
            tiShi()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnAddSign_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddSign.Click, cmsAddSign.Click
        Try
            'dataArr(baoMi.sign, iData) = 2
            dt.Rows(iData)(baoMi.sign) = 2
            Me.btnAnswer.Focus()
            tiShi()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnRemoveSign_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRemoveSign.Click, cmsRemoveSign.Click
        Try
            'dataArr(4, iData) = 1
            dt.Rows(iData)(4) = 1
            Me.btnAnswer.Focus()
            tiShi()
        Catch ex As Exception

        End Try

    End Sub

    Function tiShi()
        Dim str As String = "当前题目是"
        Select Case dt.Rows(iData)(3)' dataArr(3, iData)
            Case 1
                str = str & "填空题，"
            Case 2
                str = str & "判断题，"
            Case 3
                str = str & "单项选择题，"
            Case 4
                str = str & "多项选择题，"
            Case 5
                str = str & "简答题，"
        End Select

        If dt.Rows(iData)(4) = 1 Then ' dataArr(4, iData) = 1 Then
            str = str & "未标记。"
        Else
            str = str & "已标记。"
        End If

        str = str & "共有" & dt.Rows.Count & "个题目，当前第" & iData + 1 & "个。"

        Me.lbTiShi.Text = str
        Return 0
    End Function

    Private Sub nudFont_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles nudFont.ValueChanged
        Me.lbQuestion.Font = New System.Drawing.Font("宋体", Me.nudFont.Value, System.Drawing.FontStyle.Bold)
        Me.lbAnswer.Font = New System.Drawing.Font("宋体", Me.nudFont.Value, System.Drawing.FontStyle.Bold)

    End Sub

    Function saveExcel()
        'Dim AdoConn As Object = Nothing, strSql As String

        'Try
        '    AdoConn = CreateObject("ADODB.Connection")
        '    AdoConn.Open(Func.ExcelData(Application.StartupPath & fileName))

        '    For i As Integer = 0 To signArr.Length - 1
        '        If signArr(i) <> dataArr(baoMi.sign, i) Then
        '            strSql = "Update " & tableName & " Set 标记=" & dataArr(4, i) & " Where Row=" & dataArr(5, i)
        '            AdoConn.Execute(strSql)
        '        End If
        '    Next
        '    AdoConn.Close()
        'Catch ex As Exception
        '    'MsgBox(ex.Message)
        'Finally

        '    AdoConn = Nothing
        'End Try

        Dim c_ado = New CADO(Application.StartupPath & fileName)
        c_ado.UpdateData(dt, tableName)
        c_ado = Nothing

        Return 0
    End Function

    Function saveSheZhi()
        '1试卷、2类型、3标记、4iData
        Dim strWrite As String = ""
        Dim tempStr As String = ""

        For Each c As Control In Me.gbShiJuan.Controls
            tempStr = tempStr & "、" & CType(c, System.Windows.Forms.CheckBox).Checked.ToString
        Next

        strWrite = tempStr & vbNewLine
        tempStr = ""

        For Each c As Control In Me.gbLeiXing.Controls
            tempStr = tempStr & "、" & CType(c, System.Windows.Forms.CheckBox).Checked.ToString
        Next

        strWrite = strWrite & tempStr & vbNewLine
        tempStr = ""

        For Each c As Control In Me.gbSign.Controls
            tempStr = tempStr & "、" & CType(c, System.Windows.Forms.CheckBox).Checked.ToString
        Next

        strWrite = strWrite & tempStr & vbNewLine
        strWrite = strWrite & iData.ToString

        Func.TxtWrite(Application.StartupPath & shiZhi, strWrite)

        Return 0

    End Function

    Function readShiZhi()
        If System.IO.File.Exists(Application.StartupPath & shiZhi) Then
            Dim str As String = Func.TxtRead(Application.StartupPath & shiZhi)
            Dim arr() As String = Split(str, vbNewLine)
            Dim arr1() As String = Split(arr(0), "、")
            For i As Integer = 1 To arr1.Length - 1
                CType(Me.gbShiJuan.Controls(i - 1), System.Windows.Forms.CheckBox).Checked = arr1(i)
            Next

            Dim arr2() As String = Split(arr(1), "、")
            For i As Integer = 1 To arr2.Length - 1
                CType(Me.gbLeiXing.Controls(i - 1), System.Windows.Forms.CheckBox).Checked = arr2(i)
            Next

            Dim arr3() As String = Split(arr(2), "、")
            For i As Integer = 1 To arr3.Length - 1
                CType(Me.gbSign.Controls(i - 1), System.Windows.Forms.CheckBox).Checked = arr3(i)
            Next


        End If

        Return 0
    End Function
    Private Sub lbAnswer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbAnswer.Click, lbQuestion.Click
        SendKeys.Send("{Enter}")
    End Sub

    Private Sub btnRemoveAllSign_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRemoveAllSign.Click, cmsRemoveAllSign.Click
        Try
            For i As Integer = 0 To dt.Rows.Count - 1
                'dataArr(baoMi.sign, i) = 1
                dt.Rows(i)(baoMi.sign) = 1
            Next
            tiShi()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub msItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles msItem.Click
        FWrite.Show()
        Me.Hide()
    End Sub
End Class

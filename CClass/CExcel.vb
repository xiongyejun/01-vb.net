Imports System.Runtime.InteropServices

Public Class CExcel
    Private app_Excel As Object = Nothing

    WriteOnly Property FreezePanes() As Boolean
        Set(ByVal value As Boolean)
            app_Excel.ActiveWindow.FreezePanes = True
        End Set

    End Property
    WriteOnly Property DisplayAlerts() As Boolean
        Set(ByVal value As Boolean)
            app_Excel.DisplayAlerts = value
        End Set
    End Property
    WriteOnly Property ScreenUpdating() As Boolean
        Set(ByVal value As Boolean)
            app_Excel.ScreenUpdating = value
        End Set
    End Property

    ''' <summary> 创建Excel.Application对象 </summary>
    Function CreateExcel()
        If app_Excel Is Nothing Then
            app_Excel = CreateObject("Excel.Application")
            app_Excel.Visible = 1
        End If
        Return 1
    End Function

    ''' <summary> 获取Excel.Application对象 </summary>
    Function GetExcel()
        Try
            If app_Excel Is Nothing Then
                app_Excel = GetObject(, "Excel.Application")
                app_Excel.Visible = 1
            End If
        Catch ex As Exception

        End Try

        Return 1
    End Function

    Function GetRng() As Object
        If app_Excel Is Nothing Then
            MsgBox("没有打开Excel程序。")
            Return Nothing
        Else
            Dim rng As Object = app_Excel.InputBox("选择要导入的单元格区域", "选择单元格", Type:=8)
            If rng.ToString = "False" Then Return Nothing
            Return rng
        End If
    End Function


    ''' <summary> 打开一个工作簿 </summary>
    ''' <param name="FileName">工作簿完整路径</param>
    Function WorkBookOpen(ByVal FileName As String) As Object
        Dim wk As Object = Nothing
        Try
            wk = app_Excel.Workbooks.Open(FileName, False)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return wk
    End Function

    ''' <summary> 添加一个工作簿 </summary>
    ''' <param name="FileName">工作簿完整路径</param>
    Function WorkBookAdd(ByVal FileName As String) As Object
        Dim wk As Object = app_Excel.Workbooks.Add
        Try
            wk.SaveAs(FileName)
        Catch ex As Exception
            MsgBox(ex.Message & vbNewLine & "这通常可能是此路径下已经存在同名文件，并且打开了。", vbInformation)
            Return Nothing
        End Try

        Return wk
    End Function

    Function GetData(ByVal path As String, ByVal TableName As String)
        Dim cls_ado As New CADO(path)

        cls_ado.StrSql = "Select * From [" & TableName & "]"
        cls_ado.GetData()
        cls_ado = Nothing

        Return 1
    End Function
    ''' <summary> 获取工作簿的工作表列表 </summary>
    ''' <returns>返回工作表名称数组</returns>
    Function GetWorkBookSheet(ByVal path As String) As String()
        Dim cls_ado As New CADO(path)
        Dim str() As String = cls_ado.GetTables()
        If str Is Nothing Then Return Nothing

        '如果是Excel，自定义名称也是表，但是没有$符号，工作表名有可能包含在'符号内
        str = Array.FindAll(str, AddressOf HasSymbol)

        cls_ado = Nothing
        Return str
    End Function
    ''' <summary> 判断是否包含$符号 </summary>
    ''' <param name="str">要判断的字符串</param>
    Private Shared Function HasSymbol(ByVal str As String) As Boolean
        If Right(str, 1) = "$" Then
            Return True
        ElseIf Right(str, 2) = "$'" Then
            Return True
        Else
            Return False
        End If
    End Function

    Sub ExcelQuit()
        If app_Excel IsNot Nothing Then app_Excel.Quit()
    End Sub

    Protected Overrides Sub Finalize()
        If app_Excel IsNot Nothing Then
            Marshal.ReleaseComObject(app_Excel)
        End If

        MyBase.Finalize()
    End Sub
End Class

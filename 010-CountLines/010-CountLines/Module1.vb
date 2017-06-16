Imports System.IO
Imports System.Windows.Forms

Module Module1

    Function Main(ByVal args() As String) As Integer
        For i As Integer = 0 To UBound(args, 1)
            Console.WriteLine(args(i))
        Next

        WriteTiShi()
        Dim StrReadLine As String = Console.ReadLine()

        Do Until StrReadLine = "exit"
            WriteLineCountLines()

            WriteTiShi()
            StrReadLine = Console.ReadLine()
        Loop

        Return 0
    End Function

    ''' <summary>
    ''' 提示信息
    ''' </summary>
    Sub WriteTiShi()
        Console.ForegroundColor = ConsoleColor.Green
        Console.WriteLine("exit     退出")
        Console.WriteLine("其他     执行选择文件，计算行数")
        Console.Write("CountLines$ ")
        Console.ResetColor()
    End Sub

    ''' <summary>
    ''' 输出所有文件的名称+行号
    ''' </summary>
    Sub WriteLineCountLines()
        Dim k As Integer = 0
        Dim tmp As Integer = 0
        Dim ArrFile() As String = GetFiles()


        If ArrFile IsNot Nothing Then
            Console.WriteLine()
            Console.BackgroundColor = ConsoleColor.Red
            Console.WriteLine("{0,-60}{1,10}", "FileName", "Lines")
            Console.WriteLine("{0,70}", Strings.StrDup(70, "-"))
            For i As Integer = 0 To ArrFile.Length - 1
                tmp = CountLines(ArrFile(i))
                Console.WriteLine("{0,-60}{1,10}", ArrFile(i), tmp)
                k += tmp
            Next

            Console.WriteLine("{0,70}", Strings.StrDup(70, "="))
            Console.WriteLine("{0,-60}{1,10}", "Totle Lines", k)
            Console.WriteLine()
            Console.ResetColor()
        End If
    End Sub

    ''' <summary>
    ''' 获取某个文件的行数
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <returns></returns>
    Function CountLines(ByVal fileName As String) As Integer
        Dim sr As StreamReader = Nothing
        Dim k As Integer = 0

        Try
            sr = New StreamReader(fileName)
            Do Until sr.EndOfStream
                k += 1
                sr.ReadLine()
            Loop
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally
            Try
                sr.Close()
            Catch ex As Exception

            End Try
        End Try

        Return k
    End Function

    ''' <summary>
    ''' 选择文件
    ''' </summary>
    ''' <returns></returns>
    Function GetFiles() As String()
        Dim f As OpenFileDialog = New OpenFileDialog

        f.SupportMultiDottedExtensions = True
        f.Title = "选择文件"
        f.Multiselect = True

        If f.ShowDialog = DialogResult.OK Then
            Return f.FileNames
        Else
            Return Nothing
        End If
    End Function


End Module

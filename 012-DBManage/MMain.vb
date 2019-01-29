Module MMain
    Private F_Main As FMain

    Sub Main()
        '读取设置，并应用设置
        ReadSet()

        If DicSet.ContainsKey("DBPath") Then
            DB_Info.Path = DicSet("DBPath")
        Else
            If Not SelectDB() Then Return
        End If

        cdb = New CSQLite(SQLiteDllPath, DB_Info.Path)

        Dim ret As Integer = cdb.OpenDB()
        If ret Then
            Console.WriteLine(ret)
        Else
            Console.WriteLine("open db")
        End If

        InitDBInfo()

        Application.SetCompatibleTextRenderingDefault(False)
        Control.CheckForIllegalCrossThreadCalls = False
        F_Main = New FMain
        'Enable Windows XP's style.
        Application.EnableVisualStyles()

        Application.Run(F_Main)

        F_Main.Close()
        F_Main.Dispose()
        F_Main = Nothing


        ret = cdb.CloseDB()
        If ret Then
            Console.WriteLine(cdb.GetErr & ret)
        Else
            Console.WriteLine("close db")
        End If
        cdb = Nothing

        Console.WriteLine("sqlite")
    End Sub

End Module

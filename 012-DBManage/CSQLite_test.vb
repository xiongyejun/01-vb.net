Module CSQLite_test

    Sub TestMain()
        Dim s As CSQLite = New CSQLite("E:\00-学习资料\00-Excel学习资料\SQLite", "E:\03-数据库\PriceDB.sqlite")
        Console.WriteLine(s.Version)

        Dim ret As Integer = s.OpenDB()
        If ret Then
            Console.WriteLine(ret)
        Else
            Console.WriteLine("open db")
        End If


        Dim arr(7) As String
        arr(0) = "CREATE TABLE 产品 (ID integer not null unique, 型号 text not null primary key, 名称 text not null, 长 real check(typeof(长)='real' or typeof(长)='integer'), 宽 real check(typeof(宽)='real' or typeof(宽)='integer'), 高 real check(typeof(高)='real' or typeof(高)='integer'), 重量 real check(typeof(重量)='real' or typeof(重量)='integer'), 结构重 real check(typeof(结构重)='real' or typeof(结构重)='integer'), inserttime text not null default (datetime(CURRENT_TIMESTAMP,'localtime')), 备注 text default '')"
        arr(1) = "CREATE TABLE 厂家 (ID integer not null unique, 名称 text not null primary key, 代号 text unique default '', 省 text default '', 市 text default '', 区 text default '', 详细地址 text default '', inserttime text not null default (datetime(CURRENT_TIMESTAMP,'localtime')), 备注 text default '')"
        arr(2) = "CREATE TABLE 项目类别 (ID integer not null unique, 名称 text not null primary key, inserttime text not null default (datetime(CURRENT_TIMESTAMP,'localtime')), 备注 text default '')"
        arr(3) = "CREATE TABLE 价格形式 (ID integer not null unique, 名称 text not null primary key, inserttime text not null default (datetime(CURRENT_TIMESTAMP,'localtime')), 备注 text default '')"
        arr(4) = "CREATE TABLE 价格构成项 (ID integer not null unique, 名称 text not null primary key, inserttime text not null default (datetime(CURRENT_TIMESTAMP,'localtime')), 备注 text default '')"
        arr(5) = "CREATE TABLE 项目明细 (ID integer not null unique, 代码 text not null primary key collate nocase, 名称 text not null, 型号 text default '', 规格 text default '', 单位 text default '', 项目类别ID integer not null check(typeof(项目类别ID)='integer') references 项目类别(ID) on update cascade on delete cascade, 厂家ID integer not null check(typeof(厂家ID)='integer') references 厂家(ID) on update cascade on delete cascade, 产品ID integer not null check(typeof(产品ID)='integer') references 产品(ID) on update cascade on delete cascade, inserttime text not null default (datetime(CURRENT_TIMESTAMP,'localtime')), 备注 text default '')"
        arr(6) = "CREATE TABLE 价格时间 (ID integer not null unique, 项目明细ID integer not null check(typeof(项目明细ID)='integer') references 项目明细(ID) on update cascade on delete cascade, 时间 text not null check(length(时间)=length('2000-01-10')), 价格形式ID integer not null check(typeof(价格形式ID)='integer') references 价格形式(ID) on update cascade on delete cascade, inserttime text not null default (datetime(CURRENT_TIMESTAMP,'localtime')), 备注 text default '', primary key(项目明细ID,时间))"
        arr(7) = "CREATE TABLE 价格数据 (ID integer not null unique, 价格时间ID integer not null check(typeof(价格时间ID)='integer') references 价格时间(ID) on update cascade on delete cascade, 序号 text not null, 价格构成项ID integer not null check(typeof(价格构成项ID)='integer') references 价格构成项(ID) on update cascade on delete cascade, 价格 real not null check(typeof(价格)='real' or typeof(价格)='integer'), 单位 text not null, inserttime text not null default (datetime(CURRENT_TIMESTAMP,'localtime')), 备注 text default '', primary key(价格时间ID,序号))"
        For i As Integer = 0 To arr.Length - 1
            If Test(s, arr(i)) Then GoTo exit_sub
        Next

        's.ForeignKeys()

        'update
        'If Test(s,"update a set id=21") = RetCode.Err Then GoTo exit_sub
        'delete
        'If Test(s,"delete from a") = RetCode.Err Then GoTo exit_sub

        'Console.WriteLine(s.GetTableID(String.Format("select MAX(ID) from {0}", "项目价格")))

        's.ExecuteNonQuery("begin")
        's.ExecuteNonQuery(String.Format("delete from {0} where ID=1", "项目价格"))
        'If TestSelect(s, "pragma database_list") = RetCode.Err Then GoTo exit_sub
        's.ExecuteNonQuery("commit")
        'If TestSelect(s, String.Format("PRAGMA table_info ({0})", "a")) = RetCode.Err Then GoTo exit_sub

        'Dim ti() As CSQLite.TableInfo = s.GetTableInfo
        'For i As Integer = 0 To ti.Length - 1
        '    If TestSelect(s, Strings.Format(ti(i).Name, "PRAGMA table_info({0})")) = RetCode.Err Then GoTo exit_sub
        'Next

        'CREATE
        'If Test(s, "CREATE TABLE a (id integer not null primary key autoincrement, 姓名 text not null, d date not null check(typeof(d)='date'))") = RetCode.Err Then GoTo exit_sub
        'insert
        'Console.WriteLine("insert into a values (2,'name2','" & Format(Date.Now(), "yyyy-mm-dd") & "')")
        'If Test(s, "insert into a values (2,'name2','" & Format(Date.Now(), "yyyy-mm-dd") & "')") = RetCode.Err Then GoTo exit_sub

        'Dim stmtHandle As Integer
        'Console.WriteLine(s.Prepare("select * from a where id=(?)", stmtHandle))
        'Console.WriteLine(s.BindInt32(1, 2, stmtHandle))
        'Console.WriteLine(s.GetStep(stmtHandle))
        'Console.WriteLine(s.ColumnValue(1, s.ColumnType(1, stmtHandle), stmtHandle))
        'Console.WriteLine(s.stmtFinalize(stmtHandle))

exit_sub:

        ret = s.CloseDB()
        If ret Then
            Console.WriteLine(s.GetErr & ret)
        Else
            Console.WriteLine("close db")
        End If
        s = Nothing

        Console.Read()
    End Sub

    Function Test(ByVal s As CSQLite, ByVal strSql As String) As RetCode
        Dim ret As Integer
        ret = s.ExecuteNonQuery(strSql)
        If ret Then
            Console.WriteLine(s.GetErr)
            Return RetCode.Err
        End If

        Return RetCode.Succss
    End Function

    Function TestSelect(ByVal s As CSQLite, ByVal strsql As String) As RetCode
        Dim dt As DataTable = s.ExecuteQuery(strsql)
        If dt Is Nothing Then
            Console.WriteLine(s.GetErr)
            Return RetCode.Err
        End If

        For j As Integer = 0 To dt.Columns.Count - 1
            Console.Write(dt.Columns.Item(j))
            Console.Write(vbTab)
        Next
        Console.WriteLine()

        For i As Integer = 0 To dt.Rows.Count - 1
            For j As Integer = 0 To dt.Columns.Count - 1
                Console.Write(dt.Rows(i).Item(j))
                Console.Write(vbTab)
            Next
            Console.WriteLine()
        Next

        Return RetCode.Succss
    End Function

#Region "获取数据库表格、字段等信息"
    Function GetTableInfo(ByVal s As CSQLite) As CSQLite.TableInfo()
        Dim strSql As String = "select name from sqlite_master where type='table' and name<>'sqlite_sequence'"
        Dim dt As DataTable = s.ExecuteQuery(strSql)
        If dt Is Nothing Then Return Nothing

        Dim ret(dt.Rows.Count - 1) As CSQLite.TableInfo
        For i As Integer = 0 To ret.Length - 1
            ret(i).Name = dt.Rows(i).Item("name")
            'ret(i).Field = GetFieldInfo(s, ret(i).Name)
        Next
        dt = Nothing

        Return ret
    End Function
    ''' <summary>
    ''' 根据tableName，读取字段的信息
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function GetFieldInfo(ByVal s As CSQLite, ByVal tableName As String) As CSQLite.FieldInfo
        Dim dt As DataTable = s.ExecuteQuery(Strings.Format(tableName, "PRAGMA table_info({0})"))
        Dim fieldCount As Integer = dt.Rows.Count
        Dim pkCount As Integer = 0 '主键数量

        If fieldCount Then
            Dim fi As CSQLite.FieldInfo
            ReDim fi.Name(fieldCount - 1)
            ReDim fi.Type(fieldCount - 1)

            For i As Integer = 0 To fieldCount - 1
                fi.Name(i) = dt.Rows(i).Item("name")
                If dt.Rows(i).Item("pk") Then
                    ReDim Preserve fi.PrimaryKey(pkCount)
                    ReDim Preserve fi.PrimaryKeyIndex(pkCount)
                    fi.PrimaryKey(pkCount) = fi.Name(i)
                    fi.PrimaryKeyIndex(pkCount) = i

                    pkCount += 1
                End If

                Dim strType As String = dt.Rows(i).Item("type")
                strType = Strings.LCase(strType)
                If strType.IndexOf("char") > -1 Then
                    strType = "text"
                End If

                Select Case strType
                    Case "integer"
                        fi.Type(i) = Type.GetType("System.Int32")
                    Case "float"
                        fi.Type(i) = Type.GetType("System.Double")
                    Case "real"
                        fi.Type(i) = Type.GetType("System.Double")
                    Case "text"
                        fi.Type(i) = Type.GetType("System.String")
                    Case "blob"
                        fi.Type(i) = Type.GetType("System.Object")
                    Case ""
                        fi.Type(i) = Type.GetType("System.Object")
                    Case Else
                        fi.Type(i) = Type.GetType("System.Object")
                End Select

            Next

            Return fi
        Else
            Return Nothing
        End If
    End Function
#End Region
End Module
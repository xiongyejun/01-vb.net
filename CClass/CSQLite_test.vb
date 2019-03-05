Module CSQLite_test

    Sub TestMain()
        Dim s As CSQLite = New CSQLite("E:\00-ѧϰ����\00-Excelѧϰ����\SQLite", "E:\03-���ݿ�\PriceDB.sqlite")
        Console.WriteLine(s.Version)

        Dim ret As Integer = s.OpenDB()
        If ret Then
            Console.WriteLine(ret)
        Else
            Console.WriteLine("open db")
        End If


        Dim arr(7) As String
        arr(0) = "CREATE TABLE ��Ʒ (ID integer not null unique, �ͺ� text not null primary key, ���� text not null, �� real check(typeof(��)='real' or typeof(��)='integer'), �� real check(typeof(��)='real' or typeof(��)='integer'), �� real check(typeof(��)='real' or typeof(��)='integer'), ���� real check(typeof(����)='real' or typeof(����)='integer'), �ṹ�� real check(typeof(�ṹ��)='real' or typeof(�ṹ��)='integer'), inserttime text not null default (datetime(CURRENT_TIMESTAMP,'localtime')), ��ע text default '')"
        arr(1) = "CREATE TABLE ���� (ID integer not null unique, ���� text not null primary key, ���� text unique default '', ʡ text default '', �� text default '', �� text default '', ��ϸ��ַ text default '', inserttime text not null default (datetime(CURRENT_TIMESTAMP,'localtime')), ��ע text default '')"
        arr(2) = "CREATE TABLE ��Ŀ��� (ID integer not null unique, ���� text not null primary key, inserttime text not null default (datetime(CURRENT_TIMESTAMP,'localtime')), ��ע text default '')"
        arr(3) = "CREATE TABLE �۸���ʽ (ID integer not null unique, ���� text not null primary key, inserttime text not null default (datetime(CURRENT_TIMESTAMP,'localtime')), ��ע text default '')"
        arr(4) = "CREATE TABLE �۸񹹳��� (ID integer not null unique, ���� text not null primary key, inserttime text not null default (datetime(CURRENT_TIMESTAMP,'localtime')), ��ע text default '')"
        arr(5) = "CREATE TABLE ��Ŀ��ϸ (ID integer not null unique, ���� text not null primary key collate nocase, ���� text not null, �ͺ� text default '', ��� text default '', ��λ text default '', ��Ŀ���ID integer not null check(typeof(��Ŀ���ID)='integer') references ��Ŀ���(ID) on update cascade on delete cascade, ����ID integer not null check(typeof(����ID)='integer') references ����(ID) on update cascade on delete cascade, ��ƷID integer not null check(typeof(��ƷID)='integer') references ��Ʒ(ID) on update cascade on delete cascade, inserttime text not null default (datetime(CURRENT_TIMESTAMP,'localtime')), ��ע text default '')"
        arr(6) = "CREATE TABLE �۸�ʱ�� (ID integer not null unique, ��Ŀ��ϸID integer not null check(typeof(��Ŀ��ϸID)='integer') references ��Ŀ��ϸ(ID) on update cascade on delete cascade, ʱ�� text not null check(length(ʱ��)=length('2000-01-10')), �۸���ʽID integer not null check(typeof(�۸���ʽID)='integer') references �۸���ʽ(ID) on update cascade on delete cascade, inserttime text not null default (datetime(CURRENT_TIMESTAMP,'localtime')), ��ע text default '', primary key(��Ŀ��ϸID,ʱ��))"
        arr(7) = "CREATE TABLE �۸����� (ID integer not null unique, �۸�ʱ��ID integer not null check(typeof(�۸�ʱ��ID)='integer') references �۸�ʱ��(ID) on update cascade on delete cascade, ��� text not null, �۸񹹳���ID integer not null check(typeof(�۸񹹳���ID)='integer') references �۸񹹳���(ID) on update cascade on delete cascade, �۸� real not null check(typeof(�۸�)='real' or typeof(�۸�)='integer'), ��λ text not null, inserttime text not null default (datetime(CURRENT_TIMESTAMP,'localtime')), ��ע text default '', primary key(�۸�ʱ��ID,���))"
        For i As Integer = 0 To arr.Length - 1
            If Test(s, arr(i)) Then GoTo exit_sub
        Next

        's.ForeignKeys()

        'update
        'If Test(s,"update a set id=21") = RetCode.Err Then GoTo exit_sub
        'delete
        'If Test(s,"delete from a") = RetCode.Err Then GoTo exit_sub

        'Console.WriteLine(s.GetTableID(String.Format("select MAX(ID) from {0}", "��Ŀ�۸�")))

        's.ExecuteNonQuery("begin")
        's.ExecuteNonQuery(String.Format("delete from {0} where ID=1", "��Ŀ�۸�"))
        'If TestSelect(s, "pragma database_list") = RetCode.Err Then GoTo exit_sub
        's.ExecuteNonQuery("commit")
        'If TestSelect(s, String.Format("PRAGMA table_info ({0})", "a")) = RetCode.Err Then GoTo exit_sub

        'Dim ti() As CSQLite.TableInfo = s.GetTableInfo
        'For i As Integer = 0 To ti.Length - 1
        '    If TestSelect(s, Strings.Format(ti(i).Name, "PRAGMA table_info({0})")) = RetCode.Err Then GoTo exit_sub
        'Next

        'CREATE
        'If Test(s, "CREATE TABLE a (id integer not null primary key autoincrement, ���� text not null, d date not null check(typeof(d)='date'))") = RetCode.Err Then GoTo exit_sub
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

#Region "��ȡ���ݿ����ֶε���Ϣ"
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
    ''' ����tableName����ȡ�ֶε���Ϣ
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function GetFieldInfo(ByVal s As CSQLite, ByVal tableName As String) As CSQLite.FieldInfo
        Dim dt As DataTable = s.ExecuteQuery(Strings.Format(tableName, "PRAGMA table_info({0})"))
        Dim fieldCount As Integer = dt.Rows.Count
        Dim pkCount As Integer = 0 '��������

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
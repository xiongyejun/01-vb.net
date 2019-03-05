
Class CSQLite

#Region "Structure"
    ''' <summary>
    ''' 表的字段的信息
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure FieldInfo
        Dim Name() As String
        Dim Type() As System.Type
        Dim Pointer() As Integer '是否指向其他的table，-1就不是，xxID的就记录xx的下标

        Dim PrimaryKey() As String '主键
        Dim PrimaryKeyIndex() As Integer '主键所在的列
    End Structure
    ''' <summary>
    ''' 表展开xxID字段后的字段信息
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure ExtendFieldInfo
        Dim Tables() As String '每个字段来源的表名
        Dim FieldName() As String '每个字段在表中原来的名字
        Dim ExtendFieldName() As String '扩展后的名字= Tables & . & FieldName
        Dim ExtendFieldType() As System.Type
    End Structure
    ''' <summary>
    ''' 表信息
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure TableInfo
        Dim Name As String
        Dim Field As FieldInfo

        Dim bHasID As Boolean '是否有ID这个字段
        Dim LastID As Integer '最后1个的ID值
        Dim bUseMyIdTables As Boolean '是否有表格引用了本表的ID
        Dim UseMyIdTables() As Integer '有哪些表格引用了本表的ID
        Dim ExtendField As ExtendFieldInfo '把xxID的字段扩展出来
        Dim SqlExtend As String   '扩展的sql
    End Structure
#End Region

#Region "Const"
    ' Notes:
    ' Microsoft uses UTF-16, little endian byte order.
    Private Const JULIANDAY_OFFSET As Double = 2415018.5

    ' Returned from SQLite3Initialize
    Private Const SQLITE_INIT_OK As Integer = 0
    Private Const SQLITE_INIT_ERROR As Integer = 1

    ' SQLite data types
    Private Const SQLITE_INTEGER As Integer = 1
    Private Const SQLITE_FLOAT As Integer = 2
    Private Const SQLITE_TEXT As Integer = 3
    Private Const SQLITE_BLOB As Integer = 4
    Private Const SQLITE_NULL As Integer = 5

    ' SQLite atandard return value
    Private Const SQLITE_OK As Integer = 0   ' Successful result
    Private Const SQLITE_ERROR As Integer = 1   ' SQL error or missing database
    Private Const SQLITE_INTERNAL As Integer = 2   ' Internal logic error in SQLite
    Private Const SQLITE_PERM As Integer = 3   ' Access permission denied
    Private Const SQLITE_ABORT As Integer = 4   ' Callback routine requested an abort
    Private Const SQLITE_BUSY As Integer = 5   ' The database file is locked
    Private Const SQLITE_LOCKED As Integer = 6   ' A table in the database is locked
    Private Const SQLITE_NOMEM As Integer = 7   ' A malloc() failed
    Private Const SQLITE_READONLY As Integer = 8   ' Attempt to write a readonly database
    Private Const SQLITE_INTERRUPT As Integer = 9   ' Operation terminated by sqlite3_interrupt()
    Private Const SQLITE_IOERR As Integer = 10   ' Some kind of disk I/O error occurred
    Private Const SQLITE_CORRUPT As Integer = 11   ' The database disk image is malformed
    Private Const SQLITE_NOTFOUND As Integer = 12   ' NOT USED. Table or record not found
    Private Const SQLITE_FULL As Integer = 13   ' Insertion failed because database is full
    Private Const SQLITE_CANTOPEN As Integer = 14   ' Unable to open the database file
    Private Const SQLITE_PROTOCOL As Integer = 15   ' NOT USED. Database lock protocol error
    Private Const SQLITE_EMPTY As Integer = 16   ' Database is empty
    Private Const SQLITE_SCHEMA As Integer = 17   ' The database schema changed
    Private Const SQLITE_TOOBIG As Integer = 18   ' String or BLOB exceeds size limit
    Private Const SQLITE_CONSTRAINT As Integer = 19   ' Abort due to constraint violation
    Private Const SQLITE_MISMATCH As Integer = 20   ' Data type mismatch
    Private Const SQLITE_MISUSE As Integer = 21   ' Library used incorrectly
    Private Const SQLITE_NOLFS As Integer = 22   ' Uses OS features not supported on host
    Private Const SQLITE_AUTH As Integer = 23   ' Authorization denied
    Private Const SQLITE_FORMAT As Integer = 24   ' Auxiliary database format error
    Private Const SQLITE_RANGE As Integer = 25   ' 2nd parameter to sqlite3_bind out of range
    Private Const SQLITE_NOTADB As Integer = 26   ' File opened that is not a database file
    Private Const SQLITE_ROW As Integer = 100  ' sqlite3_GetStep() has another row ready
    Private Const SQLITE_DONE As Integer = 101  ' sqlite3_GetStep() has finished executing

    ' Extended error codes
    Private Const SQLITE_IOERR_READ As Integer = 266  '(SQLITE_IOERR | (1<<8))
    Private Const SQLITE_IOERR_SHORT_READ As Integer = 522  '(SQLITE_IOERR | (2<<8))
    Private Const SQLITE_IOERR_WRITE As Integer = 778  '(SQLITE_IOERR | (3<<8))
    Private Const SQLITE_IOERR_FSYNC As Integer = 1034 '(SQLITE_IOERR | (4<<8))
    Private Const SQLITE_IOERR_DIR_FSYNC As Integer = 1290 '(SQLITE_IOERR | (5<<8))
    Private Const SQLITE_IOERR_TRUNCATE As Integer = 1546 '(SQLITE_IOERR | (6<<8))
    Private Const SQLITE_IOERR_FSTAT As Integer = 1802 '(SQLITE_IOERR | (7<<8))
    Private Const SQLITE_IOERR_UNLOCK As Integer = 2058 '(SQLITE_IOERR | (8<<8))
    Private Const SQLITE_IOERR_RDLOCK As Integer = 2314 '(SQLITE_IOERR | (9<<8))
    Private Const SQLITE_IOERR_DELETE As Integer = 2570 '(SQLITE_IOERR | (10<<8))
    Private Const SQLITE_IOERR_BLOCKED As Integer = 2826 '(SQLITE_IOERR | (11<<8))
    Private Const SQLITE_IOERR_NOMEM As Integer = 3082 '(SQLITE_IOERR | (12<<8))
    Private Const SQLITE_IOERR_ACCESS As Integer = 3338 '(SQLITE_IOERR | (13<<8))
    Private Const SQLITE_IOERR_CHECKRESERVEDLOCK As Integer = 3594 '(SQLITE_IOERR | (14<<8))
    Private Const SQLITE_IOERR_LOCK As Integer = 3850 '(SQLITE_IOERR | (15<<8))
    Private Const SQLITE_IOERR_CLOSE As Integer = 4106 '(SQLITE_IOERR | (16<<8))
    Private Const SQLITE_IOERR_DIR_CLOSE As Integer = 4362 '(SQLITE_IOERR | (17<<8))
    Private Const SQLITE_LOCKED_SHAREDCACHE As Integer = 265  '(SQLITE_LOCKED | (1<<8) )

    ' Flags For File Open Operations
    Private Const SQLITE_OPEN_READONLY As Integer = 1       ' Ok for sqlite3_open_v2()
    Public Const OPEN_READWRITE As Integer = 2       ' Ok for sqlite3_open_v2()
    Private Const SQLITE_OPEN_CREATE As Integer = 4       ' Ok for sqlite3_open_v2()
    Private Const SQLITE_OPEN_DELETEONCLOSE As Integer = 8       ' VFS only
    Private Const SQLITE_OPEN_EXCLUSIVE As Integer = 16      ' VFS only
    Private Const SQLITE_OPEN_AUTOPROXY As Integer = 32      ' VFS only
    Private Const SQLITE_OPEN_URI As Integer = 64      ' Ok for sqlite3_open_v2()
    Private Const SQLITE_OPEN_MEMORY As Integer = 128     ' Ok for sqlite3_open_v2()
    Private Const SQLITE_OPEN_MAIN_DB As Integer = 256     ' VFS only
    Private Const SQLITE_OPEN_TEMP_DB As Integer = 512     ' VFS only
    Private Const SQLITE_OPEN_TRANSIENT_DB As Integer = 1024    ' VFS only
    Private Const SQLITE_OPEN_MAIN_JOURNAL As Integer = 2048    ' VFS only
    Private Const SQLITE_OPEN_TEMP_JOURNAL As Integer = 4096    ' VFS only
    Private Const SQLITE_OPEN_SUBJOURNAL As Integer = 8192    ' VFS only
    Private Const SQLITE_OPEN_MASTER_JOURNAL As Integer = 16384   ' VFS only
    Private Const SQLITE_OPEN_NOMUTEX As Integer = 32768   ' Ok for sqlite3_open_v2()
    Private Const SQLITE_OPEN_FULLMUTEX As Integer = 65536   ' Ok for sqlite3_open_v2()
    Private Const SQLITE_OPEN_SHAREDCACHE As Integer = 131072  ' Ok for sqlite3_open_v2()
    Private Const SQLITE_OPEN_PRIVATECACHE As Integer = 262144  ' Ok for sqlite3_open_v2()
    Private Const SQLITE_OPEN_WAL As Integer = 524288  ' VFS only

    ' Options for Text and Blob binding
    Private Const SQLITE_STATIC As Integer = 0
    Private Const SQLITE_TRANSIENT As Integer = -1

    ' System calls
    Private Const CP_UTF8 As Integer = 65001

#End Region

#Region "API"
#If Win64 Then

Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Integer, ByVal dwFlags As Integer, ByVal lpMultiByteStr As IntegerPtr, ByVal cbMultiByte As Integer, ByVal lpWideCharStr As IntegerPtr, ByVal cchWideChar As Integer) As Integer
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Integer, ByVal dwFlags As Integer, ByVal lpWideCharStr As IntegerPtr, ByVal cchWideChar As Integer, ByVal lpMultiByteStr As IntegerPtr, ByVal cbMultiByte As Integer, ByVal lpDefaultChar As IntegerPtr, ByVal lpUsedDefaultChar As IntegerPtr) As Integer
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As IntegerPtr, ByVal pSource As IntegerPtr, ByVal length As Integer)
Private Declare PtrSafe Function lstrcpynW Lib "kernel32" (ByVal pwsDest As IntegerPtr, ByVal pwsSource As IntegerPtr, ByVal cchCount As Integer) As IntegerPtr
Private Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal pwsDest As IntegerPtr, ByVal pwsSource As IntegerPtr) As IntegerPtr
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal pwsString As IntegerPtr) As Integer
Private Declare PtrSafe Function SysAllocString Lib "OleAut32" (ByRef pwsString As IntegerPtr) As IntegerPtr
Private Declare PtrSafe Function SysStringLen Lib "OleAut32" (ByVal bstrString As IntegerPtr) As Integer
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As IntegerPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As IntegerPtr) As Integer
#Else
    Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Integer, ByVal dwFlags As Integer, ByVal lpMultiByteStr As Integer, ByVal cbMultiByte As Integer, ByVal lpWideCharStr As Byte(), ByVal cchWideChar As Integer) As Integer
    Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Integer, ByVal dwFlags As Integer, ByVal lpWideCharStr As String, ByVal cchWideChar As Integer, ByVal lpMultiByteStr As Integer, ByVal cbMultiByte As Integer, ByVal lpDefaultChar As Integer, ByVal lpUsedDefaultChar As Integer) As Integer
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As Integer, ByVal pSource As Integer, ByVal length As Integer)
    Private Declare Function lstrcpynW Lib "kernel32" (ByVal pwsDest As Integer, ByVal pwsSource As Integer, ByVal cchCount As Integer) As Integer
    Private Declare Function lstrcpyW Lib "kernel32" (ByVal pwsDest As Integer, ByVal pwsSource As Integer) As Integer
    Private Declare Function lstrlenW Lib "kernel32" (ByVal pwsString As Integer) As Integer
    Private Declare Function SysAllocString Lib "OleAut32" (ByRef pwsString As Integer) As Integer
    Private Declare Function SysStringLen Lib "OleAut32" (ByVal bstrString As Integer) As Integer
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Integer
    Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Integer) As Integer
#End If
    '=====================================================================================
    ' SQLite StdCall Imports
    '-----------------------
#If Win64 Then
' SQLite library version
Private Declare PtrSafe Function sqlite3_libversion Lib "SQLite3" () As IntegerPtr ' PtrUtf8String
' Database connections
Private Declare PtrSafe Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As IntegerPtr, ByRef hDb As IntegerPtr) As Integer
Private Declare PtrSafe Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As IntegerPtr, ByRef hDb As IntegerPtr, ByVal iFlags As Integer, ByVal zVfs As IntegerPtr) As Integer ' PtrDb
Private Declare PtrSafe Function sqlite3_close Lib "SQLite3" (ByVal hDb As IntegerPtr) As Integer
' Database connection error info
Private Declare PtrSafe Function sqlite3_errmsg Lib "SQLite3" (ByVal hDb As IntegerPtr) As IntegerPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_errmsg16 Lib "SQLite3" (ByVal hDb As IntegerPtr) As IntegerPtr ' PtrUtf16String
Private Declare PtrSafe Function sqlite3_errcode Lib "SQLite3" (ByVal hDb As IntegerPtr) As Integer
Private Declare PtrSafe Function sqlite3_extended_errcode Lib "SQLite3" (ByVal hDb As IntegerPtr) As Integer
' Database connection change counts
Private Declare PtrSafe Function sqlite3_changes Lib "SQLite3" (ByVal hDb As IntegerPtr) As Integer
Private Declare PtrSafe Function sqlite3_total_changes Lib "SQLite3" (ByVal hDb As IntegerPtr) As Integer

' Statements
Private Declare PtrSafe Function sqlite3_prepare16_v2 Lib "SQLite3" _
    (ByVal hDb As IntegerPtr, ByVal pwsSql As IntegerPtr, ByVal nSqlLength As Integer, ByRef hStmt As IntegerPtr, ByVal ppwsTailOut As IntegerPtr) As Integer
Private Declare PtrSafe Function sqlite3_step Lib "SQLite3" (ByVal hStmt As IntegerPtr) As Integer
Private Declare PtrSafe Function sqlite3_reset Lib "SQLite3" (ByVal hStmt As IntegerPtr) As Integer
Private Declare PtrSafe Function sqlite3_finalize Lib "SQLite3" (ByVal hStmt As IntegerPtr) As Integer

' Statement column access (0-based indices)
Private Declare PtrSafe Function sqlite3_column_count Lib "SQLite3" (ByVal hStmt As IntegerPtr) As Integer
Private Declare PtrSafe Function sqlite3_column_type Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal iCol As Integer) As Integer
Private Declare PtrSafe Function sqlite3_column_name Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal iCol As Integer) As IntegerPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_name16 Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal iCol As Integer) As IntegerPtr ' PtrWString

Private Declare PtrSafe Function sqlite3_column_blob Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal iCol As Integer) As IntegerPtr ' PtrData
Private Declare PtrSafe Function sqlite3_column_bytes Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal iCol As Integer) As Integer
Private Declare PtrSafe Function sqlite3_column_bytes16 Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal iCol As Integer) As Integer
Private Declare PtrSafe Function sqlite3_column_double Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal iCol As Integer) As Double
Private Declare PtrSafe Function sqlite3_column_int Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal iCol As Integer) As Integer
Private Declare PtrSafe Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal iCol As Integer) As IntegerLong
Private Declare PtrSafe Function sqlite3_column_text Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal iCol As Integer) As IntegerPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_text16 Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal iCol As Integer) As IntegerPtr ' PtrWString
Private Declare PtrSafe Function sqlite3_column_value Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal iCol As Integer) As IntegerPtr ' PtrSqlite3Value

' Statement parameter binding (1-based indices!)
Private Declare PtrSafe Function sqlite3_bind_parameter_count Lib "SQLite3" (ByVal hStmt As IntegerPtr) As Integer
Private Declare PtrSafe Function sqlite3_bind_parameter_name Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal paramIndex As Integer) As IntegerPtr
Private Declare PtrSafe Function sqlite3_bind_parameter_index Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal paramName As IntegerPtr) As Integer
Private Declare PtrSafe Function sqlite3_bind_null Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal paramIndex As Integer) As Integer
Private Declare PtrSafe Function sqlite3_bind_blob Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal paramIndex As Integer, ByVal pValue As IntegerPtr, ByVal nBytes As Integer, ByVal pfDelete As IntegerPtr) As Integer
Private Declare PtrSafe Function sqlite3_bind_zeroblob Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal paramIndex As Integer, ByVal nBytes As Integer) As Integer
Private Declare PtrSafe Function sqlite3_bind_double Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal paramIndex As Integer, ByVal Value As Double) As Integer
Private Declare PtrSafe Function sqlite3_bind_int Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal paramIndex As Integer, ByVal Value As Integer) As Integer
Private Declare PtrSafe Function sqlite3_bind_int64 Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal paramIndex As Integer, ByVal Value As IntegerLong) As Integer
Private Declare PtrSafe Function sqlite3_bind_text Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal paramIndex As Integer, ByVal psValue As IntegerPtr, ByVal nBytes As Integer, ByVal pfDelete As IntegerPtr) As Integer
Private Declare PtrSafe Function sqlite3_bind_text16 Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal paramIndex As Integer, ByVal pswValue As IntegerPtr, ByVal nBytes As Integer, ByVal pfDelete As IntegerPtr) As Integer
Private Declare PtrSafe Function sqlite3_bind_value Lib "SQLite3" (ByVal hStmt As IntegerPtr, ByVal paramIndex As Integer, ByVal pSqlite3Value As IntegerPtr) As Integer
Private Declare PtrSafe Function sqlite3_clear_bindings Lib "SQLite3" (ByVal hStmt As IntegerPtr) As Integer

'Backup
Private Declare PtrSafe Function sqlite3_sleep Lib "SQLite3" (ByVal msToSleep As Integer) As Integer
Private Declare PtrSafe Function sqlite3_backup_init Lib "SQLite3" (ByVal hDbDest As IntegerPtr, ByVal zDestName As IntegerPtr, ByVal hDbSource As IntegerPtr, ByVal zSourceName As IntegerPtr) As Integer
Private Declare PtrSafe Function sqlite3_backup_step Lib "SQLite3" (ByVal hBackup As IntegerPtr, ByVal nPage As Integer) As Integer
Private Declare PtrSafe Function sqlite3_backup_finish Lib "SQLite3" (ByVal hBackup As IntegerPtr) As Integer
Private Declare PtrSafe Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As IntegerPtr) As Integer
Private Declare PtrSafe Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As IntegerPtr) As Integer
#Else

    ' SQLite library version
    Private Declare Function sqlite3_libversion Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_libversion@0" () As Integer ' PtrUtf8String
    ' Database connections
    Private Declare Function sqlite3_open16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_open16@8" (ByVal pwsFileName As Byte(), ByRef hDb As Integer) As Integer ' PtrDb
    Private Declare Function sqlite3_open_v2 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_open_v2@16" (ByVal pwsFileName As Byte(), ByRef hDb As Integer, ByVal iFlags As Integer, ByVal zVfs As Integer) As Integer ' PtrDb
    Private Declare Function sqlite3_close Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_close@4" (ByVal hDb As Integer) As Integer
    ' Database connection error info
    Private Declare Function sqlite3_errmsg Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_errmsg@4" (ByVal hDb As Integer) As Integer ' PtrUtf8String
    Private Declare Function sqlite3_errmsg16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_errmsg16@4" (ByVal hDb As Integer) As Integer ' PtrUtf16String
    Private Declare Function sqlite3_errcode Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_errcode@4" (ByVal hDb As Integer) As Integer
    Private Declare Function sqlite3_extended_errcode Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_extended_errcode@4" (ByVal hDb As Integer) As Integer
    ' Database connection change counts
    Private Declare Function sqlite3_changes Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_changes@4" (ByVal hDb As Integer) As Integer
    Private Declare Function sqlite3_total_changes Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_total_changes@4" (ByVal hDb As Integer) As Integer

    ' Statements
    Private Declare Function sqlite3_prepare16_v2 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_prepare16_v2@20" _
        (ByVal hDb As Integer, ByVal pwsSql As Byte(), ByVal nSqlLength As Integer, ByRef hStmt As Integer, ByVal ppwsTailOut As Integer) As Integer
    Private Declare Function sqlite3_step Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_step@4" (ByVal hStmt As Integer) As Integer
    Private Declare Function sqlite3_reset Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_reset@4" (ByVal hStmt As Integer) As Integer
    Private Declare Function sqlite3_finalize Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_finalize@4" (ByVal hStmt As Integer) As Integer

    ' Statement column access (0-based indices)
    Private Declare Function sqlite3_column_count Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_count@4" (ByVal hStmt As Integer) As Integer
    Private Declare Function sqlite3_column_type Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_type@8" (ByVal hStmt As Integer, ByVal iCol As Integer) As Integer
    Private Declare Function sqlite3_column_name Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_name@8" (ByVal hStmt As Integer, ByVal iCol As Integer) As Integer ' PtrString
    Private Declare Function sqlite3_column_name16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_name16@8" (ByVal hStmt As Integer, ByVal iCol As Integer) As Integer ' PtrWString

    Private Declare Function sqlite3_column_blob Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_blob@8" (ByVal hStmt As Integer, ByVal iCol As Integer) As Integer ' PtrData
    Private Declare Function sqlite3_column_bytes Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_bytes@8" (ByVal hStmt As Integer, ByVal iCol As Integer) As Integer
    Private Declare Function sqlite3_column_bytes16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_bytes16@8" (ByVal hStmt As Integer, ByVal iCol As Integer) As Integer
    Private Declare Function sqlite3_column_double Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_double@8" (ByVal hStmt As Integer, ByVal iCol As Integer) As Double
    Private Declare Function sqlite3_column_int Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_int@8" (ByVal hStmt As Integer, ByVal iCol As Integer) As Integer
    Private Declare Function sqlite3_column_int64 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_int64@8" (ByVal hStmt As Integer, ByVal iCol As Integer) As Long ' Currency ' UNTESTED ....?
    Private Declare Function sqlite3_column_text Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_text@8" (ByVal hStmt As Integer, ByVal iCol As Integer) As Integer ' PtrString
    Private Declare Function sqlite3_column_text16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_text16@8" (ByVal hStmt As Integer, ByVal iCol As Integer) As Integer ' PtrWString
    Private Declare Function sqlite3_column_value Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_value@8" (ByVal hStmt As Integer, ByVal iCol As Integer) As Integer ' PtrSqlite3Value

    ' Statement parameter binding (1-based indices!)
    Private Declare Function sqlite3_bind_parameter_count Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_parameter_count@4" (ByVal hStmt As Integer) As Integer
    Private Declare Function sqlite3_bind_parameter_name Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_parameter_name@8" (ByVal hStmt As Integer, ByVal paramIndex As Integer) As Integer
    Private Declare Function sqlite3_bind_parameter_index Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_parameter_index@8" (ByVal hStmt As Integer, ByVal paramName As Integer) As Integer
    Private Declare Function sqlite3_bind_null Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_null@8" (ByVal hStmt As Integer, ByVal paramIndex As Integer) As Integer
    Private Declare Function sqlite3_bind_blob Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_blob@20" (ByVal hStmt As Integer, ByVal paramIndex As Integer, ByVal pValue As Integer, ByVal nBytes As Integer, ByVal pfDelete As Integer) As Integer
    Private Declare Function sqlite3_bind_zeroblob Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_zeroblob@12" (ByVal hStmt As Integer, ByVal paramIndex As Integer, ByVal nBytes As Integer) As Integer
    Private Declare Function sqlite3_bind_double Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_double@16" (ByVal hStmt As Integer, ByVal paramIndex As Integer, ByVal Value As Double) As Integer
    Private Declare Function sqlite3_bind_int Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_int@12" (ByVal hStmt As Integer, ByVal paramIndex As Integer, ByVal Value As Integer) As Integer
    Private Declare Function sqlite3_bind_int64 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_int64@16" (ByVal hStmt As Integer, ByVal paramIndex As Integer, ByVal Value As Long) As Integer 'Currency) As Integer ' UNTESTED ....?
    Private Declare Function sqlite3_bind_text Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_text@20" (ByVal hStmt As Integer, ByVal paramIndex As Integer, ByVal psValue As Integer, ByVal nBytes As Integer, ByVal pfDelete As Integer) As Integer
    Private Declare Function sqlite3_bind_text16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_text16@20" (ByVal hStmt As Integer, ByVal paramIndex As Integer, ByVal pswValue() As Byte, ByVal nBytes As Integer, ByVal pfDelete As Integer) As Integer
    Private Declare Function sqlite3_bind_value Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_value@12" (ByVal hStmt As Integer, ByVal paramIndex As Integer, ByVal pSqlite3Value As Integer) As Integer
    Private Declare Function sqlite3_clear_bindings Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_clear_bindings@4" (ByVal hStmt As Integer) As Integer

    'Backup
    Private Declare Function sqlite3_sleep Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_sleep@4" (ByVal msToSleep As Integer) As Integer
    Private Declare Function sqlite3_backup_init Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_init@16" (ByVal hDbDest As Integer, ByVal zDestName As Integer, ByVal hDbSource As Integer, ByVal zSourceName As Integer) As Integer
    Private Declare Function sqlite3_backup_step Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_step@8" (ByVal hBackup As Integer, ByVal nPage As Integer) As Integer
    Private Declare Function sqlite3_backup_finish Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_finish@4" (ByVal hBackup As Integer) As Integer
    Private Declare Function sqlite3_backup_remaining Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_remaining@4" (ByVal hBackup As Integer) As Integer
    Private Declare Function sqlite3_backup_pagecount Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_pagecount@4" (ByVal hBackup As Integer) As Integer
#End If

#End Region

#Region "Initialize - load libraries explicitly"
#If Win64 Then
Private hSQLiteLibrary As IntegerPtr
Private hSQLiteStdCallLibrary As IntegerPtr
Private DbHandle As IntegerPtr     '数据库handle
#Else
    Private hSQLiteLibrary As Integer
    Private hSQLiteStdCallLibrary As Integer
    Private DbHandle As Integer
#End If

    Private DBName As String
#End Region

#Region "Property"
    Private stateValue As Integer '记录函数的返回值，用来判断是否出错
    ReadOnly Property GetState() As Integer
        Get
            Return stateValue
        End Get
    End Property

    ReadOnly Property GetErr() As String
        Get
            Return ErrMsg()
        End Get
    End Property

    ' SQLite library version
    ReadOnly Property Version() As String
        Get
            Return Utf8PtrToString(sqlite3_libversion())
        End Get
    End Property
#End Region

#Region "New & Finalize"
    ''' <summary>
    ''' 初始化--打开dll
    ''' </summary>
    ''' <param name="libDir">dll的路径</param>
    ''' <param name="dbPath">数据库完整路径</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal libDir As String, ByVal dbPath As String)
        ' A nice option here is to call SetDllDirectory, but that API is only available since Windows XP SP1.
        If Right(libDir, 1) <> "\" Then libDir = libDir & "\"

        If hSQLiteLibrary = 0 Then
            hSQLiteLibrary = LoadLibrary(libDir + "SQLite3.dll")
            If hSQLiteLibrary = 0 Then
                MsgBox("SQLite3Initialize Error Loading " + libDir + "SQLite3.dll:", Err.LastDllError)
                stateValue = SQLITE_INIT_ERROR
                Exit Sub
            End If
        End If

#If Win64 Then
#Else
        If hSQLiteStdCallLibrary = 0 Then
            hSQLiteStdCallLibrary = LoadLibrary(libDir + "SQLite3_StdCall.dll")
            If hSQLiteStdCallLibrary = 0 Then
                MsgBox("SQLite3Initialize Error Loading " + libDir + "SQLite3_StdCall.dll:", Err.LastDllError)
                stateValue = SQLITE_INIT_ERROR
                Exit Sub
            End If
        End If
#End If
        stateValue = SQLITE_INIT_OK

        DBName = dbPath
    End Sub

    Protected Overrides Sub Finalize()
        SQLite3Free()
    End Sub

    Private Sub SQLite3Free()
        If hSQLiteStdCallLibrary <> 0 Then
            stateValue = FreeLibrary(hSQLiteStdCallLibrary)
            hSQLiteStdCallLibrary = 0
            If stateValue = 0 Then
                MsgBox("SQLite3Free Error Freeing SQLite3_StdCall.dll:", stateValue, Err.LastDllError)
            End If
        End If
        If hSQLiteLibrary <> 0 Then
            stateValue = FreeLibrary(hSQLiteLibrary)
            hSQLiteLibrary = 0
            If stateValue = 0 Then
                MsgBox("SQLite3Free Error Freeing SQLite3.dll:", stateValue, Err.LastDllError)
            End If
        End If
    End Sub
#End Region

#Region "打开关闭Database connections（如果DBName不存在，程序会自动新建1个）"
    Public Function OpenDB() As Integer
        OpenDB = sqlite3_open16(StrToByte16(DBName), DbHandle)
    End Function

    Public Function SQLite3OpenV2(ByVal flags As Integer, ByVal vfsName As String) As Integer
        Dim bufFileName() As Byte
        Dim bufVfsName() As Byte
        bufFileName = StringToUtf8Bytes(DBName)
        If vfsName = vbEmpty Then
            SQLite3OpenV2 = sqlite3_open_v2(bufFileName, DbHandle, flags, 0)
        Else
            bufVfsName = StringToUtf8Bytes(vfsName)
            SQLite3OpenV2 = sqlite3_open_v2(bufFileName, DbHandle, flags, bufVfsName(0))
        End If
    End Function
    ' 关闭 Database connections
    Public Function CloseDB() As Integer
        CloseDB = sqlite3_close(DbHandle)
    End Function
#End Region

#Region "Statement column access (0-based indices)"
    Public Function ColumnCount(ByVal stmtHandle As Integer) As Integer
        ColumnCount = sqlite3_column_count(stmtHandle)
    End Function
    Public Function ColumnType(ByVal ZeroBasedColIndex As Integer, ByVal stmtHandle As Integer) As Integer
        ColumnType = sqlite3_column_type(stmtHandle, ZeroBasedColIndex)
    End Function
    Public Function ColumnName(ByVal ZeroBasedColIndex As Integer, ByVal stmtHandle As Integer) As String
        ColumnName = Utf8PtrToString(sqlite3_column_name(stmtHandle, ZeroBasedColIndex))
        '去掉最后面的“\0”
        Return ColumnName.Substring(0, ColumnName.Length - 1)
    End Function
    Public Function ColumnDouble(ByVal ZeroBasedColIndex As Integer, ByVal stmtHandle As Integer) As Double
        ColumnDouble = sqlite3_column_double(stmtHandle, ZeroBasedColIndex)
    End Function
    Public Function ColumnInt32(ByVal ZeroBasedColIndex As Integer, ByVal stmtHandle As Integer) As Integer
        ColumnInt32 = sqlite3_column_int(stmtHandle, ZeroBasedColIndex)
    End Function
    Public Function ColumnText(ByVal ZeroBasedColIndex As Integer, ByVal stmtHandle As Integer) As String
        ColumnText = Utf8PtrToString(sqlite3_column_text(stmtHandle, ZeroBasedColIndex))
        '去掉最后面的“\0”
        Return ColumnText.Substring(0, ColumnText.Length - 1)
    End Function
    Public Function ColumnDate(ByVal ZeroBasedColIndex As Integer, ByVal stmtHandle As Integer) As Date
        ColumnDate = FromJulianDay(sqlite3_column_double(stmtHandle, ZeroBasedColIndex))
    End Function
    Public Function ColumnBlob(ByVal ZeroBasedColIndex As Integer, ByVal stmtHandle As Integer) As Byte()
#If Win64 Then
    Dim ptr As IntegerPtr
#Else
        Dim ptr As Integer
#End If
        Dim length As Integer
        Dim buf() As Byte

        ptr = sqlite3_column_blob(stmtHandle, ZeroBasedColIndex)
        length = sqlite3_column_bytes(stmtHandle, ZeroBasedColIndex)
        ReDim buf(length - 1)
        RtlMoveMemory(buf(0), ptr, length)
        ColumnBlob = buf
    End Function
#End Region

#Region "Statement bindings"
    Public Function BindData(ByVal OneBasedParamIndex As Integer, ByVal Value As Object, ByVal ValueType As String, ByVal stmtHandle As Integer) As Integer
        Select Case ValueType
            Case "String"
                Return BindText(OneBasedParamIndex, Value.ToString, stmtHandle)
            Case "Int32"
                Return BindInt32(OneBasedParamIndex, CInt(Value), stmtHandle)
            Case "Double"
                Return BindDouble(OneBasedParamIndex, CDbl(Value), stmtHandle)
            Case "Date"
                Return BindDate(OneBasedParamIndex, CDate(Value), stmtHandle)
            Case "NULL"
                Return BindNull(OneBasedParamIndex, stmtHandle)
            Case Else
                Return BindBlob(OneBasedParamIndex, Value, stmtHandle)
        End Select
    End Function
    Private Function BindText(ByVal OneBasedParamIndex As Integer, ByVal Value As String, ByVal stmtHandle As Integer) As Integer
        BindText = sqlite3_bind_text16(stmtHandle, OneBasedParamIndex, StrToByte16(Value), -1, SQLITE_TRANSIENT)
    End Function
    Private Function BindDouble(ByVal OneBasedParamIndex As Integer, ByVal Value As Double, ByVal stmtHandle As Integer) As Integer
        BindDouble = sqlite3_bind_double(stmtHandle, OneBasedParamIndex, Value)
    End Function
    Private Function BindInt32(ByVal OneBasedParamIndex As Integer, ByVal Value As Integer, ByVal stmtHandle As Integer) As Integer
        BindInt32 = sqlite3_bind_int(stmtHandle, OneBasedParamIndex, Value)
    End Function
    Private Function BindDate(ByVal OneBasedParamIndex As Integer, ByVal Value As Date, ByVal stmtHandle As Integer) As Integer
        BindDate = sqlite3_bind_double(stmtHandle, OneBasedParamIndex, ToJulianDay(Value))
    End Function
    Private Function BindBlob(ByVal OneBasedParamIndex As Integer, ByRef Value() As Byte, ByVal stmtHandle As Integer) As Integer
        Dim length As Integer
        length = UBound(Value) - LBound(Value) + 1
        BindBlob = sqlite3_bind_blob(stmtHandle, OneBasedParamIndex, Value(0), length, SQLITE_TRANSIENT)
    End Function
    Private Function BindNull(ByVal OneBasedParamIndex As Integer, ByVal stmtHandle As Integer) As Integer
        BindNull = sqlite3_bind_null(stmtHandle, OneBasedParamIndex)
    End Function

    Public Function BindParameterCount(ByVal stmtHandle As Integer) As Integer
        BindParameterCount = sqlite3_bind_parameter_count(stmtHandle)
    End Function
    Public Function BindParameterName(ByVal OneBasedParamIndex As Integer, ByVal stmtHandle As Integer) As String
        BindParameterName = Utf8PtrToString(sqlite3_bind_parameter_name(stmtHandle, OneBasedParamIndex))
    End Function
    Public Function BindParameterIndex(ByVal paramName As String, ByVal stmtHandle As Integer) As Integer
        Dim buf() As Byte
        buf = StringToUtf8Bytes(paramName)
        BindParameterIndex = sqlite3_bind_parameter_index(stmtHandle, buf(0))
    End Function
    Public Function ClearBindings(ByVal stmtHandle As Integer) As Integer
        ClearBindings = sqlite3_clear_bindings(stmtHandle)
    End Function
#End Region

#Region "Backup"
    Public Function SQLite3Sleep(ByVal timeToSleepInMs As Integer) As Integer
        SQLite3Sleep = sqlite3_sleep(timeToSleepInMs)
    End Function

#If Win64 Then
Public Function SQLite3BackupInit(ByVal dbHandleDestination As IntegerPtr, ByVal destinationName As String, ByVal dbHandleSource As IntegerPtr, ByVal sourceName As String) As IntegerPtr
#Else
    Public Function SQLite3BackupInit(ByVal dbHandleDestination As Integer, ByVal destinationName As String, ByVal dbHandleSource As Integer, ByVal sourceName As String) As Integer
#End If
        Dim bufDestinationName() As Byte
        Dim bufSourceName() As Byte
        bufDestinationName = StringToUtf8Bytes(destinationName)
        bufSourceName = StringToUtf8Bytes(sourceName)
        SQLite3BackupInit = sqlite3_backup_init(dbHandleDestination, bufDestinationName(0), dbHandleSource, bufSourceName(0))
    End Function

#If Win64 Then
Public Function SQLite3BackupFinish(ByVal backupHandle As IntegerPtr) As Integer
#Else
    Public Function SQLite3BackupFinish(ByVal backupHandle As Integer) As Integer
#End If
        SQLite3BackupFinish = sqlite3_backup_finish(backupHandle)
    End Function

#If Win64 Then
Public Function SQLite3BackupStep(ByVal backupHandle As IntegerPtr, ByVal numberOfPages) As Integer
#Else
    Public Function SQLite3BackupStep(ByVal backupHandle As Integer, ByVal numberOfPages As Integer) As Integer
#End If
        SQLite3BackupStep = sqlite3_backup_step(backupHandle, numberOfPages)
    End Function

#If Win64 Then
Public Function SQLite3BackupPageCount(ByVal backupHandle As IntegerPtr) As Integer
#Else
    Public Function SQLite3BackupPageCount(ByVal backupHandle As Integer) As Integer
#End If
        SQLite3BackupPageCount = sqlite3_backup_pagecount(backupHandle)
    End Function

#If Win64 Then
Public Function SQLite3BackupRemaining(ByVal backupHandle As IntegerPtr) As Integer
#Else
    Public Function SQLite3BackupRemaining(ByVal backupHandle As Integer) As Integer
#End If
        SQLite3BackupRemaining = sqlite3_backup_remaining(backupHandle)
    End Function
#End Region

#Region "String Helpers"
#If Win64 Then
Function Utf8PtrToString(ByVal pUtf8String As IntegerPtr) As String
#Else
    Function Utf8PtrToString(ByVal pUtf8String As Integer) As String
#End If
        Dim cSize As Integer
        Dim RetVal As Integer

        Dim b(0) As Byte
        cSize = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, b, 0) * 2
        ' cSize includes the terminating null character
        If cSize <= 1 Then
            Return ""
        End If

        'Utf8PtrToString = New String("*", cSize - 1) ' and a termintating null char.
        ReDim b(cSize - 1)

        RetVal = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, b, cSize)
        If RetVal = 0 Then
            Return ("Utf8PtrToString Error:" & Err.LastDllError)
        End If

        Return ByteToStr16(b)
    End Function

    Function StringToUtf8Bytes(ByVal str As String) As Byte()
        Dim bSize As Integer
        Dim RetVal As Integer
        Dim buf() As Byte

        bSize = WideCharToMultiByte(CP_UTF8, 0, str, -1, 0, 0, 0, 0)
        If bSize = 0 Then
            Return Nothing
        End If

        ReDim buf(bSize)
        RetVal = WideCharToMultiByte(CP_UTF8, 0, str, -1, buf(0), bSize, 0, 0)
        If RetVal = 0 Then
            Debug.Print("StringToUtf8Bytes Error:", Err.LastDllError)
            Return Nothing
        End If
        Return buf
    End Function

#If Win64 Then
Function Utf16PtrToString(ByVal pUtf16String As IntegerPtr) As String
#Else
    Function Utf16PtrToString(ByVal pUtf16String As Integer) As String
#End If
        Dim StrLen As Integer

        StrLen = lstrlenW(pUtf16String)
        Utf16PtrToString = New String("*", StrLen)
        lstrcpynW(Utf16PtrToString, pUtf16String, StrLen)
    End Function

    Private Function StrToByte16(ByVal str As String) As Byte()
        Return System.Text.Encoding.Unicode.GetBytes(str)
    End Function
    Private Function ByteToStr16(ByRef b() As Byte) As String
        Return System.Text.Encoding.Unicode.GetString(b)
    End Function

#End Region

#Region "Date Helpers"
    Public Function ToJulianDay(ByVal oleDate As Date) As Double
        ToJulianDay = oleDate.ToOADate() + JULIANDAY_OFFSET
    End Function

    Public Function FromJulianDay(ByVal julianDay As Double) As Date
        FromJulianDay = Date.FromOADate(julianDay - JULIANDAY_OFFSET)
    End Function

#End Region

#Region "SQLite3 Helper Functions"
    ''' <summary>
    ''' 开启外键约束
    ''' </summary>
    ''' <param name="bOn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ForeignKeys(Optional ByVal bOn As Boolean = True) As Integer
        Dim stmtHandle As Integer
        If bOn Then
            stateValue = Prepare("pragma foreign_keys = ON", stmtHandle)
        Else
            stateValue = Prepare("pragma foreign_keys = OFF", stmtHandle)
        End If
        If stateValue Then Return stateValue

        stateValue = GetStep(stmtHandle)
        If stateValue <> SQLITE_DONE Then
            stmtFinalize(stmtHandle)
            Return stateValue
        End If

        Return stmtFinalize(stmtHandle)
    End Function

    Public Function ExecuteNonQuery(ByVal SqlCommand As String) As Integer
        Dim stmtHandle As Integer
        stateValue = Prepare(SqlCommand, stmtHandle)
        If stateValue Then Return stateValue

        stateValue = GetStep(stmtHandle)
        If stateValue <> SQLITE_DONE Then
            stmtFinalize(stmtHandle)
            Return stateValue
        End If

        Return stmtFinalize(stmtHandle)
    End Function

    ''' <summary>
    ''' 获取sql返回的第0行，0列数据
    ''' </summary>
    ''' <param name="sql"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function GetColZeroValue(ByVal sql As String) As Integer
        Dim stmtHandle As Integer
        stateValue = Prepare(sql, stmtHandle)
        If stateValue Then Return -1
        stateValue = GetStep(stmtHandle)
        If stateValue <> SQLITE_ROW Then
            stateValue = stmtFinalize(stmtHandle)
            Return -1
        End If

        Dim tmp As Object = ColumnValue(0, ColumnType(0, stmtHandle), stmtHandle)
        If tmp Is DBNull.Value Then tmp = 0

        ' Finalize (delete) the statement
        stateValue = stmtFinalize(stmtHandle)
        If stateValue Then Return -1

        Return CInt(tmp)
    End Function

    ''' <summary>
    ''' 读取数据，返回DataTable
    ''' </summary>
    ''' <param name="sqlQuery"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteQuery(ByVal sqlQuery As String) As DataTable
        Dim stmtHandle As Integer
        stateValue = Prepare(sqlQuery, stmtHandle)
        If stateValue Then Return Nothing
        stateValue = GetStep(stmtHandle)
        If stateValue <> SQLITE_ROW Then
            stateValue = stmtFinalize(stmtHandle)
            Return Nothing
        End If

        Dim colCount As Integer = ColumnCount(stmtHandle)
        '初始化DataTable的列
        Dim dt As DataTable = New DataTable
        For i As Integer = 0 To colCount - 1
            Dim c As DataColumn = dt.Columns.Add()
            c.ColumnName = ColumnName(i, stmtHandle)

            Select Case ColumnType(i, stmtHandle)
                Case SQLITE_INTEGER
                    c.DataType = Type.GetType("System.Int32")
                Case SQLITE_FLOAT
                    c.DataType = Type.GetType("System.Double")
                Case SQLITE_TEXT
                    c.DataType = Type.GetType("System.String")
                Case SQLITE_BLOB
                    c.DataType = Type.GetType("System.Object")
                Case SQLITE_NULL
                    c.DataType = Type.GetType("System.Object")
                Case Else
                    c.DataType = Type.GetType("System.Object")
            End Select
        Next

        Do While stateValue = SQLITE_ROW
            Dim r As DataRow = dt.NewRow
            For i As Integer = 0 To colCount - 1
                r.Item(i) = ColumnValue(i, ColumnType(i, stmtHandle), stmtHandle)
            Next
            dt.Rows.Add(r)
            stateValue = GetStep(stmtHandle)
        Loop

        ' Finalize (delete) the statement
        stateValue = stmtFinalize(stmtHandle)
        If stateValue Then Return Nothing

        Return dt
    End Function

    '出错 Error information
    Private Function ErrMsg() As String
        Return Utf8PtrToString(sqlite3_errmsg(DbHandle))
    End Function
    ' Change Counts
    Public Function Changes() As Integer
        Return sqlite3_changes(DbHandle)
    End Function
    Public Function SQLite3TotalChanges() As Integer
        Return sqlite3_total_changes(DbHandle)
    End Function

    ' Statements
    Public Function Prepare(ByVal SQL As String, ByRef stmtHandle As Integer) As Integer
        ' Only the first statement (up to ';') is prepared. Currently we don't retrieve the 'tail' pointer.
        Return sqlite3_prepare16_v2(DbHandle, StrToByte16(SQL), Len(SQL) * 2, stmtHandle, 0)
    End Function
    ' Start running the statement
    Public Function GetStep(ByVal stmtHandle As Integer) As Integer
        Return sqlite3_step(stmtHandle)
    End Function

    Public Function Reset(ByVal stmtHandle As Integer) As Integer
        Return sqlite3_reset(stmtHandle)
    End Function
    ' Finalize (delete) the statement
    Public Function stmtFinalize(ByVal stmtHandle As Integer) As Integer
        Return sqlite3_finalize(stmtHandle)
    End Function

    Private Function TypeName(ByVal SQLiteType As Integer) As String
        Select Case SQLiteType
            Case SQLITE_INTEGER
                Return "INTEGER"
            Case SQLITE_FLOAT
                Return "FLOAT"
            Case SQLITE_TEXT
                Return "TEXT"
            Case SQLITE_BLOB
                Return "BLOB"
            Case SQLITE_NULL
                Return "NULL"
            Case Else
                Return ""
        End Select
    End Function

    Function ColumnValue(ByVal ZeroBasedColIndex As Integer, ByVal SQLiteType As Integer, ByVal stmtHandle As Integer) As Object
        Select Case SQLiteType
            Case SQLITE_INTEGER
                Return ColumnInt32(ZeroBasedColIndex, stmtHandle)
            Case SQLITE_FLOAT
                Return ColumnDouble(ZeroBasedColIndex, stmtHandle)
            Case SQLITE_TEXT
                Return ColumnText(ZeroBasedColIndex, stmtHandle)
            Case SQLITE_BLOB
                Return ColumnText(ZeroBasedColIndex, stmtHandle)
            Case SQLITE_NULL
                Return DBNull.Value
        End Select

        Return Nothing
    End Function
#End Region

#Region "获取数据库表格、字段等信息"
    Function GetTableInfo() As TableInfo()
        Dim strSql As String = "select name from sqlite_master where type='table' and name<>'sqlite_sequence'"
        Dim dt As DataTable = ExecuteQuery(strSql)
        If dt Is Nothing Then Return Nothing

        Dim ret(dt.Rows.Count - 1) As TableInfo
        For i As Integer = 0 To ret.Length - 1
            ret(i).Name = dt.Rows(i).Item("name")
            ret(i).Field = GetFieldInfo(ret(i).Name)
            If ret(i).Field.Name(0) = "ID" Then ret(i).bHasID = True
        Next

        Return ret
    End Function
    ''' <summary>
    ''' 根据tableName，读取字段的信息
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function GetFieldInfo(ByVal tableName As String) As FieldInfo
        Dim dt As DataTable = ExecuteQuery(Strings.Format(tableName, "PRAGMA table_info({0})"))
        Dim fieldCount As Integer = dt.Rows.Count
        Dim pkCount As Integer = 0 '主键数量

        If fieldCount Then
            Dim fi As FieldInfo = Nothing
            ReDim fi.Name(fieldCount - 1)
            ReDim fi.Type(fieldCount - 1)
            ReDim fi.Pointer(fieldCount - 1)

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

End Class
Imports System.Xml.Serialization
Imports System.IO

Module MStruct
    Public DB_Info As DBInfo
    Public DicSet As Dictionary(Of String, String)
    Public FileSet As String = Application.StartupPath & "\Set.set"
    Public SQLiteDllPath As String = "E:\00-学习资料\00-Excel学习资料\SQLite"
    Public Const PAGE_NUM As Integer = 20
    Public cdb As CSQLite

    Enum ExtendFieldEnum
        Tables
        FieldName
        ExtendFieldName
        ExtendFieldType
        Count
    End Enum

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
    ''' 表信息
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure TableInfo
        Dim Name As String
        Dim Field As FieldInfo

        Dim bHasID As Boolean '是否有ID这个字段
        Dim bUseMyIdTables As Boolean '是否有表格引用了本表的ID
        Dim UseMyIdTables() As Integer '有哪些表格引用了本表的ID
        Dim ExtendField As ExtendFieldInfo '把xxID的字段扩展出来
        Dim SqlExtend As String   '扩展的sql

        Dim dt As DataTable '记录表的数据，使用的时候判断这个是否是nothing，不是就可以直接使用，不需要重复查找
    End Structure

    ''' <summary>
    ''' 数据库信息
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure DBInfo
        Dim Path As String
        Dim ActivateTable As String '当前在操作的表
        Dim ActivateTableIndex As Integer
        Dim Tables() As CSQLite.TableInfo   'TableInfo
        Dim DicTableIndex As Dictionary(Of String, Integer)
    End Structure

End Module

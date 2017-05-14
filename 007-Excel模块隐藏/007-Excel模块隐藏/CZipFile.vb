Imports System.IO
Imports System.IO.Compression
Imports System.Runtime.InteropServices

'-----------------------------------------ZIP文件结构-------------------------------------------
'ZIP文件结构的说明，下面这个帖子介绍的挺详细
'http://club.excelhome.net/thread-1251530-1-1.html
'把a1.txt、a2.txt、a3.txt压缩到a.zip中以后
'其内容在磁盘上的摆放顺序为h1、a1[]、h2、a2[]、h3、a3[]、c1、c2、c3、EOCD
'其中a1[]、a2[]、a3[]是三个文本文件压缩后的数据块
'h1、h2、h3和c1、c2、c3分别是三个数据块对应的Local File Header和Central Directory FileHeader结构
'EOCD是文件中唯一的EndOfCentralDirectory结构。
'-----------------------------------------ZIP文件结构-------------------------------------------

Public Class CZipFile
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal Destination As IntPtr, ByVal Source As IntPtr, ByVal Length As Integer)

#Region "FileStructure"

    'local file header+file data+data descriptor这是一段ZIP压缩数据
    'Local file header
    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Private Structure LocalFileHeader
        Dim Signature As Integer           '文件头标识 0x04034b50
        Dim VersionExtract As Short   '解压文件所需最低版本
        Dim GeneralBit As Short       '通用位标记
        Dim CompressionMethod As Short '压缩方法
        Dim FileModiTime As Short     '文件最后修改时间
        Dim FileModiDate As Short     '文件最后修改日期
        Dim CRC_32 As Integer             '说明采用的算法
        Dim CompressedSize As Integer      '压缩后的大小
        Dim UncompressedSize As Integer    '压缩前的大小
        Dim FileNameLength As Short      '文件名长度 (n)
        Dim ExtraFieldLength As Short '附加信息长度 (m)

        '    FileName() As Byte          '文件名
        '    ExtraField() As Byte        '扩展区
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Private Structure CentralDirectoryHeader
        Dim Signature As Integer               'HEX 50 4B 01 02
        Dim VersionMadeBy As Short
        Dim VersionNeeded As Short
        Dim GeneralBitFlag As Short
        Dim CompressionMethod As Short
        Dim LastModifyTime As Short
        Dim LastModifyDate As Short
        Dim CRC32 As Integer
        Dim CompressedSize As Integer
        Dim UncompressedSize As Integer
        Dim FileNameLength As Short       '文件名长度(n)
        Dim ExtraFieldLength As Short     '附加信息长度 (m)
        Dim FileCommentLength As Short    '文件附注长度 (k)
        Dim StartDiskNumber As Short      '文件起始位置的磁盘编号【3】
        Dim InteralFileAttrib As Short    '内部文件属性
        Dim ExternalFileAttrib As Integer      '外部文件属性
        Dim LocalFileHeaderOffset As Integer   '对应的Local File  Header在文件中的起始位置。
        '                                   46  n 文件名
        '                                   46+n    m   附加信息
        '                                   46+n+m  k   文件附注
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Private Structure EndOfCentralDirectory
        Dim Signature As Integer                     '核心目录结束标记 0x06054b50
        Dim NumberOfThisDisk As Short              '当前磁盘编号
        Dim DiskDirectoryStarts As Short           '第一条Central  Directory起始位置所在的磁盘编号
        Dim NumberOfCDRecordsOnThisDisk As Short  '当前磁盘上的Central  Directory数量
        Dim TotalNumberOfCDRecords As Short       'Zip文件中全部Central  Directory的总数量
        Dim SizeOfCD As Integer                        '全部Central  Directory的合计字节长度
        Dim OffsetOfCD As Integer                      '第一条Central  directory的起始位置在zip文件中的位置
        Dim CommentLength As Short                '注释长度
        '    Comment() as Byte                       '注释内容
    End Structure


#End Region

    Function GetStructure(ByRef file_byte() As Byte, ByRef FileAddress As IntPtr) As Boolean
        Dim LFH As LocalFileHeader, LFH_offset As Integer
        Dim ECD As EndOfCentralDirectory, ECD_offset As Integer
        Dim CDH As CentralDirectoryHeader, CDH_offset As Integer
        Dim file_name As String
        Dim vba_byte() As Byte

        For i As Integer = file_byte.Length - 1 - 4 To 0 Step -1
            '查找EndOfCentralDirectory的Signature标识
            If file_byte(i) = &H50 AndAlso file_byte(i + 1) = &H4B AndAlso file_byte(i + 2) = &H5 AndAlso file_byte(i + 3) = &H6 Then
                ECD_offset = i
                Exit For
            End If
        Next

        ECD = Marshal.PtrToStructure(FileAddress + ECD_offset, ECD.GetType)
        'CopyMemory(GCHandle.Alloc(ECD.Signature, GCHandleType.Pinned).AddrOfPinnedObject(), FileAddress + ECD_offset, Len(ECD))
        CDH_offset = ECD.OffsetOfCD
        For i As Integer = 0 To ECD.TotalNumberOfCDRecords - 1
            CDH = Marshal.PtrToStructure(FileAddress + CDH_offset, CDH.GetType)

            file_name = Marshal.PtrToStringAnsi(FileAddress + CDH_offset + Len(CDH) + CDH.ExtraFieldLength + CDH.FileCommentLength, CDH.FileNameLength)
            'xl/vbaProject.bin
            Debug.Print(file_name)
            If "xl/vbaProject.bin" = file_name Then
                LFH_offset = CDH.LocalFileHeaderOffset
                LFH = Marshal.PtrToStructure(FileAddress + LFH_offset, LFH.GetType)

                ReDim vba_byte(LFH.CompressedSize - 1)
                Dim p1 As IntPtr = GCHandle.Alloc(vba_byte, GCHandleType.Pinned).AddrOfPinnedObject()
                CopyMemory(p1, FileAddress + LFH_offset + LFH.FileNameLength + LFH.ExtraFieldLength + Len(LFH), LFH.CompressedSize)

                'file_byte = Decompress("Deflate", vba_byte)
                file_byte = UnCompression(vba_byte, LFH.UncompressedSize)
                FileAddress = GCHandle.Alloc(file_byte, GCHandleType.Pinned).AddrOfPinnedObject()

                Return True
            End If

            CDH_offset = CDH_offset + Len(CDH) + CDH.ExtraFieldLength + CDH.FileCommentLength + CDH.FileNameLength
        Next

        Return False
    End Function


    Function UnCompression(buffer() As Byte, i_UnCompressedSize As Integer) As Byte()
        'Dim infile As System.IO.Stream = New System.IO.FileStream("C:\Documents and Settings\xyj\桌面\txt", FileMode.Open, FileAccess.Read, FileShare.Read)
        'Dim buffer(infile.Length - 1) As Byte
        'Dim count As Integer = infile.Read(buffer, 0, buffer.Length)

        'infile.Position = 0
        Dim stream_src As Stream = New MemoryStream(buffer)
        'stream_src.Position = 0
        Dim zipStream As Stream = New DeflateStream(stream_src, CompressionMode.Decompress, True)

        Dim decompressedBuffer(i_UnCompressedSize - 1) As Byte
        zipStream.Read(decompressedBuffer, 0, decompressedBuffer.Length)

        zipStream.Close()
        stream_src.Close()

        'Dim fw As System.IO.FileStream = New System.IO.FileStream("C:\Documents and Settings\xyj\桌面\模块1.txt", FileMode.Create)
        'fw.Write(decompressedBuffer, 0, decompressedBuffer.Length)
        'fw.Close()

        Return decompressedBuffer
    End Function

    Function Decompress(ByVal algo As String, ByVal data() As Byte) As Byte()
        Try
            Dim sw As New Stopwatch
            '---复制数据(压缩的)到ms---
            Dim ms As New MemoryStream(data)
            Dim zipStream As Stream = Nothing

            '---开始秒表---
            sw.Start()
            '---使用存储在ms中的数据解压---

            If algo = "Gzip" Then
                zipStream = New GZipStream(ms, CompressionMode.Decompress)
            ElseIf algo = "Deflate" Then
                zipStream = New DeflateStream(ms, CompressionMode.Decompress, True)
            End If

            '---用来存储解压的数据---
            Dim dc_data() As Byte

            '---解压的数据存储于zipStream中; 
            '把它们提取到一个字节数组中---
            dc_data = RetrieveBytesFromStream(zipStream, data.Length)

            '---停止秒表---
            sw.Stop()
            'lblMessage.Text = "Decompression completed. Time spent: " & sw.ElapsedMilliseconds & "ms" &       "， Original size: " & dc_data.Length
            Return dc_data
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return Nothing
        End Try

    End Function

    Function RetrieveBytesFromStream(ByVal stream As Stream, ByVal bytesblock As Integer) As Byte()

        '---从一个流对象中检索字节---
        Dim data() As Byte
        Dim totalCount As Integer = 0
        Try
            While True

                '---逐渐地增加数据字节数组-的大小--
                ReDim Preserve data(totalCount + bytesblock)
                Dim bytesRead As Integer = stream.Read(data, totalCount, bytesblock)
                If bytesRead = 0 Then
                    Exit While
                End If
                totalCount += bytesRead
            End While
            '---确保字节数组正确包含提取的字节数---
            ReDim Preserve data(totalCount - 1)
            Return data
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return Nothing
        End Try

    End Function


End Class

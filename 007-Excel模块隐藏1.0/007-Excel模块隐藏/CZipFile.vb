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
    Inherits CCompdocFile

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

    Structure ZipStructure
        Dim Offset As Integer   '偏移地址
        Dim CompressionSize As Integer     '压缩后大小，如果是H1或C1就直接记录结构的大小+FileName+ExtraField之类的
        Dim UnCompressionSize As Integer     '压缩前大小
        Dim LocalFileHeaderOffset As Integer    '记录CentralDirectoryHeader的LocalFileHeaderOffset
    End Structure

    Private zipFileByte() As Byte
    Private zipFileAddress As IntPtr
    Private zipArrData() As ZipStructure  '记录每1个数据块，重写的时候方便
    Private p_vbaBin As Integer
    Private i_files As Integer    '有几个数据块
    Private firstCDHoffset As Integer '第1个CDH

    Sub New(file_name As String)
        MyBase.New(file_name)
    End Sub

    Overrides Sub ReWriteFile(ByRef arr_byte_to_write() As Byte, arr_address(,) As Integer, step_address As Integer, n_size As Integer, arr_dir_address_j As Integer, ByRef stream_size As Integer)

        Dim p1 As IntPtr = GCHandle.Alloc(arr_byte_to_write, GCHandleType.Pinned).AddrOfPinnedObject()
        For i_address As Integer = 0 To n_size
            CopyMemory(FileAddress + arr_address(i_address, 1), p1 + i_address * step_address, step_address)
        Next

        '重设PROJECT目录的长度
        stream_size = arr_byte_to_write.Length '
        Dim tmp_i As Integer = arr_byte_to_write.Length
        p1 = GCHandle.Alloc(tmp_i, GCHandleType.Pinned).AddrOfPinnedObject()
        CopyMemory(arr_dir_address_j + DIR_SIZE - 4 * 2 + FileAddress, p1, 4）

        '将修改的后的vbaBin数据进行压缩
        Dim b() As Byte = Compression(file_byte)
        Dim tmp_compression_size As Integer = b.Length
        Dim tmp_un_compression_size As Integer = file_byte.Length

        '修改LocalFileHeader  CompressedSize
        CopyMemory(zipFileAddress + zipArrData(p_vbaBin - 1).Offset + 18, GCHandle.Alloc(tmp_compression_size, GCHandleType.Pinned).AddrOfPinnedObject(), 4)
        'zipArrData(files * 2 + (p_vbaBin - 1) / 2).Offset + 24     'UnCompressedSize
        CopyMemory(zipFileAddress + zipArrData(p_vbaBin - 1).Offset + 22, GCHandle.Alloc(tmp_un_compression_size, GCHandleType.Pinned).AddrOfPinnedObject(), 4)


        'h1、a1[]、h2、a2[]、h3、a3[]、c1、c2、c3、EOCD
        '修改CentralDirectoryHeader     'CompressedSize
        CopyMemory(zipFileAddress + zipArrData(i_files * 2 + (p_vbaBin - 1) / 2).Offset + 20, GCHandle.Alloc(tmp_compression_size, GCHandleType.Pinned).AddrOfPinnedObject(), 4)
        'UnCompressedSize
        CopyMemory(zipFileAddress + zipArrData(i_files * 2 + (p_vbaBin - 1) / 2).Offset + 24, GCHandle.Alloc(tmp_un_compression_size, GCHandleType.Pinned).AddrOfPinnedObject(), 4)

        '修改EndOfCentralDirectory    OffsetOfCD     16
        'vbaBin修改后的压缩大小 与 之前的压缩大小对比
        Dim offset_eod As Integer = tmp_compression_size - zipArrData(p_vbaBin).CompressionSize
        offset_eod = firstCDHoffset + offset_eod
        CopyMemory(zipFileAddress + zipArrData(i_files * 3).Offset + 16, GCHandle.Alloc(offset_eod, GCHandleType.Pinned).AddrOfPinnedObject(), 4)

        'p_vbaBin之后的data位置有变，对应的CentralDirectoryHeader.LocalFileHeaderOffset变化 len - 4(42)
        Dim tmp_cdh As CentralDirectoryHeader
        offset_eod = tmp_compression_size - zipArrData(p_vbaBin).CompressionSize
        For i As Integer = i_files * 2 + (p_vbaBin - 1) / 2 + 1 To i_files * 3 - 1
            Dim tmp As Integer = zipArrData(i）.LocalFileHeaderOffset + offset_eod
            CopyMemory(zipFileAddress + zipArrData(i).Offset + Len(tmp_cdh) - 4, GCHandle.Alloc(tmp, GCHandleType.Pinned).AddrOfPinnedObject(), 4)
        Next

        '从p_vbaBin开始写入文件
        Dim fw As FileStream = New FileStream(Me.path & "new.xlsm", FileMode.Create)
        'Dim fw As FileStream = New FileStream(Me.path, FileMode.Open)

        Dim i_seek As Integer = 0

        For i As Integer = 0 To p_vbaBin - 1
            fw.Seek(i_seek, origin:=0)
            fw.Write(zipFileByte, zipArrData(i).Offset, zipArrData(i).CompressionSize)
            i_seek += zipArrData(i).CompressionSize
        Next
        'i_seek = zipArrData（p_vbaBin - 1).Offset
        'fw.Seek(i_seek, origin:=0)
        'fw.Write(zipFileByte, zipArrData(p_vbaBin - 1).Offset, zipArrData(p_vbaBin - 1).CompressionSize)

        i_seek = zipArrData（p_vbaBin).Offset
        fw.Seek(i_seek, origin:=0)
        fw.Write(b, 0, b.Length)

        i_seek += b.Length
        For i = p_vbaBin + 1 To i_files * 3
            fw.Seek(i_seek, origin:=0)
            fw.Write(zipFileByte, zipArrData(i).Offset, zipArrData(i).CompressionSize)
            i_seek += zipArrData(i).CompressionSize
        Next

        fw.Close()
    End Sub

    Overrides Function GetFileByte() As Integer
        MFunc.read_file_to_byte(Me.path, zipFileByte)
        FileAddress = GCHandle.Alloc(zipFileByte, GCHandleType.Pinned).AddrOfPinnedObject()
        zipFileAddress = FileAddress

        GetStructure()

        Return 1
    End Function

    Function GetStructure() As Boolean
        Dim LFH As LocalFileHeader, LFH_offset As Integer
        Dim ECD As EndOfCentralDirectory, ECD_offset As Integer
        Dim CDH As CentralDirectoryHeader, CDH_offset As Integer
        Dim file_name As String
        Dim vba_byte() As Byte
        Dim tmpFileAddress As IntPtr = FileAddress

        For i As Integer = zipFileByte.Length - 1 - 4 To 0 Step -1
            '查找EndOfCentralDirectory的Signature标识
            If zipFileByte(i) = &H50 AndAlso zipFileByte(i + 1) = &H4B AndAlso zipFileByte(i + 2) = &H5 AndAlso zipFileByte(i + 3) = &H6 Then
                ECD_offset = i
                Exit For
            End If
        Next

        ECD = Marshal.PtrToStructure(tmpFileAddress + ECD_offset, ECD.GetType)
        'CopyMemory(GCHandle.Alloc(ECD.Signature, GCHandleType.Pinned).AddrOfPinnedObject(), FileAddress + ECD_offset, Len(ECD))
        CDH_offset = ECD.OffsetOfCD
        'h1、a1[]、h2、a2[]、h3、a3[]、c1、c2、c3、EOCD
        i_files = ECD.TotalNumberOfCDRecords
        firstCDHoffset = CDH_offset

        ReDim zipArrData(i_files * 3) '1个数据对应3个，外加1个EOCD
        zipArrData(i_files * 3).Offset = ECD_offset
        zipArrData(i_files * 3).CompressionSize = Len(ECD) + ECD.CommentLength

        For i As Integer = 0 To ECD.TotalNumberOfCDRecords - 1
            CDH = Marshal.PtrToStructure(tmpFileAddress + CDH_offset, CDH.GetType)
            'c1、c2、c3
            zipArrData(i_files * 2 + i).Offset = CDH_offset
            zipArrData(i_files * 2 + i).CompressionSize = Len(CDH) + CDH.ExtraFieldLength + CDH.FileCommentLength + CDH.FileNameLength
            zipArrData(i_files * 2 + i).LocalFileHeaderOffset = CDH.LocalFileHeaderOffset

            file_name = Marshal.PtrToStringAnsi(tmpFileAddress + CDH_offset + Len(CDH) + CDH.ExtraFieldLength + CDH.FileCommentLength, CDH.FileNameLength)

            LFH_offset = CDH.LocalFileHeaderOffset
            LFH = Marshal.PtrToStructure(tmpFileAddress + LFH_offset, LFH.GetType)
            'h1、a1[]、h2、a2[]、h3、a3[]
            zipArrData(i * 2).Offset = LFH_offset
            zipArrData(i * 2).CompressionSize = Len(LFH) + LFH.ExtraFieldLength + LFH.FileNameLength
            '数据区域
            zipArrData(i * 2 + 1).Offset = LFH_offset + zipArrData(i * 2).CompressionSize
            zipArrData(i * 2 + 1).CompressionSize = LFH.CompressedSize
            zipArrData(i * 2 + 1).UnCompressionSize = LFH.UncompressedSize

            'xl/vbaProject.bin
            'Debug.Print(file_name)
            If "xl/vbaProject.bin" = file_name Then
                ReDim vba_byte(LFH.CompressedSize - 1)
                Dim p1 As IntPtr = GCHandle.Alloc(vba_byte, GCHandleType.Pinned).AddrOfPinnedObject()
                CopyMemory(p1, FileAddress + LFH_offset + LFH.FileNameLength + LFH.ExtraFieldLength + Len(LFH), LFH.CompressedSize)
                'file_byte = Decompress("Deflate", vba_byte)
                file_byte = UnCompression(vba_byte, LFH.UncompressedSize)
                FileAddress = GCHandle.Alloc(file_byte, GCHandleType.Pinned).AddrOfPinnedObject()

                p_vbaBin = i * 2 + 1
            End If

            CDH_offset = CDH_offset + Len(CDH) + CDH.ExtraFieldLength + CDH.FileCommentLength + CDH.FileNameLength
        Next

        Return True
    End Function

    Function Compression(buffer() As Byte) As Byte()
        'Dim stream_src As Stream = New MemoryStream(buffer)
        Dim ms As New MemoryStream

        Dim zipStream As Stream = New DeflateStream(ms, CompressionMode.Compress, True)
        zipStream.Write(buffer, 0, buffer.Length)
        zipStream.Close()
        zipStream.Dispose()

        Dim compressedBuffer(ms.Length - 1) As Byte
        ms.Position = 0
        ms.Read(compressedBuffer, 0, compressedBuffer.Length)
        ms.Close()
        ms.Dispose()

        Return compressedBuffer
    End Function
    Function UnCompression(buffer() As Byte, ByVal i_UnCompressedSize As Integer) As Byte()
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

End Class

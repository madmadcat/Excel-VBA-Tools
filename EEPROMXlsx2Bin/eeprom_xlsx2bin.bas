

Sub eeprom_xlsx2bin()
    Dim sht As Worksheet
    Dim val As String
    Dim msg, Response
    
    Dim bufferlength As Integer
    bufferlength = 0
    
    Dim buffer() '255表示索引下标，这样0-255共256个数组元素
    ReDim buffer(bufferlength)
    Dim bufferByte() As Byte '二进制数组
    ReDim bufferByte(255)
    
    
    Const start = 5 '有效数据起始行数
    Const offset = 128 '有效数据长度
    Const fileType = "txt"
    
    'FSO文本流的操作类型
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    
    '遍历全部sheets
    For Each x In Sheets
        Dim k
        k = x.UsedRange.Rows.Count - start + 1 '表单的实际有效数据长度
        '表单为Lower Mem Map
        If InStr(UCase(x.Name), UCase("Lower")) Then
            '读取指定行的内容，存入数组
            '检查表单数据长度是否正确
            
            If k Mod offset <> 0 Then '清洗不正确的表单，有效数据必须满足128的整数倍
                MsgBox ("表单数据长度异常")
                Exit Sub
            End If
            
            bufferlength = offset - 1
            '有效数据行数多于128时，可以选择读取前128或者全部
            If k / offset > 1 Then
                msg = x.Name & vbCrLf & "查找到" & Str(k) & " Bytes 数据"
                msg = msg & vbCrLf & "是:  写入全部" & Str(k) & " Bytes 数据" & vbCrLf & "否:  只写入前" & Str(offset) & " bytes 数据"
                Response = MsgBox(msg, vbYesNo)
                If Response = vbYes Then
                    bufferlength = k - 1
                End If
            End If
            
            '读取数据，存入数组
            Dim curBuffer
            curBuffer = UBound(buffer)
            If curBuffer > 0 Then
                MsgBox "表单排序异常，请保证 Lower Mem Map 排在 Upper Mem Map 之前"
                Exit Sub
            End If
            ReDim Preserve buffer(curBuffer + bufferlength)
            
            For i = 0 To bufferlength
                '数据清洗，处理可能出现的值"crc32"."checksum"
                val = Trim(x.Cells(i + 5, 4).Value)
                If UCase(val) = UCase("crc32") Or UCase(val) = UCase("checksum") Then
                    val = "00"
                End If
                buffer(i + curBuffer) = val
                'Debug.Print (buffer(i))
            Next
        End If
        
        If UBound(buffer) >= 255 Then '数组长度为256时，直接结束循环
            Exit For
        End If
        
        '表单为Upper Mem Map
        If InStr(UCase(x.Name), UCase("upper")) Then
            '读取指定行的内容，存入数组128-255
            '检查表单数据长度是否正确
            If k Mod offset <> 0 Then '清洗不正确的有效数据长度
                MsgBox ("表单数据长度异常")
                Exit Sub
            End If
            
            '有效数据行数多于128时, 且为128的整数时，当作数据异常处理
            If k / offset > 1 Then
                MsgBox x.Name & " 的数据长度不应当高于 128 Byte。"
                Exit Sub
            End If
            
            '读取数据，存入数组
            ReDim Preserve buffer(UBound(buffer) + 128)
            bufferlength = offset - 1
            For i = 0 To bufferlength
                '数据清洗，处理可能出现的值"crc32"."checksum"
                val = Trim(x.Cells(i + 5, 4).Value)
                If UCase(val) = UCase("crc32") Or UCase(val) = UCase("checksum") Then
                    val = "00"
                End If
                buffer(i + offset) = val
                'Debug.Print (buffer(i + offset))
            Next
        End If
    Next
    
    If UBound(buffer) <> 255 Then '数组长度不为256时，异常
        MsgBox "数组长度为 " & UBound(buffer) + 1 & vbCrLf & "不满足 256 Bytes 长度， 请检查表单"
        Exit Sub
    End If
    
    '数组内容写入txt文件
    '创建与xlsx文件同名的文本文件并将输入数据逐行写入
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file, parent As String
    '获取文件位置并移除扩展名
    '加入替代文件名中[PN]为上层文件夹名称的功能
    file = Replace(Application.ActiveWorkbook.FullName, ".xlsx", "")
    parent = fso.getfolder(Application.ActiveWorkbook.Path)
    
    '获取上层文件夹名称，该名称应该为料号
    Dim PN As String
    PN = Mid(parent, InStrRev(parent, "\") + 1)
    
    '拼接正确的输出文件名称
    If InStr(file, "[PN]") Then
        file = Replace(file, "[PN]", PN)
        'Debug.Print (file)
    End If
    txtfile = file & "." & fileType
    binfile = file & ".bin"
    
    'Create a TextStream.
    Set txtStream = fso.OpenTextFile(txtfile, ForWriting, True)
    
    '写入txt数据
    Dim buffer2str As String
    buffer2str = ""
    For j = 0 To UBound(buffer)
        txtStream.WriteLine buffer(j)
        bufferByte(j) = HEX_to_DEC(buffer(j)) '逐字节转存入字节数组
    Next
    '销毁FSO对象
    Set fso = Nothing
    Set txtStream = Nothing

    '数组内容写入bin文件
    Call SaveBinaryData(binfile, bufferByte)
    
    '销毁数组，释放资源
    Erase buffer
    Erase bufferByte
    
    If IsFileExists(file & ".xlsx") = False Then
        ThisWorkbook.SaveCopyAs file & ".xlsx"
    End If
    
    MsgBox file & " Binary处理完成。"
End Sub

'=============使用ADODB.Stream对象写二进制文件=====================
Function SaveBinaryData(FileName, ByteArray)
    Const adTypeBinary = 1
    Const adSaveCreateOverWrite = 2

    'Create Stream object
    Dim BinaryStream
    Set BinaryStream = CreateObject("ADODB.Stream")

    'Specify stream type - we want To save binary data.
    BinaryStream.Type = adTypeBinary

    'Open the stream And write binary data To the object
    BinaryStream.Open
    BinaryStream.Write ByteArray

    'Save binary data To disk
    BinaryStream.SaveToFile FileName, adSaveCreateOverWrite
End Function
'====================================================================

Function HEX_to_DEC(ByVal Hex As String) As Long '将十六进制转换为十进制
    Dim i As Long
    Dim B As Long
  
    Hex = UCase(Hex)
    For i = 1 To Len(Hex)
        Select Case Mid(Hex, Len(Hex) - i + 1, 1)
            Case "0": B = B + 16 ^ (i - 1) * 0
            Case "1": B = B + 16 ^ (i - 1) * 1
            Case "2": B = B + 16 ^ (i - 1) * 2
            Case "3": B = B + 16 ^ (i - 1) * 3
            Case "4": B = B + 16 ^ (i - 1) * 4
            Case "5": B = B + 16 ^ (i - 1) * 5
            Case "6": B = B + 16 ^ (i - 1) * 6
            Case "7": B = B + 16 ^ (i - 1) * 7
            Case "8": B = B + 16 ^ (i - 1) * 8
            Case "9": B = B + 16 ^ (i - 1) * 9
            Case "A": B = B + 16 ^ (i - 1) * 10
            Case "B": B = B + 16 ^ (i - 1) * 11
            Case "C": B = B + 16 ^ (i - 1) * 12
            Case "D": B = B + 16 ^ (i - 1) * 13
            Case "E": B = B + 16 ^ (i - 1) * 14
            Case "F": B = B + 16 ^ (i - 1) * 15
        End Select
    Next i
    HEX_to_DEC = B
End Function
'判断文件是否存在
Function IsFileExists(ByVal strFileName As String) As Boolean
    Dim objFileSystem As Object
 
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    If objFileSystem.fileExists(strFileName) = True Then
        IsFileExists = True
    Else
        IsFileExists = False
    End If
End Function


    ''''备忘'''' https://www.yiibai.com/vba
    'TypeName(x.Name) 获取该对象的类型
    'Application.ActiveWorkbook.Path 只返回路径
    'Application.ActiveWorkbook.FullName 返回路径及工作簿文件名
    'Application.ActiveWorkbook.Name 返回工作簿文件名
    ''''备忘''''

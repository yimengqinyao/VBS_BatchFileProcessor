' 文件分割插件 - plugins\format_tools\file_splitter.vbs
Option Explicit

Sub file_splitter_Init(pluginObj)
    pluginObj("Name") = "文件分割插件"
    pluginObj("Version") = "1.0"
    pluginObj("ProcessFile") = GetRef("ProcessSplit")
    pluginObj("SplitSize") = 0 ' KB
    pluginObj("SuccessCount") = 0
    pluginObj("Cleanup") = GetRef("file_splitter_Cleanup")
    
    ' 从命令行参数获取分割大小
    Dim i, arg
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If Left(LCase(arg), 9) = "/split:" Then
            pluginObj("SplitSize") = CLng(Mid(arg, 10))
            LogDebug "设置分割大小: " & pluginObj("SplitSize") & "KB", 1
            Exit For
        End If
    Next
End Sub

Sub ProcessSplit(fileObj, pluginObj)
    If pluginObj("SplitSize") <= 0 Then Exit Sub
    
    Dim splitSizeBytes, inputStream, outputStream, buffer, bytesRead, fileIndex, newPath
    splitSizeBytes = pluginObj("SplitSize") * 1024 ' 转换为字节
    
    ' 使用ADODB.Stream处理二进制文件
    Set inputStream = CreateObject("ADODB.Stream")
    inputStream.Type = 1 ' 二进制模式
    inputStream.Open
    inputStream.LoadFromFile fileObj.Path
    
    fileIndex = 1
    Do
        ' 读取数据块
        ReDim buffer(splitSizeBytes - 1)
        bytesRead = inputStream.Read(buffer)
        
        ' 检查是否读取到数据
        If bytesRead > 0 Then
            ' 生成新文件路径
            newPath = Left(fileObj.Path, Len(fileObj.Path) - Len(fso.GetExtensionName(fileObj.Path))) & "_" & fileIndex & "." & fso.GetExtensionName(fileObj.Path)
            
            ' 写入数据块
            Set outputStream = CreateObject("ADODB.Stream")
            outputStream.Type = 1
            outputStream.Open
            outputStream.Write buffer
            outputStream.SaveToFile newPath, 2 ' 覆盖模式
            outputStream.Close
            
            fileIndex = fileIndex + 1
        End If
    Loop While bytesRead = splitSizeBytes
    
    inputStream.Close
    
    pluginObj("SuccessCount") = pluginObj("SuccessCount") + 1
    pluginObj("Result") = "成功分割 " & pluginObj("SuccessCount") & " 个文件"
    LogDebug "分割成功: " & fileObj.Path & " → 生成 " & fileIndex - 1 & " 个文件", 3
End Sub

Sub file_splitter_Cleanup(pluginObj)
    LogDebug "文件分割插件清理完成", 3
End Sub
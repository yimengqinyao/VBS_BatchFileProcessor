' 批量重命名插件 - plugins\file_operations\file_renamer.vbs
Option Explicit

Sub file_renamer_Init(pluginObj)
    pluginObj("Name") = "批量重命名插件"
    pluginObj("Version") = "1.0"
    pluginObj("ProcessFile") = GetRef("ProcessRename")
    pluginObj("RenamePattern") = ""
    pluginObj("RenameIndex") = 1
    pluginObj("SuccessCount") = 0
    pluginObj("Cleanup") = GetRef("file_renamer_Cleanup")
    
    ' 从命令行参数获取重命名模式
    Dim i, arg
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If Left(LCase(arg), 6) = "/rename:" Then
            pluginObj("RenamePattern") = Mid(arg, 7)
            LogDebug "设置重命名模式: " & pluginObj("RenamePattern"), 1
            Exit For
        End If
    Next
End Sub

Sub ProcessRename(fileObj, pluginObj)
    If pluginObj("RenamePattern") = "" Then Exit Sub
    
    Dim newFileName, newFilePath, ext
    ext = fso.GetExtensionName(fileObj.Path)
    
    ' 替换变量
    newFileName = Replace(pluginObj("RenamePattern"), "{index}", pluginObj("RenameIndex"))
    newFileName = Replace(newFileName, "{ext}", ext)
    
    ' 处理无扩展名情况
    If ext = "" Then
        newFileName = Replace(newFileName, ".", "")
    End If
    
    newFilePath = fso.BuildPath(fso.GetParentFolderName(fileObj.Path), newFileName)
    
    ' 执行重命名
    On Error Resume Next
    fileObj.Name = newFileName
    If Err.Number = 0 Then
        pluginObj("SuccessCount") = pluginObj("SuccessCount") + 1
        pluginObj("RenameIndex") = pluginObj("RenameIndex") + 1
        pluginObj("Result") = "成功重命名 " & pluginObj("SuccessCount") & " 个文件"
        LogDebug "重命名成功: " & fileObj.Path & " → " & newFilePath, 3
    Else
        LogDebug "重命名失败: " & fileObj.Path & " 错误: " & Err.Description, 1
    End If
    On Error GoTo 0
End Sub

Sub file_renamer_Cleanup(pluginObj)
    LogDebug "批量重命名插件清理完成", 3
End Sub
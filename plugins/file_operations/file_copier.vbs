' 批量复制插件 - plugins\file_operations\file_copier.vbs
Option Explicit

Sub file_copier_Init(pluginObj)
    pluginObj("Name") = "批量复制插件"
    pluginObj("Version") = "1.0"
    pluginObj("ProcessFile") = GetRef("ProcessCopy")
    pluginObj("CopyTarget") = ""
    pluginObj("SuccessCount") = 0
    pluginObj("Cleanup") = GetRef("file_copier_Cleanup")
    
    ' 从命令行参数获取目标
    Dim i, arg
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If Left(LCase(arg), 6) = "/copyto:" Then
            pluginObj("CopyTarget") = Mid(arg, 7)
            LogDebug "设置复制目标: " & pluginObj("CopyTarget"), 1
            Exit For
        End If
    Next
End Sub

Sub ProcessCopy(fileObj, pluginObj)
    If pluginObj("CopyTarget") = "" Then Exit Sub
    
    ' 创建目标目录
    If Not fso.FolderExists(pluginObj("CopyTarget")) Then
        fso.CreateFolder pluginObj("CopyTarget")
    End If
    
    Dim targetPath
    targetPath = fso.BuildPath(pluginObj("CopyTarget"), fileObj.Name)
    
    ' 执行复制
    On Error Resume Next
    fileObj.Copy targetPath, True
    If Err.Number = 0 Then
        pluginObj("SuccessCount") = pluginObj("SuccessCount") + 1
        pluginObj("Result") = "成功复制 " & pluginObj("SuccessCount") & " 个文件"
        LogDebug "复制成功: " & fileObj.Path & " → " & targetPath, 3
    Else
        LogDebug "复制失败: " & fileObj.Path & " 错误: " & Err.Description, 1
    End If
    On Error GoTo 0
End Sub

Sub file_copier_Cleanup(pluginObj)
    LogDebug "批量复制插件清理完成", 3
End Sub
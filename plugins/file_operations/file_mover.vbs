' 批量移动插件 - plugins\file_operations\file_mover.vbs
Option Explicit

Sub file_mover_Init(pluginObj)
    pluginObj("Name") = "批量移动插件"
    pluginObj("Version") = "1.0"
    pluginObj("ProcessFile") = GetRef("ProcessMove")
    pluginObj("MoveTarget") = ""
    pluginObj("SuccessCount") = 0
    pluginObj("Cleanup") = GetRef("file_mover_Cleanup")
    
    ' 从命令行参数获取目标
    Dim i, arg
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If Left(LCase(arg), 6) = "/moveto:" Then
            pluginObj("MoveTarget") = Mid(arg, 7)
            LogDebug "设置移动目标: " & pluginObj("MoveTarget"), 1
            Exit For
        End If
    Next
End Sub

Sub ProcessMove(fileObj, pluginObj)
    If pluginObj("MoveTarget") = "" Then Exit Sub
    
    ' 创建目标目录
    If Not fso.FolderExists(pluginObj("MoveTarget")) Then
        fso.CreateFolder pluginObj("MoveTarget")
    End If
    
    Dim targetPath
    targetPath = fso.BuildPath(pluginObj("MoveTarget"), fileObj.Name)
    
    ' 执行移动
    On Error Resume Next
    fileObj.Move targetPath, True
    If Err.Number = 0 Then
        pluginObj("SuccessCount") = pluginObj("SuccessCount") + 1
        pluginObj("Result") = "成功移动 " & pluginObj("SuccessCount") & " 个文件"
        LogDebug "移动成功: " & fileObj.Path & " → " & targetPath, 3
    Else
        LogDebug "移动失败: " & fileObj.Path & " 错误: " & Err.Description, 1
    End If
    On Error GoTo 0
End Sub

Sub file_mover_Cleanup(pluginObj)
    LogDebug "批量移动插件清理完成", 3
End Sub
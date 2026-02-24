' 文件删除插件 - plugins\file_operations\file_deleter.vbs
Option Explicit

Sub file_deleter_Init(pluginObj)
    pluginObj("Name") = "文件删除插件"
    pluginObj("Version") = "1.0"
    pluginObj("ProcessFile") = GetRef("ProcessDelete")
    pluginObj("DeleteEnabled") = False
    pluginObj("SuccessCount") = 0
    pluginObj("Cleanup") = GetRef("file_deleter_Cleanup")
    
    ' 从命令行参数获取开关
    Dim i, arg
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If LCase(arg) = "/delete" Then
            pluginObj("DeleteEnabled") = True
            LogDebug "启用文件删除功能", 1
            Exit For
        End If
    Next
End Sub

Sub ProcessDelete(fileObj, pluginObj)
    If Not pluginObj("DeleteEnabled") Then Exit Sub
    
    ' 执行删除
    On Error Resume Next
    fileObj.Delete True
    If Err.Number = 0 Then
        pluginObj("SuccessCount") = pluginObj("SuccessCount") + 1
        pluginObj("Result") = "成功删除 " & pluginObj("SuccessCount") & " 个文件"
        LogDebug "删除成功: " & fileObj.Path, 3
    Else
        LogDebug "删除失败: " & fileObj.Path & " 错误: " & Err.Description, 1
    End If
    On Error GoTo 0
End Sub

Sub file_deleter_Cleanup(pluginObj)
    LogDebug "文件删除插件清理完成", 3
End Sub
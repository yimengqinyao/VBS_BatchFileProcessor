' 内容替换插件 - plugins\content_tools\content_replacer.vbs
Option Explicit

Sub content_replacer_Init(pluginObj)
    pluginObj("Name") = "内容替换插件"
    pluginObj("Version") = "1.0"
    pluginObj("ProcessFile") = GetRef("ProcessReplace")
    pluginObj("ReplacePattern") = ""
    pluginObj("ReplaceFrom") = ""
    pluginObj("ReplaceTo") = ""
    pluginObj("SuccessCount") = 0
    pluginObj("Cleanup") = GetRef("content_replacer_Cleanup")
    
    ' 从命令行参数获取替换模式
    Dim i, arg, parts
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If Left(LCase(arg), 9) = "/replace:" Then
            pluginObj("ReplacePattern") = Mid(arg, 10)
            parts = Split(pluginObj("ReplacePattern"), "→")
            If UBound(parts) = 1 Then
                pluginObj("ReplaceFrom") = parts(0)
                pluginObj("ReplaceTo") = parts(1)
                LogDebug "设置替换模式: " & pluginObj("ReplaceFrom") & " → " & pluginObj("ReplaceTo"), 1
            End If
            Exit For
        End If
    Next
End Sub

Sub ProcessReplace(fileObj, pluginObj)
    If pluginObj("ReplaceFrom") = "" Then Exit Sub
    
    Dim stream, content
    On Error Resume Next
    Set stream = fso.OpenTextFile(fileObj.Path, 1, False, True)
    content = stream.ReadAll()
    stream.Close
    
    ' 执行替换
    content = Replace(content, pluginObj("ReplaceFrom"), pluginObj("ReplaceTo"), 1, -1, vbTextCompare)
    
    ' 写入替换后的内容
    Set stream = fso.CreateTextFile(fileObj.Path, True, True)
    stream.Write content
    stream.Close
    
    If Err.Number = 0 Then
        pluginObj("SuccessCount") = pluginObj("SuccessCount") + 1
        pluginObj("Result") = "成功替换 " & pluginObj("SuccessCount") & " 个文件"
        LogDebug "替换成功: " & fileObj.Path, 3
    Else
        LogDebug "替换失败: " & fileObj.Path & " 错误: " & Err.Description, 1
    End If
    On Error GoTo 0
End Sub

Sub content_replacer_Cleanup(pluginObj)
    LogDebug "内容替换插件清理完成", 3
End Sub
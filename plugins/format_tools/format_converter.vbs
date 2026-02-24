' 格式转换插件 - plugins\format_tools\format_converter.vbs
Option Explicit

Sub format_converter_Init(pluginObj)
    pluginObj("Name") = "格式转换插件"
    pluginObj("Version") = "1.0"
    pluginObj("ProcessFile") = GetRef("ProcessConvert")
    pluginObj("ConvertPattern") = ""
    pluginObj("ConvertFrom") = ""
    pluginObj("ConvertTo") = ""
    pluginObj("SuccessCount") = 0
    pluginObj("Cleanup") = GetRef("format_converter_Cleanup")
    
    ' 从命令行参数获取转换模式
    Dim i, arg, parts
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If Left(LCase(arg), 8) = "/convert:" Then
            pluginObj("ConvertPattern") = Mid(arg, 9)
            parts = Split(pluginObj("ConvertPattern"), "→")
            If UBound(parts) = 1 Then
                pluginObj("ConvertFrom") = UCase(parts(0))
                pluginObj("ConvertTo") = UCase(parts(1))
                LogDebug "设置格式转换: " & pluginObj("ConvertFrom") & " → " & pluginObj("ConvertTo"), 1
            End If
            Exit For
        End If
    Next
End Sub

Sub ProcessConvert(fileObj, pluginObj)
    If pluginObj("ConvertFrom") = "" Or pluginObj("ConvertTo") = "" Then Exit Sub
    
    ' 检查文件格式
    Dim ext
    ext = UCase(fso.GetExtensionName(fileObj.Path))
    If ext <> pluginObj("ConvertFrom") Then Exit Sub
    
    Dim newPath, stream, content
    newPath = Left(fileObj.Path, Len(fileObj.Path) - Len(ext)) & pluginObj("ConvertTo")
    
    ' 读取文件内容
    On Error Resume Next
    Set stream = fso.OpenTextFile(fileObj.Path, 1, False, True)
    content = stream.ReadAll()
    stream.Close
    
    ' 根据目标格式转换
    Select Case pluginObj("ConvertTo")
        Case "HTML"
            content = "<html><head><meta charset='UTF-8'></head><body><pre>" & content & "</pre></body></html>"
        Case "CSV"
            content = Replace(content, vbTab, ",")
    End Select
    
    ' 写入转换后的内容
    Set stream = fso.CreateTextFile(newPath, True, True)
    stream.Write content
    stream.Close
    
    If Err.Number = 0 Then
        pluginObj("SuccessCount") = pluginObj("SuccessCount") + 1
        pluginObj("Result") = "成功转换 " & pluginObj("SuccessCount") & " 个文件"
        LogDebug "转换成功: " & fileObj.Path & " → " & newPath, 3
    Else
        LogDebug "转换失败: " & fileObj.Path & " 错误: " & Err.Description, 1
    End If
    On Error GoTo 0
End Sub

Sub format_converter_Cleanup(pluginObj)
    LogDebug "格式转换插件清理完成", 3
End Sub
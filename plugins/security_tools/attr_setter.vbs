' 属性设置插件 - plugins\security_tools\attr_setter.vbs
Option Explicit

Sub attr_setter_Init(pluginObj)
    pluginObj("Name") = "属性设置插件"
    pluginObj("Version") = "1.0"
    pluginObj("ProcessFile") = GetRef("ProcessSetAttr")
    pluginObj("Attributes") = ""
    pluginObj("SuccessCount") = 0
    pluginObj("Cleanup") = GetRef("attr_setter_Cleanup")
    
    ' 从命令行参数获取属性
    Dim i, arg
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If Left(LCase(arg), 7) = "/attr:" Then
            pluginObj("Attributes") = Mid(arg, 8)
            LogDebug "设置文件属性: " & pluginObj("Attributes"), 1
            Exit For
        End If
    Next
End Sub

Sub ProcessSetAttr(fileObj, pluginObj)
    If pluginObj("Attributes") = "" Then Exit Sub
    
    Dim attrValue
    attrValue = 0
    
    ' 解析属性
    If InStr(1, pluginObj("Attributes"), "R", vbTextCompare) > 0 Then attrValue = attrValue + 1
    If InStr(1, pluginObj("Attributes"), "H", vbTextCompare) > 0 Then attrValue = attrValue + 2
    If InStr(1, pluginObj("Attributes"), "S", vbTextCompare) > 0 Then attrValue = attrValue + 4
    If InStr(1, pluginObj("Attributes"), "A", vbTextCompare) > 0 Then attrValue = attrValue + 32
    
    ' 设置属性
    On Error Resume Next
    fileObj.Attributes = attrValue
    If Err.Number = 0 Then
        pluginObj("SuccessCount") = pluginObj("SuccessCount") + 1
        pluginObj("Result") = "成功设置属性 " & pluginObj("SuccessCount") & " 个文件"
        LogDebug "属性设置成功: " & fileObj.Path, 3
    Else
        LogDebug "属性设置失败: " & fileObj.Path & " 错误: " & Err.Description, 1
    End If
    On Error GoTo 0
End Sub

Sub attr_setter_Cleanup(pluginObj)
    LogDebug "属性设置插件清理完成", 3
End Sub
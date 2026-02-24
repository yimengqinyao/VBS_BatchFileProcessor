' 内容搜索插件 - plugins\content_tools\content_searcher.vbs
Option Explicit

Sub content_searcher_Init(pluginObj)
    pluginObj("Name") = "内容搜索插件"
    pluginObj("Version") = "1.0"
    pluginObj("ProcessFile") = GetRef("ProcessSearch")
    pluginObj("SearchText") = ""
    pluginObj("MatchCount") = 0
    pluginObj("Cleanup") = GetRef("content_searcher_Cleanup")
    
    ' 从命令行参数获取搜索文本
    Dim i, arg
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If Left(LCase(arg), 5) = "/find:" Then
            pluginObj("SearchText") = Mid(arg, 6)
            LogDebug "设置搜索文本: " & pluginObj("SearchText"), 1
            Exit For
        End If
    Next
End Sub

Sub ProcessSearch(fileObj, pluginObj)
    If pluginObj("SearchText") = "" Then Exit Sub
    
    Dim stream, content, found
    On Error Resume Next
    Set stream = fso.OpenTextFile(fileObj.Path, 1, False, True)
    content = stream.ReadAll()
    stream.Close
    
    found = (InStr(1, content, pluginObj("SearchText"), vbTextCompare) > 0)
    If found Then
        pluginObj("MatchCount") = pluginObj("MatchCount") + 1
        pluginObj("Result") = "找到匹配文件 " & pluginObj("MatchCount") & " 个"
        LogDebug "找到匹配: " & fileObj.Path, 3
    End If
    On Error GoTo 0
End Sub

Sub content_searcher_Cleanup(pluginObj)
    LogDebug "内容搜索插件清理完成", 3
End Sub
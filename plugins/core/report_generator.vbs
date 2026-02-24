' 报告生成插件 - plugins\core\report_generator.vbs
Option Explicit

Sub report_generator_Init(pluginObj)
    pluginObj("Name") = "报告生成插件"
    pluginObj("Version") = "1.0"
    pluginObj("Generate") = GetRef("GenerateReport")
    pluginObj("Cleanup") = GetRef("report_generator_Cleanup")
    LogDebug "报告生成插件初始化完成", 2
End Sub

Sub GenerateReport(pluginObj)
    Dim stream, scanner, md5Plugin
    Set stream = fso.CreateTextFile(logFilePath, True, True)
    
    ' 基本信息
    stream.WriteLine "=== 批量文件处理报告 ==="
    stream.WriteLine "处理时间: " & Now()
    
    ' 获取扫描统计
    If plugins.Exists("file_scanner") Then
        Set scanner = plugins("file_scanner")
        stream.WriteLine "目标目录: " & Join(scanner("TargetFolders"), ", ")
        stream.WriteLine "总处理文件数: " & scanner("TotalFiles")
    End If
    
    ' 文件类型统计
    stream.WriteLine vbCrLf & "=== 文件类型统计 ==="
    Dim ext
    For Each ext In stats.Keys
        stream.WriteLine ext & ": " & stats(ext) & " 个文件"
    Next
    
    ' MD5统计
    If plugins.Exists("md5_calculator") Then
        Set md5Plugin = plugins("md5_calculator")
        stream.WriteLine vbCrLf & "=== MD5统计 ==="
        stream.WriteLine "唯一MD5数量: " & md5Plugin("MD5Dict").Count
    End If
    
    ' 插件结果汇总
    stream.WriteLine vbCrLf & "=== 插件处理结果 ==="
    Dim pluginName
    For Each pluginName In plugins.Keys
        If plugins(pluginName).Exists("Result") Then
            stream.WriteLine plugins(pluginName)("Name") & ": " & plugins(pluginName)("Result")
        End If
    Next
    
    stream.Close
    LogDebug "报告生成完成: " & logFilePath, 1
End Sub

Sub report_generator_Cleanup(pluginObj)
    LogDebug "报告生成插件清理完成", 3
End Sub
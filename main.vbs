' 批量文件处理工具主程序 - main.vbs
Option Explicit

' ====================== 自定义辅助函数 ======================
' 检查指定名称的函数是否存在
Function IsFunction(funcName)
    Dim funcRef
    IsFunction = False
    
    On Error Resume Next
    Set funcRef = GetRef(funcName)
    If Err.Number = 0 Then
        IsFunction = True
    End If
    On Error GoTo 0
End Function

' 检查指定名称的过程是否存在
Function IsSub(subName)
    Dim subRef
    IsSub = False
    
    On Error Resume Next
    Set subRef = GetRef(subName)
    If Err.Number = 0 Then
        IsSub = True
    End If
    On Error GoTo 0
End Function

' 获取缓存键（结合文件路径、大小和修改时间）
Function GetCacheKey(filePath)
    Dim fileObj, size, mtime
    On Error Resume Next
    Set fileObj = fso.GetFile(filePath)
    size = fileObj.Size
    mtime = FormatDateTime(fileObj.DateLastModified, vbGeneralDate)
    GetCacheKey = filePath & "|" & size & "|" & mtime
    On Error GoTo 0
End Function

' 读取缓存
Function ReadCache(cacheKey, cacheFolder)
    ReadCache = ""
    Dim cacheFile, cachePath
    
    If Not fso.FolderExists(cacheFolder) Then
        fso.CreateFolder cacheFolder
        Exit Function
    End If
    
    cachePath = fso.BuildPath(cacheFolder, Replace(cacheKey, "\", "_") & ".cache")
    If fso.FileExists(cachePath) Then
        ' 检查缓存是否过期
        Dim fileObj, currentTime, fileTime, diff
        Set fileObj = fso.GetFile(cachePath)
        currentTime = Now()
        fileTime = fileObj.DateLastModified
        diff = DateDiff("s", fileTime, currentTime)
        
        If diff <= CLng(config("CACHE_DURATION")) Then
            ReadCache = fso.OpenTextFile(cachePath).ReadAll()
        Else
            ' 删除过期缓存
            fso.DeleteFile cachePath
        End If
    End If
End Function

' 写入缓存
Sub WriteCache(cacheKey, cacheValue, cacheFolder)
    Dim cachePath
    
    If Not fso.FolderExists(cacheFolder) Then
        fso.CreateFolder cacheFolder
    End If
    
    cachePath = fso.BuildPath(cacheFolder, Replace(cacheKey, "\", "_") & ".cache")
    Dim stream
    Set stream = fso.CreateTextFile(cachePath, True, True)
    stream.Write cacheValue
    stream.Close
End Sub

' 清理过期缓存
Sub CleanupCache(cacheFolder)
    If Not fso.FolderExists(cacheFolder) Then Exit Sub
    
    Dim folder, files, file, currentTime, diff
    Set folder = fso.GetFolder(cacheFolder)
    Set files = folder.Files
    currentTime = Now()
    
    For Each file In files
        diff = DateDiff("s", file.DateLastModified, currentTime)
        If diff > CLng(config("CACHE_DURATION")) Then
            fso.DeleteFile file.Path, True
            LogDebug "清理过期缓存: " & file.Name, 3
        End If
    Next
End Sub

' ====================== 全局配置 ======================
Const APP_NAME = "批量文件处理工具"
Const APP_VERSION = "v1.0.1"
Const PLUGIN_FOLDER = "plugins\"
Const CONFIG_FILE = "config.ini"
Const DOCS_FOLDER = "docs\"

' 全局对象
Dim fso, shell, wi, config, plugins, stats, cacheEnabled, cacheFolder
Dim debugLevel, logFilePath, debugLogPath

' 程序入口
Sub Main()
    ' 初始化全局对象
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")
    Set wi = CreateObject("WindowsInstaller.Installer") ' 保留原MD5计算方式
    Set config = CreateObject("Scripting.Dictionary")
    Set plugins = CreateObject("Scripting.Dictionary")
    Set stats = CreateObject("Scripting.Dictionary")

    ' 加载配置
    LoadConfig
    
    ' 初始化日志
    InitLogs
    
    ' 初始化缓存
    InitCache
    
    ' 加载所有插件
    LoadAllPlugins
    
    ' 执行主逻辑
    ExecuteMainProcess
    
    ' 清理资源
    Cleanup
End Sub

' ====================== 框架核心函数 ======================
Sub LoadConfig()
    ' 从config.ini加载配置
    If fso.FileExists(CONFIG_FILE) Then
        Dim configStream, line, parts, section
        section = "global"
        Set configStream = fso.OpenTextFile(CONFIG_FILE, 1, False, True)
        Do While Not configStream.AtEndOfStream
            line = Trim(configStream.ReadLine())
            If line <> "" And Left(line, 1) <> ";" Then
                ' 处理节
                If Left(line, 1) = "[" And Right(line, 1) = "]" Then
                    section = Mid(line, 2, Len(line)-2)
                Else
                    parts = Split(line, "=", 2)
                    If UBound(parts) = 1 Then
                        config.Add UCase(section & "_" & Trim(parts(0))), Trim(parts(1))
                    End If
                End If
            End If
        Loop
        configStream.Close
    End If

    ' 设置默认值
    If Not config.Exists("GLOBAL_DEBUG_LEVEL") Then config("GLOBAL_DEBUG_LEVEL") = 2
    If Not config.Exists("GLOBAL_LOG_FILE") Then config("GLOBAL_LOG_FILE") = "FileProcessReport.log"
    If Not config.Exists("GLOBAL_DEBUG_LOG") Then config("GLOBAL_DEBUG_LOG") = "ProcessDebug.log"
    If Not config.Exists("GLOBAL_TARGET_FOLDERS") Then config("GLOBAL_TARGET_FOLDERS") = ""
    If Not config.Exists("GLOBAL_FILE_FILTER") Then config("GLOBAL_FILE_FILTER") = "*.*"
    If Not config.Exists("GLOBAL_SKIP_SUBFOLDERS") Then config("GLOBAL_SKIP_SUBFOLDERS") = "False"
    If Not config.Exists("GLOBAL_MIN_FILE_SIZE") Then config("GLOBAL_MIN_FILE_SIZE") = 0
    If Not config.Exists("GLOBAL_MAX_FILE_SIZE") Then config("GLOBAL_MAX_FILE_SIZE") = 0
    If Not config.Exists("GLOBAL_DATE_FILTER") Then config("GLOBAL_DATE_FILTER") = ""
    
    ' 性能配置默认值
    If Not config.Exists("PERFORMANCE_ENABLE_CACHE") Then config("PERFORMANCE_ENABLE_CACHE") = "True"
    If Not config.Exists("PERFORMANCE_CACHE_FOLDER") Then config("PERFORMANCE_CACHE_FOLDER") = "cache\"
    If Not config.Exists("PERFORMANCE_CACHE_DURATION") Then config("PERFORMANCE_CACHE_DURATION") = "86400"
    If Not config.Exists("PERFORMANCE_MAX_CACHE_ITEMS") Then config("PERFORMANCE_MAX_CACHE_ITEMS") = "1000"

    ' 全局变量赋值
    debugLevel = CInt(config("GLOBAL_DEBUG_LEVEL"))
    logFilePath = config("GLOBAL_LOG_FILE")
    debugLogPath = config("GLOBAL_DEBUG_LOG")
    cacheEnabled = CBool(config("PERFORMANCE_ENABLE_CACHE"))
    cacheFolder = config("PERFORMANCE_CACHE_FOLDER")
End Sub

Sub InitLogs()
    Dim debugStream
    If fso.FileExists(debugLogPath) Then
        Set debugStream = fso.OpenTextFile(debugLogPath, 8, True, True)
        debugStream.WriteLine vbCrLf & "=== 新运行开始 " & Now() & " ==="
    Else
        Set debugStream = fso.CreateTextFile(debugLogPath, True, True)
        debugStream.WriteLine "=== " & APP_NAME & " " & APP_VERSION & " 调试日志 ==="
    End If
    debugStream.Close
End Sub

Sub InitCache()
    If cacheEnabled Then
        ' 创建缓存目录
        If Not fso.FolderExists(cacheFolder) Then
            fso.CreateFolder cacheFolder
            LogDebug "创建缓存目录: " & cacheFolder, 2
        End If
        
        ' 清理过期缓存
        CleanupCache cacheFolder
    End If
End Sub

Sub LoadAllPlugins()
    Dim pluginFolders, folder, path
    pluginFolders = Array("core", "file_operations", "content_tools", "security_tools", "format_tools")
    
    For Each folder In pluginFolders
        path = fso.BuildPath(PLUGIN_FOLDER, folder)
        If fso.FolderExists(path) Then
            LoadPluginsFromFolder path
        End If
    Next
End Sub

Sub LoadPluginsFromFolder(folderPath)
    Dim folder, files, file, pluginName, code
    Set folder = fso.GetFolder(folderPath)
    Set files = folder.Files
    
    For Each file In files
        If LCase(fso.GetExtensionName(file.Name)) = "vbs" Then
            pluginName = Left(file.Name, Len(file.Name)-4)
            code = fso.OpenTextFile(file.Path).ReadAll()
            ExecuteGlobal code
            
            ' 初始化插件
            If IsFunction(pluginName & "_Init") Then
                Dim pluginObj
                Set pluginObj = CreateObject("Scripting.Dictionary")
                Execute pluginName & "_Init pluginObj"
                plugins.Add pluginName, pluginObj
                LogDebug "插件加载成功: " & pluginName, 2
            End If
        End If
    Next
End Sub

Sub ExecuteMainProcess()
    WScript.Echo "=== " & APP_NAME & " " & APP_VERSION & " ==="
    WScript.Echo "开始时间: " & Now()
    WScript.Echo "缓存状态: " & IIf(cacheEnabled, "已启用", "已禁用")
    
    ' 执行核心扫描
    If plugins.Exists("file_scanner") Then
        Dim scanner, processProc
        Set scanner = plugins("file_scanner")
        Set processProc = scanner("Process")  ' 必须使用Set获取过程引用
        processProc scanner  ' 正确的过程调用方式
    End If
    
    ' 生成报告
    If plugins.Exists("report_generator") Then
        Dim reporter, generateProc
        Set reporter = plugins("report_generator")
        Set generateProc = reporter("Generate")
        generateProc reporter
    End If
    
    ' 生成HTML报告
    If plugins.Exists("html_report_generator") Then
        plugins("html_report_generator")("GenerateHTMLReport") plugins("html_report_generator")
    End If

End Sub

Sub Cleanup()
    LogDebug "开始清理资源", 1
    
    ' 清理插件
    Dim pluginName
    For Each pluginName In plugins.Keys
        If plugins(pluginName).Exists("Cleanup") Then
            Dim cleanupProc
            Set cleanupProc = plugins(pluginName)("Cleanup")
            cleanupProc plugins(pluginName)
        End If
    Next
    
    ' 清理缓存
    If cacheEnabled Then
        CleanupCache cacheFolder
    End If
    
    ' 释放全局对象
    Set fso = Nothing
    Set shell = Nothing
    Set wi = Nothing
    Set config = Nothing
    Set plugins = Nothing
    Set stats = Nothing
    
    WScript.Echo "处理完成！报告文件: " & logFilePath
    WScript.Echo "调试日志: " & debugLogPath
End Sub

' ====================== 通用工具函数 ======================
Sub LogDebug(message, level)
    If level <= debugLevel Then
        Dim debugStream
        Set debugStream = fso.OpenTextFile(debugLogPath, 8, True, True)
        debugStream.WriteLine "[" & Now() & "] [" & level & "] " & message
        debugStream.Close
    End If
End Sub

' 启动程序
Main
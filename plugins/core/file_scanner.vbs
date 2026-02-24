' 文件扫描插件 - plugins\core\file_scanner.vbs
Option Explicit

Sub file_scanner_Init(pluginObj)
    pluginObj("Name") = "文件扫描插件"
    pluginObj("Version") = "1.0"
    pluginObj("Process") = GetRef("ProcessScan")
    pluginObj("TargetFolders") = GetTargetFolders()
    pluginObj("FileFilter") = config("FILE_FILTER")
    pluginObj("SkipSubFolders") = CBool(config("SKIP_SUBFOLDERS"))
    pluginObj("MinFileSize") = CLng(config("MIN_FILE_SIZE"))
    pluginObj("MaxFileSize") = CLng(config("MAX_FILE_SIZE"))
    pluginObj("DateFilter") = config("DATE_FILTER")
    pluginObj("TotalFiles") = 0
    pluginObj("Cleanup") = GetRef("file_scanner_Cleanup")
    
    ' 处理命令行参数
    ProcessCommandLineArgs pluginObj
    
    LogDebug "文件扫描插件初始化完成", 2
End Sub

Function GetTargetFolders()
    Dim folders, defaultPath
    If config("TARGET_FOLDERS") <> "" Then
        folders = Split(config("TARGET_FOLDERS"), ",")
    Else
        defaultPath = fso.GetParentFolderName(WScript.ScriptFullName)
        folders = Array(defaultPath)
    End If
    GetTargetFolders = folders
End Function

Sub ProcessScan(pluginObj)
    Dim folderPath
    For Each folderPath In pluginObj("TargetFolders")
        If fso.FolderExists(folderPath) Then
            LogDebug "扫描目录: " & folderPath, 2
            If pluginObj("SkipSubFolders") Then
                ScanNonRecursive folderPath, pluginObj
            Else
                ScanRecursive folderPath, pluginObj
            End If
        Else
            LogDebug "目录不存在: " & folderPath, 1
        End If
    Next
End Sub

' 非递归扫描
Sub ScanNonRecursive(folderPath, pluginObj)
    Dim folder, files, fileObj
    Set folder = fso.GetFolder(folderPath)
    Set files = folder.Files
    
    For Each fileObj In files
        If IsFileMatch(fileObj, pluginObj) Then
            ProcessFile fileObj, pluginObj
        End If
    Next
End Sub

' 递归扫描（改为队列实现，避免栈溢出）
Sub ScanRecursive(folderPath, pluginObj)
    Dim queue, currentFolder, subFolders, subFolder
    Set queue = CreateObject("System.Collections.Queue")
    queue.Enqueue folderPath
    
    Do While queue.Count > 0
        currentFolder = queue.Dequeue()
        Set folder = fso.GetFolder(currentFolder)
        
        ' 处理当前目录文件
        Dim files, fileObj
        Set files = folder.Files
        For Each fileObj In files
            If IsFileMatch(fileObj, pluginObj) Then
                ProcessFile fileObj, pluginObj
            End If
        Next
        
        ' 添加子目录到队列
        Set subFolders = folder.SubFolders
        For Each subFolder In subFolders
            queue.Enqueue subFolder.Path
        Next
    Loop
    
    Set queue = Nothing
End Sub

' 文件匹配判断
Function IsFileMatch(fileObj, pluginObj)
    IsFileMatch = False
    
    ' 大小过滤
    If (pluginObj("MinFileSize") > 0 And fileObj.Size < pluginObj("MinFileSize")) Or _
       (pluginObj("MaxFileSize") > 0 And fileObj.Size > pluginObj("MaxFileSize")) Then
        Exit Function
    End If
    
    ' 时间过滤
    If pluginObj("DateFilter") <> "" And fileObj.DateLastModified < CDate(pluginObj("DateFilter")) Then
        Exit Function
    End If
    
    ' 文件名过滤
    Dim filters, filter
    filters = Split(pluginObj("FileFilter"), ",")
    For Each filter In filters
        If fileObj.Name Like filter Then
            IsFileMatch = True
            Exit Function
        End If
    Next
End Function

' 处理单个文件
Sub ProcessFile(fileObj, pluginObj)
    pluginObj("TotalFiles") = pluginObj("TotalFiles") + 1
    LogDebug "处理文件: " & fileObj.Path, 3
    
    ' 更新文件类型统计
    Dim ext
    ext = UCase(fso.GetExtensionName(fileObj.Path))
    If ext = "" Then ext = "(无扩展名)"
    If stats.Exists(ext) Then
        stats(ext) = stats(ext) + 1
    Else
        stats.Add ext, 1
    End If
    
    ' 调用所有插件处理
    Dim pluginName
    For Each pluginName In plugins.Keys
        If pluginName <> "file_scanner" And pluginName <> "report_generator" Then
            If plugins(pluginName).Exists("ProcessFile") Then
                On Error Resume Next
                Dim processProc
                Set processProc = plugins(pluginName)("ProcessFile")
                processProc fileObj, plugins(pluginName)
                If Err.Number <> 0 Then
                    LogDebug "插件处理错误 [" & pluginName & "]: " & Err.Description, 1
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
    Next
    
    ' 计算MD5 - 修复过程调用方式
    If plugins.Exists("md5_calculator") Then
        Dim md5Plugin, calcProc, updateProc, md5Hash
        Set md5Plugin = plugins("md5_calculator")
        
        ' 获取过程引用
        Set calcProc = md5Plugin("CalculateMD5")
        Set updateProc = md5Plugin("UpdateMD5Dict")
        
        ' 调用过程
        md5Hash = calcProc(fileObj.Path)
        updateProc md5Plugin, fileObj.Path, md5Hash
    End If
End Sub

' 命令行参数处理
Sub ProcessCommandLineArgs(pluginObj)
    Dim arg, i
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        
        ' 调试级别
        If IsNumeric(arg) Then
            debugLevel = CInt(arg)
            LogDebug "设置调试级别: " & debugLevel, 1
        ' 目标目录
        ElseIf Left(LCase(arg), 3) = "/d:" Then
            Dim folders, validFolders(), validCount, folder
            folders = Split(Mid(arg, 4), ",")
            validCount = 0
            
            For Each folder In folders
                If fso.FolderExists(folder) Then
                    ReDim Preserve validFolders(validCount)
                    validFolders(validCount) = folder
                    validCount = validCount + 1
                End If
            Next
            
            If validCount > 0 Then
                pluginObj("TargetFolders") = validFolders
            End If
        ' 文件过滤
        ElseIf Left(LCase(arg), 3) = "/f:" Then
            pluginObj("FileFilter") = Mid(arg, 4)
        ' 忽略子目录
        ElseIf LCase(arg) = "/s" Then
            pluginObj("SkipSubFolders") = True
        ' 其他参数...
        ' 保留原有的所有参数处理逻辑
        End If
    Next
End Sub

Sub file_scanner_Cleanup(pluginObj)
    LogDebug "文件扫描插件清理完成", 3
End Sub
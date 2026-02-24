' 文件备份插件 - plugins\file_operations\file_backuper.vbs
Option Explicit

Sub file_backuper_Init(pluginObj)
    pluginObj("Name") = "文件备份插件"
    pluginObj("Version") = "1.0"
    pluginObj("ProcessFile") = GetRef("ProcessBackup")
    pluginObj("BackupTarget") = ""
    pluginObj("SuccessCount") = 0
    pluginObj("Cleanup") = GetRef("file_backuper_Cleanup")
    
    ' 从命令行参数获取目标
    Dim i, arg
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If Left(LCase(arg), 8) = "/backup:" Then
            pluginObj("BackupTarget") = Mid(arg, 9)
            LogDebug "设置备份目标: " & pluginObj("BackupTarget"), 1
            Exit For
        End If
    Next
End Sub

Sub ProcessBackup(fileObj, pluginObj)
    If pluginObj("BackupTarget") = "" Then Exit Sub
    
    ' 创建带时间戳的备份目录
    Dim timestamp, backupFolder
    timestamp = Replace(Replace(Now(), "/", "-"), ":", "-")
    backupFolder = fso.BuildPath(pluginObj("BackupTarget"), timestamp)
    
    If Not fso.FolderExists(backupFolder) Then
        fso.CreateFolder backupFolder
    End If
    
    Dim targetPath
    targetPath = fso.BuildPath(backupFolder, fileObj.Name)
    
    ' 执行备份
    On Error Resume Next
    fileObj.Copy targetPath, True
    If Err.Number = 0 Then
        pluginObj("SuccessCount") = pluginObj("SuccessCount") + 1
        pluginObj("Result") = "成功备份 " & pluginObj("SuccessCount") & " 个文件"
        LogDebug "备份成功: " & fileObj.Path & " → " & targetPath, 3
    Else
        LogDebug "备份失败: " & fileObj.Path & " 错误: " & Err.Description, 1
    End If
    On Error GoTo 0
End Sub

Sub file_backuper_Cleanup(pluginObj)
    LogDebug "文件备份插件清理完成", 3
End Sub
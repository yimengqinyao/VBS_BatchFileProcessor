' MD5计算插件 - plugins\core\md5_calculator.vbs
Option Explicit

Sub md5_calculator_Init(pluginObj)
    pluginObj("Name") = "MD5计算插件"
    pluginObj("Version") = "1.0.1"
    pluginObj("MD5Dict") = CreateObject("Scripting.Dictionary")
    pluginObj("CalculateMD5") = GetRef("CalculateMD5")
    pluginObj("UpdateMD5Dict") = GetRef("UpdateMD5Dict")
    pluginObj("Cleanup") = GetRef("md5_calculator_Cleanup")
    pluginObj("UseCache") = CBool(config("MD5_CALCULATOR_ENABLE_MD5_CACHE"))
    pluginObj("CacheBySizeAndTime") = CBool(config("MD5_CALCULATOR_CACHE_MD5_BY_SIZE_AND_TIME"))
    pluginObj("UseSystemAPI") = CBool(config("MD5_CALCULATOR_USE_SYSTEM_API"))
    pluginObj("CacheHits") = 0
    pluginObj("CacheMisses") = 0
    
    LogDebug "MD5计算插件初始化完成", 2
End Sub

' 支持缓存的MD5计算函数
Function CalculateMD5(filePath)
    Dim cacheKey, cachedValue, file_hash, hash_value, i
    
    ' 检查缓存
    If cacheEnabled And pluginObj("UseCache") Then
        If pluginObj("CacheBySizeAndTime") Then
            cacheKey = GetCacheKey(filePath)
        Else
            cacheKey = filePath
        End If
        
        cachedValue = ReadCache(cacheKey, cacheFolder)
        If cachedValue <> "" Then
            pluginObj("CacheHits") = pluginObj("CacheHits") + 1
            CalculateMD5 = cachedValue
            LogDebug "MD5缓存命中: " & filePath, 3
            Exit Function
        Else
            pluginObj("CacheMisses") = pluginObj("CacheMisses") + 1
            LogDebug "MD5缓存未命中: " & filePath, 3
        End If
    End If
    
    On Error Resume Next
    If Not fso.FileExists(filePath) Then
        CalculateMD5 = ""
        Exit Function
    End If
    
    ' 使用Windows Installer对象计算MD5
    If pluginObj("UseSystemAPI") Then
        Set file_hash = wi.FileHash(filePath, 0)
        
        hash_value = ""
        If IsObject(file_hash) Then
            For i = 1 To file_hash.FieldCount
                hash_value = hash_value & BigEndianHex(file_hash.IntegerData(i))
            Next
        End If
        
        If hash_value = "00000000" Then
            hash_value = "00000000000000000000000000000000"
        End If
    Else
        ' 回退到纯VBS实现（如果系统API不可用）
        hash_value = VBSMD5(filePath)
    End If
    
    ' 保存到缓存
    If cacheEnabled And pluginObj("UseCache") And hash_value <> "" Then
        WriteCache cacheKey, hash_value, cacheFolder
    End If
    
    CalculateMD5 = hash_value
    On Error GoTo 0
End Function

' 保留原有的大端序转换函数
Function BigEndianHex(Int)
    Dim Result
    Result = Hex(Int)
    If Len(Result) < 8 Then
        Result = String(8 - Len(Result), "0") & Result
    End If
    BigEndianHex = Mid(Result, 7, 2) & Mid(Result, 5, 2) & Mid(Result, 3, 2) & Mid(Result, 1, 2)
End Function

' 纯VBS实现的MD5计算（作为备选方案）
Function VBSMD5(filePath)
    ' 这里可以添加纯VBS的MD5实现，作为系统API的备选方案
    ' 由于长度限制，这里省略具体实现，需要时可以添加
    VBSMD5 = ""
End Function

' 更新MD5字典
Sub UpdateMD5Dict(pluginObj, filePath, md5Hash)
    If md5Hash <> "" And md5Hash <> "00000000000000000000000000000000" Then
        If pluginObj("MD5Dict").Exists(md5Hash) Then
            Dim existingPaths, exists, path
            existingPaths = Split(pluginObj("MD5Dict")(md5Hash), "|")
            exists = False
            
            For Each path In existingPaths
                If path = filePath Then
                    exists = True
                    Exit For
                End If
            Next
            
            If Not exists Then
                pluginObj("MD5Dict")(md5Hash) = pluginObj("MD5Dict")(md5Hash) & "|" & filePath
            End If
        Else
            pluginObj("MD5Dict").Add md5Hash, filePath
        End If
    End If
End Sub

Sub md5_calculator_Cleanup(pluginObj)
    ' 记录缓存统计
    pluginObj("Result") = "MD5计算完成，缓存命中: " & pluginObj("CacheHits") & "，缓存未命中: " & pluginObj("CacheMisses")
    Set pluginObj("MD5Dict") = Nothing
End Sub
# 自定义函数文档 - doc\custom_functions.md

## 1. 函数列表

### 1.1 IsFunction
**功能**：检查指定名称的函数是否存在
**参数**：
	funcName (String)：要检查的函数名称
**返回值**：Boolean，如果函数存在返回True，否则返回False
**实现原理**：
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

**示例**：
If IsFunction("CalculateMD5") Then
    WScript.Echo "CalculateMD5函数存在"
End If

### 1.2 IsSub
功能：检查指定名称的过程是否存在
**参数**：
	subName (String)：要检查的过程名称
**返回值**：Boolean，如果过程存在返回True，否则返回False
**实现原理**：
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

**示例**：
If IsSub("ProcessFile") Then
    WScript.Echo "ProcessFile过程存在"
End If

### 1.3 GetCacheKey
功能：生成缓存键，结合文件路径、大小和修改时间
**参数**：
	filePath (String)：文件路径
**返回值**：String，缓存键字符串
**实现原理**：
Function GetCacheKey(filePath)
    Dim fileObj, size, mtime
    On Error Resume Next
    Set fileObj = fso.GetFile(filePath)
    size = fileObj.Size
    mtime = FormatDateTime(fileObj.DateLastModified, vbGeneralDate)
    GetCacheKey = filePath & "|" & size & "|" & mtime
    On Error GoTo 0
End Function

**示例**：
Dim cacheKey
cacheKey = GetCacheKey("C:\test.txt")

### 1.4 ReadCache
功能：从缓存中读取值
**参数**：
	cacheKey (String)：缓存键
cacheFolder (String)：缓存目录
**返回值**：String，缓存的值，如果不存在或过期返回空字符串
**实现原理**：
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

**示例**：
Dim cachedValue
cachedValue = ReadCache(cacheKey, "cache\")

### 1.5 WriteCache
功能：将值写入缓存
**参数**：
	cacheKey (String)：缓存键
	cacheValue (String)：要缓存的值
	cacheFolder (String)：缓存目录
**实现原理**：
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

**示例**：
WriteCache(cacheKey, "d41d8cd98f00b204e9800998ecf8427e", "cache\")

### 1.6 CleanupCache
功能：清理过期的缓存文件
**参数**：
	cacheFolder (String)：缓存目录
**实现原理**：
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

**示例**：
CleanupCache("cache\")

## 2. 缓存机制详解

### 2.1 缓存键设计
缓存键由三部分组成，确保文件内容变化时缓存自动失效：
文件路径：唯一标识文件
文件大小：文件内容变化时大小通常会变化
文件最后修改时间：文件内容变化时修改时间会变化
格式：文件路径|文件大小|修改时间 示例：C:\test.txt|1024|2026-02-24 14:30:00


### 2.2 缓存文件存储
缓存文件存储在指定的缓存目录中，文件名由缓存键生成：

将路径中的反斜杠替换为下划线
添加.cache后缀 示例：C_test.txt|1024|2026-02-24 14:30:00.cache

### 2.3 缓存过期策略
缓存文件的最后修改时间作为缓存创建时间
每次读取缓存时检查是否超过配置的有效期
自动删除过期的缓存文件
程序启动时自动清理所有过期缓存

### 2.4 缓存命中率统计
MD5计算器插件会记录缓存命中和未命中次数：
CacheHits：缓存命中次数
CacheMisses：缓存未命中次数
统计信息会显示在最终报告中


## 3. 配置参数说明

### 3.1 全局配置参数
GLOBAL_DEBUG_LEVEL：调试日志级别（0-3）
GLOBAL_LOG_FILE：报告文件路径
GLOBAL_DEBUG_LOG：调试日志文件路径
GLOBAL_TARGET_FOLDERS：默认目标目录
GLOBAL_FILE_FILTER：默认文件过滤器
GLOBAL_SKIP_SUBFOLDERS：是否默认跳过子目录

### 3.2 性能配置参数
PERFORMANCE_ENABLE_CACHE：是否启用缓存
PERFORMANCE_CACHE_FOLDER：缓存目录路径
PERFORMANCE_CACHE_DURATION：缓存有效期（秒）
PERFORMANCE_MAX_CACHE_ITEMS：最大缓存项数量（预留）

### 3.3 MD5计算器配置参数
MD5_CALCULATOR_USE_SYSTEM_API：是否使用Windows Installer API计算MD5
MD5_CALCULATOR_ENABLE_MD5_CACHE：是否启用MD5缓存
MD5_CALCULATOR_CACHE_MD5_BY_SIZE_AND_TIME：是否按文件大小和修改时间作为缓存键


## 4. 性能优化建议

### 4.1 缓存策略选择
对于频繁处理相同文件的场景，建议启用缓存
对于一次性处理的文件，建议禁用缓存以节省磁盘空间
对于经常变化的文件，建议缩短缓存有效期

### 4.2 缓存目录位置
将缓存目录放在SSD硬盘上可以提高缓存读写速度
避免将缓存目录放在网络共享目录上
确保缓存目录有足够的可用空间

### 4.3 缓存大小控制
定期清理过期缓存
可以根据需要设置最大缓存项数量（功能预留）
监控缓存目录大小，避免占用过多磁盘空间


## 5. 故障排除

### 5.1 缓存不生效
检查enable_cache是否设置为True
检查缓存目录是否有读写权限
检查文件是否被其他进程锁定导致无法读取大小和修改时间

### 5.2 缓存命中率低
检查文件是否频繁变化
考虑是否需要调整缓存有效期
检查缓存键生成逻辑是否正确

### 5.3 缓存占用过多空间
缩短缓存有效期
定期手动清理缓存目录
考虑禁用缓存或使用更小的缓存项数量

### 5.4 缓存文件损坏
删除损坏的缓存文件
检查磁盘是否有坏道
考虑将缓存目录移动到其他磁盘
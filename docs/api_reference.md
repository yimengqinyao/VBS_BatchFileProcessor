# API参考文档 - doc\api_reference.md

## 1. 全局对象

### 1.1 fso
**类型**：Scripting.FileSystemObject
**说明**：文件系统操作对象，用于文件和目录的创建、删除、读取等操作
**权限**：全局可访问
**示例**：
If fso.FileExists(filePath) Then
    WScript.Echo "文件存在"
End If

### 1.2 shell
**类型**：WScript.Shell
**说明**：Shell操作对象，用于执行系统命令、读取环境变量等
**权限**：全局可访问
**示例**：
shell.Run "notepad.exe"

### 1.3 wi
**类型**：WindowsInstaller.Installer
**说明**：Windows Installer对象，用于计算文件MD5值
**权限**：全局可访问
**示例**：
Set file_hash = wi.FileHash(filePath, 0)

### 1.4 config
**类型**：Scripting.Dictionary
**说明**：配置字典，存储从config.ini加载的配置参数
**权限**：全局可访问
**示例**：
debugLevel = CInt(config("GLOBAL_DEBUG_LEVEL"))

### 1.5 plugins
**类型**：Scripting.Dictionary
**说明**：插件字典，存储所有已加载的插件对象
**权限**：全局可访问
**示例**：
If plugins.Exists("md5_calculator") Then
    ' 调用MD5计算器插件
End If

### 1.6 stats
**类型**：Scripting.Dictionary
**说明**：统计信息字典，存储文件类型统计等信息
**权限**：全局可访问
**示例**：
stats.Add "txt", 100


## 2. 核心函数

### 2.1 Main
**类型**：Sub
**说明**：程序入口函数
**参数**：无
**返回值**：无
**示例**：
Main

### 2.2 LoadConfig
**类型**：Sub
**说明**：加载配置文件到config字典
**参数**：无
**返回值**：无
**示例**：
LoadConfig

### 2.3 InitLogs
**类型**：Sub
**说明**：初始化日志文件
**参数**：无
**返回值**：无
**示例**：
InitLogs

### 2.4 InitCache
**类型**：Sub
**说明**：初始化缓存系统
**参数**：无
**返回值**：无
**示例**：
InitCache

### 2.5 LoadAllPlugins
**类型**：Sub
**说明**：加载所有插件
**参数**：无
**返回值**：无
**示例**：
LoadAllPlugins

### 2.6 ExecuteMainProcess
**类型**：Sub
**说明**：执行主处理流程
**参数**：无
**返回值**：无
**示例**：
ExecuteMainProcess

### 2.7 Cleanup
**类型**：Sub
**说明**：清理资源和缓存
**参数**：无
**返回值**：无
**示例**：
Cleanup

### 2.8 LogDebug
**类型**：Sub
**说明**：记录调试日志
**参数**：
	message (String)：日志消息
	level (Integer)：日志级别（0=无，1=错误，2=信息，3=调试）
**返回值**：无
**示例**：
LogDebug "文件处理完成", 2


## 3. 自定义辅助函数

### 3.1 IsFunction
**类型**：Function
**说明**：检查指定名称的函数是否存在
**参数**：
	funcName (String)：要检查的函数名称
**返回值**：Boolean，如果函数存在返回True，否则返回False
**示例**：
If IsFunction("CalculateMD5") Then
    WScript.Echo "CalculateMD5函数存在"
End If

### 3.2 IsSub
**类型**：Function
**说明**：检查指定名称的过程是否存在
**参数**：
	subName (String)：要检查的过程名称
**返回值**：Boolean，如果过程存在返回True，否则返回False
**示例**：
If IsSub("ProcessFile") Then
    WScript.Echo "ProcessFile过程存在"
End If

### 3.3 GetCacheKey
**类型**：Function
**说明**：生成缓存键，结合文件路径、大小和修改时间
**参数**：
	filePath (String)：文件路径
**返回值**：String，缓存键字符串
**示例**：
Dim cacheKey
cacheKey = GetCacheKey("C:\test.txt")

### 3.4 ReadCache
**类型**：Function
**说明**：从缓存中读取值
**参数**：
	cacheKey (String)：缓存键
	cacheFolder (String)：缓存目录
**返回值**：String，缓存的值，如果不存在或过期返回空字符串
**示例**：
Dim cachedValue
cachedValue = ReadCache(cacheKey, "cache\")

### 3.5 WriteCache
**类型**：Sub
**说明**：将值写入缓存
**参数**：
	cacheKey (String)：缓存键
	cacheValue (String)：要缓存的值
	cacheFolder (String)：缓存目录
**示例**：
WriteCache(cacheKey, "d41d8cd98f00b204e9800998ecf8427e", "cache\")

### 3.6 CleanupCache
**类型**：Sub
**说明**：清理过期的缓存文件
**参数**：
	cacheFolder (String)：缓存目录
**示例**：
CleanupCache("cache\")


## 4. 插件API

### 4.1 插件初始化
**类型**：Function 格式：pluginName_Init(pluginObj)
**说明**：插件初始化函数，每个插件都必须实现此函数
**参数**：
	pluginObj (Scripting.Dictionary)：插件对象字典，用于存储插件的配置、函数和状态
**示例**：
Sub my_plugin_Init(pluginObj)
    pluginObj("Name") = "我的插件"
    pluginObj("Version") = "1.0"
    pluginObj("ProcessFile") = GetRef("ProcessFile")
End Sub

### 4.2 文件处理函数
**类型**：Sub 格式：ProcessFile(fileObj, pluginObj)
**说明**：处理单个文件的函数，插件可以选择实现此函数
**参数**：
	fileObj (Scripting.FileObject)：文件对象
	pluginObj (Scripting.Dictionary)：插件对象字典
**示例**：
Sub ProcessFile(fileObj, pluginObj)
    ' 处理文件逻辑
End Sub

### 4.3 清理函数
**类型**：Sub 格式：pluginName_Cleanup(pluginObj)
**说明**：插件清理函数，用于释放资源
**参数**：
	pluginObj (Scripting.Dictionary)：插件对象字典
**示例**：
Sub my_plugin_Cleanup(pluginObj)
    ' 清理逻辑
End Sub


## 5. 核心插件API

### 5.1 MD5计算器插件

#### 5.1.1 CalculateMD5
**类型**：Function
**说明**：计算文件的MD5值
**参数**：
	filePath (String)：文件路径
**返回值**：String，MD5哈希值
**示例**：
Dim md5Hash
md5Hash = CalculateMD5("C:\test.txt")

#### 5.1.2 UpdateMD5Dict
**类型**：Sub
**说明**：更新MD5字典，存储文件路径和MD5值的映射
**参数**：
	pluginObj (Scripting.Dictionary)：插件对象字典
	filePath (String)：文件路径
	md5Hash (String)：MD5哈希值
**示例**：
UpdateMD5Dict pluginObj, "C:\test.txt", "d41d8cd98f00b204e9800998ecf8427e"

### 5.2 文件扫描插件

#### 5.2.1 ProcessScan
**类型**：Sub
**说明**：扫描目标目录中的文件
**参数**：
	pluginObj (Scripting.Dictionary)：插件对象字典
**示例**：
ProcessScan pluginObj

#### 5.2.2 ScanNonRecursive
**类型**：Sub
**说明**：非递归扫描指定目录
**参数**：
	folderPath (String)：目录路径
	pluginObj (Scripting.Dictionary)：插件对象字典
**示例**：
ScanNonRecursive "C:\test", pluginObj

#### 5.2.3 ScanRecursive
**类型**：Sub
**说明**：递归扫描指定目录及子目录
**参数**：
	folderPath (String)：目录路径
	pluginObj (Scripting.Dictionary)：插件对象字典
**示例**：
ScanRecursive "C:\test", pluginObj

### 5.3 报告生成插件

#### 5.3.1 GenerateReport
**类型**：Sub
**说明**：生成处理报告
**参数**：
	pluginObj (Scripting.Dictionary)：插件对象字典
**示例**：
GenerateReport pluginObj

### 5.4 HTML报告生成器插件

#### 5.4.1 GenerateHTMLReport
**类型**：Sub
**说明**：生成包含重复文件检测的HTML报告
**参数**：
	pluginObj (Scripting.Dictionary)：插件对象字典
**示例**：
plugins("html_report_generator")("GenerateHTMLReport") plugins("html_report_generator")
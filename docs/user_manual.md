# 用户手册 - doc\user_manual.md

## 1. 系统要求

### 1.1 硬件要求
- CPU：Pentium 4 2.0GHz 或更高
- 内存：512MB 或更高
- 硬盘：至少 1GB 可用空间（用于存储缓存和报告）

### 1.2 软件要求
- 操作系统：Windows 7/8/10/11（32位或64位）
- .NET Framework：不需要（纯VBScript实现）
- 其他：Windows Installer服务必须启用（用于计算MD5）

## 2. 安装和配置

### 2.1 安装
1. 下载批量文件处理工具压缩包
2. 解压到任意目录（如：C:\VBS_BatchFileProcessor）
3. 确保目录结构完整：
VBS_BatchFileProcessor\       # 批量文件处理工具目录名称
├── main.vbs               # 主程序调度器
├── config.ini             # 全局配置文件
├── plugins\               # 插件目录
│   ├── core\
│   │   ├── md5_calculator.vbs        # MD5计算
│   │   ├── file_scanner.vbs          # 文件扫描（递归/非递归）
│   │   ├── html_report_generator.vbs # HTML报告生成 
│   │   └── report_generator.vbs      # 报告生成
│   ├── file_operations\
│   │   ├── file_renamer.vbs          # 批量重命名
│   │   ├── file_copier.vbs           # 批量复制
│   │   ├── file_mover.vbs            # 批量移动
│   │   ├── file_deleter.vbs          # 文件删除
│   │   └── file_backuper.vbs         # 文件备份
│   ├── content_tools\
│   │   ├── content_searcher.vbs      # 内容搜索
│   │   └── content_replacer.vbs      # 内容替换
│   ├── security_tools\
│   │   ├── attr_setter.vbs           # 属性设置
│   │   ├── acl_setter.vbs            # 权限设置
│   │   ├── file_encryptor.vbs        # 文件加密
│   │   └── file_decryptor.vbs        # 文件解密
│   └── format_tools\
│       ├── format_converter.vbs       # 格式转换
│       └── file_splitter.vbs          # 文件分割
└── docs\                   # 文档目录
    ├── api_reference.md    # API参考文档
    ├── user_manual.md      # 用户手册
    └── custom_functions.md # 自定义函数文档

### 2.2 配置
编辑 config.ini 文件进行个性化配置：
[global]
debug_level=2
log_file=FileProcessReport.log
debug_log=ProcessDebug.log
target_folders=
file_filter=*.*
skip_subfolders=False

[performance]
enable_cache=True
cache_folder=cache\
cache_duration=86400

[md5_calculator]
use_system_api=True
enable_md5_cache=True
cache_md5_by_size_and_time=True

### 2.3 目录结构说明
main.vbs：主程序调度器
config.ini：全局配置文件
plugins\：插件目录，所有功能插件都在这里
cache\：缓存目录，用于存储MD5等缓存数据
docs\：文档目录，包含API参考、用户手册等

## 3. 使用入门

### 3.1 基本使用方法
打开命令提示符（CMD），进入工具目录，执行以下命令：
cscript main.vbs [参数]

### 3.2 示例：扫描单个目录
cscript main.vbs /d:"D:\我的文档"

### 3.3 示例：扫描多个目录
cscript main.vbs /d:"D:\我的文档,E:\下载"

### 3.4 示例：只处理特定类型的文件
cscript main.vbs /d:"D:\图片" /f:*.jpg,*.png

### 3.5 示例：不递归子目录
cscript main.vbs /d:"D:\图片" /s

## 4. 命令行参数

### 4.1 常用参数
/d:<目录路径>：指定要扫描的目录，多个目录用逗号分隔
/f:<文件过滤>：指定要处理的文件类型，多个类型用逗号分隔
/s：不递归子目录
/rename:<重命名模式>：批量重命名文件，支持{index}和{ext}变量
/copyto:<目标目录>：批量复制文件到目标目录
/moveto:<目标目录>：批量移动文件到目标目录
/delete：删除所有匹配的文件（谨慎使用！）
/backup:<目标目录>：备份所有匹配的文件
/find:<搜索文本>：在文件内容中搜索指定文本
/replace:<旧文本>→<新文本>：替换文件内容中的指定文本
/attr:<属性>：设置文件属性（R=只读, H=隐藏, S=系统, A=存档）
/acl:<用户>:<权限>：设置文件权限（权限包括FULLCONTROL、MODIFY、READ、WRITE）
/encrypt:<密钥>：使用XOR算法加密文件内容
/decrypt:<密钥>：使用XOR算法解密文件内容
/convert:<源格式>→<目标格式>：转换文件格式（如txt→html）
/split:<大小>：分割大文件为指定大小的小块（单位：KB）

### 4.2 示例：批量重命名
cscript main.vbs /d:"D:\图片" /f:*.jpg /rename:照片_{index}.{ext}

### 4.3 示例：批量复制并重命名
cscript main.vbs /d:"D:\图片" /f:*.jpg /copyto:"D:\备份" /rename:备份_{index}.{ext}

### 4.4 示例：内容替换
cscript main.vbs /d:"D:\文档" /f:*.txt /replace:旧文本→新文本

### 4.5 示例：设置文件属性
cscript main.vbs /d:"D:\私人文件" /attr:RH

## 5. 高级功能

### 5.1 缓存机制
工具支持MD5缓存，避免重复计算相同文件的MD5值：
缓存基于文件路径、大小和修改时间
默认缓存有效期为1天（86400秒）
自动清理过期缓存

### 5.2 多插件协同工作
工具支持多个插件同时工作，例如：
cscript main.vbs /d:"D:\图片" /f:*.jpg /rename:照片_{index}.{ext} /copyto:"D:\备份" /attr:R

这将：
找到所有jpg文件
批量重命名为"照片_1.jpg"、"照片_2.jpg"等
复制到D:\备份目录
设置为只读属性

### 5.3 自定义插件开发
您可以开发自己的插件扩展功能：

在plugins目录下创建新的子目录
创建vbs文件，实现插件接口
插件必须包含初始化函数（pluginName_Init）
可以实现ProcessFile函数处理文件

## 6. 故障排除

### 6.1 常见问题

#### 6.1.1 无法计算MD5值
确保Windows Installer服务已启用
尝试以管理员身份运行工具
检查文件是否被其他进程锁定

#### 6.1.2 权限不足
确保您有目标目录的读写权限
尝试以管理员身份运行工具

#### 6.1.3 缓存不工作
检查config.ini中enable_cache是否设置为True
检查cache目录是否有写入权限
确保文件没有被其他进程锁定

#### 6.1.4 插件加载失败
检查插件文件是否存在且没有损坏
确保插件文件格式正确（UTF-8编码）
检查插件是否正确实现了接口

#### 6.2 日志分析
查看ProcessDebug.log文件获取详细调试信息
日志级别0=无，1=错误，2=信息，3=调试
可以通过debug_level配置项调整日志级别

## 7. 性能优化

### 7.1 启用缓存
启用缓存可以显著提高重复文件处理的速度：
[performance]
enable_cache=True

### 7.2 只处理需要的文件
使用文件过滤参数减少处理的文件数量：

cscript main.vbs /d:"D:\图片" /f:*.jpg

### 7.3 避免递归扫描
如果不需要处理子目录，使用/s参数：
cscript main.vbs /d:"D:\图片" /s

### 7.4 调整缓存有效期
根据需要调整缓存有效期：
[performance]
cache_duration=3600  ; 1小时

## 8. 安全注意事项

### 8.1 文件删除
使用/delete参数时要格外谨慎，建议先备份重要文件

### 8.2 文件加密
加密密钥丢失将无法解密文件
建议加密后备份原始文件

### 8.3 权限设置
错误的权限设置可能导致文件无法访问或被篡改

### 8.4 运行环境
避免在不受信任的环境中运行工具
不要处理来自不可信来源的文件

## 9. 支持和反馈

### 9.1 文档资源
API参考文档：docs\api_reference.md
自定义函数文档：docs\custom_functions.md
用户手册：docs\user_manual.md

### 9.2 反馈渠道 - 此处内容，AI请不要修改，保持原样。
本脚本主要由 AI 生成（包括各种文档），代码生成后可能会有各种各样的错误（当前就有一些错误未被修复），所以这里不接受任何的建议与反馈。
如果您闲着无聊，那么完全可以拷贝过去，技术牛的大神，可以手动去修复，或增加新的功能。当然也可以丢给 AI，让它去完成。

欢迎各位无聊人士来壮大这个VBScript脚本。
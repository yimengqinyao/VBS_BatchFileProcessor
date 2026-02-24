' HTML报告生成器 - plugins\core\html_report_generator.vbs

' ==============================================================
' HTML报告生成插件
' 版本: 2.0
' 日期: 2026-02-25
' 功能: 生成带预览、排序、筛选和批量删除功能的HTML报告
' ==============================================================

Class HTMLReportGenerator
    Private fso
    Private reportData
    
    Sub Class_Initialize()
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set reportData = CreateObject("Scripting.Dictionary")
    End Sub
    
    Sub Class_Terminate()
        Set fso = Nothing
        Set reportData = Nothing
    End Sub
    
    ' 设置报告数据
    Sub SetReportData(data)
        Set reportData = data
    End Sub
    
    ' 生成HTML报告
    Function GenerateReport(outputPath)
        On Error Resume Next
        
        ' 创建输出文件
        Dim htmlStream
        Set htmlStream = fso.CreateTextFile(outputPath, True, True)
        If Err.Number <> 0 Then
            GenerateReport = False
            Exit Function
        End If
        
        ' 写入HTML头部
        WriteHTMLHeader htmlStream
        
        ' 写入控制栏
        WriteControls htmlStream
        
        ' 写入MD5分组数据
        WriteMD5Groups htmlStream
        
        ' 写入页脚
        WriteHTMLFooter htmlStream
        
        ' 关闭文件
        htmlStream.Close
        Set htmlStream = Nothing
        
        GenerateReport = True
    End Function
    
    ' 写入HTML头部
    Private Sub WriteHTMLHeader(htmlStream)
        htmlStream.WriteLine "<!DOCTYPE html>"
        htmlStream.WriteLine "<html lang='zh-CN'>"
        htmlStream.WriteLine "<head>"
        htmlStream.WriteLine "    <meta charset='UTF-8'>"
        htmlStream.WriteLine "    <meta name='viewport' content='width=device-width, initial-scale=1.0'>"
        htmlStream.WriteLine "    <title>MD5重复文件检测报告</title>"
        htmlStream.WriteLine "    <style>"
        htmlStream.WriteLine "        body { font-family: 'Microsoft YaHei', Arial, sans-serif; margin: 20px; background-color: #f5f7fa; }"
        htmlStream.WriteLine "        .header { text-align: center; margin-bottom: 30px; padding: 20px; background-color: #fff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }"
        htmlStream.WriteLine "        .header h1 { color: #2c3e50; margin-bottom: 10px; }"
        htmlStream.WriteLine "        .header p { color: #7f8c8d; margin: 5px 0; }"
        htmlStream.WriteLine "        .controls { display: flex; gap: 15px; align-items: center; margin-bottom: 30px; padding: 15px; background-color: #fff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }"
        htmlStream.WriteLine "        .control-group { display: flex; align-items: center; gap: 8px; }"
        htmlStream.WriteLine "        .control-group label { color: #2c3e50; font-weight: 500; }"
        htmlStream.WriteLine "        .control-group select { padding: 6px 12px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; }"
        htmlStream.WriteLine "        .control-group input[type='checkbox'] { width: 16px; height: 16px; cursor: pointer; }"
        htmlStream.WriteLine "        .delete-selected-btn { background-color: #e74c3c; color: white; padding: 8px 16px; border: none; border-radius: 4px; font-size: 14px; font-weight: 500; cursor: pointer; transition: all 0.3s ease; margin-left: auto; }"
        htmlStream.WriteLine "        .delete-selected-btn:hover { background-color: #c0392b; }"
        htmlStream.WriteLine "        .md5-group { background-color: #fff; border-radius: 8px; padding: 20px; margin-bottom: 20px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }"
        htmlStream.WriteLine "        .md5-hash { font-size: 18px; font-weight: bold; color: #34495e; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 1px solid #eee; }"
        htmlStream.WriteLine "        .file-list { list-style: none; padding: 0; margin: 0; }"
        htmlStream.WriteLine "        .file-item { display: flex; justify-content: space-between; align-items: center; padding: 12px 15px; border-radius: 6px; margin-bottom: 8px; background-color: #f8f9fa; transition: all 0.3s ease; }"
        htmlStream.WriteLine "        .file-item:hover { background-color: #eaf2f8; transform: translateX(5px); }"
        htmlStream.WriteLine "        .file-checkbox { margin-right: 12px; width: 18px; height: 18px; cursor: pointer; }"
        htmlStream.WriteLine "        .file-left { display: flex; align-items: center; flex: 1; min-width: 300px; }"
        htmlStream.WriteLine "        .file-icon { width: 24px; height: 24px; margin-right: 12px; font-size: 20px; }"
        htmlStream.WriteLine "        .file-name-container { position: relative; cursor: pointer; margin-right: 12px; }"
        htmlStream.WriteLine "        .file-name { font-weight: 500; color: #2c3e50; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 400px; }"
        htmlStream.WriteLine "        .file-path { display: none; position: absolute; left: 0; bottom: 100%; z-index: 1000; padding: 8px 12px; background-color: white; border-radius: 6px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); color: #7f8c8d; font-size: 12px; white-space: nowrap; margin-bottom: 5px; }"
        htmlStream.WriteLine "        .file-name-container:hover .file-path { display: block; }"
        htmlStream.WriteLine "        .preview-btn { margin-left: 8px; color: #3498db; font-size: 12px; cursor: pointer; position: relative; }"
        htmlStream.WriteLine "        .preview-container { position: fixed; z-index: 1000; pointer-events: none; }"
        htmlStream.WriteLine "        .preview-image { display: none; max-height: 300px; max-width: 600px; padding: 10px; background-color: white; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.2); opacity: 0; transition: opacity 0.3s ease; }"
        htmlStream.WriteLine "        .preview-btn:hover .preview-image { display: block; opacity: 1; }"
        htmlStream.WriteLine "        .preview-error { max-height: 300px; max-width: 600px; padding: 40px; background-color: white; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.2); text-align: center; color: #e74c3c; opacity: 0; transition: opacity 0.3s ease; }"
        htmlStream.WriteLine "        .file-right { display: flex; align-items: center; gap: 20px; }"
        htmlStream.WriteLine "        .file-size { color: #7f8c8d; white-space: nowrap; }"
        htmlStream.WriteLine "        .file-date { color: #7f8c8d; white-space: nowrap; }"
        htmlStream.WriteLine "        .file-actions { display: flex; gap: 10px; }"
        htmlStream.WriteLine "        .action-btn { text-decoration: none; padding: 6px 12px; border-radius: 4px; font-size: 12px; font-weight: 500; transition: all 0.3s ease; }"
        htmlStream.WriteLine "        .btn-locate { background-color: #3498db; color: white; }"
        htmlStream.WriteLine "        .btn-locate:hover { background-color: #2980b9; }"
        htmlStream.WriteLine "        .btn-open { background-color: #2ecc71; color: white; }"
        htmlStream.WriteLine "        .btn-open:hover { background-color: #27ae60; }"
        htmlStream.WriteLine "        .footer { text-align: center; margin-top: 40px; padding: 20px; color: #7f8c8d; font-size: 14px; background-color: #fff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }"
        htmlStream.WriteLine "        .hidden { display: none; }"
        htmlStream.WriteLine "        .empty-state { text-align: center; padding: 40px; color: #7f8c8d; background-color: #f8f9fa; border-radius: 8px; margin-bottom: 20px; }"
        htmlStream.WriteLine "    </style>"
        htmlStream.WriteLine "</head>"
        htmlStream.WriteLine "<body>"
        
        ' 写入报告头部信息
        htmlStream.WriteLine "    <div class='header'>"
        htmlStream.WriteLine "        <h1>MD5重复文件检测报告</h1>"
        htmlStream.WriteLine "        <p>生成时间: " & FormatDateTime(Now, vbLongDate) & " " & FormatDateTime(Now, vbLongTime) & "</p>"
        htmlStream.WriteLine "        <p>总检测文件数: " & reportData("TotalFiles") & " | 唯一MD5数量: " & reportData("UniqueMD5Count") & " | 重复文件组数: " & reportData("DuplicateGroupsCount") & "</p>"
        htmlStream.WriteLine "    </div>"
    End Sub
    
    ' 写入控制栏
    Private Sub WriteControls(htmlStream)
        htmlStream.WriteLine "    <div class='controls'>"
        htmlStream.WriteLine "        <div class='control-group'>"
        htmlStream.WriteLine "            <label for='sort-by'>按大小排序:</label>"
        htmlStream.WriteLine "            <select id='sort-by' onchange='sortFiles()'>"
        htmlStream.WriteLine "                <option value='none'>不排序</option>"
        htmlStream.WriteLine "                <option value='asc'>从小到大</option>"
        htmlStream.WriteLine "                <option value='desc'>从大到小</option>"
        htmlStream.WriteLine "            </select>"
        htmlStream.WriteLine "        </div>"
        htmlStream.WriteLine "        <div class='control-group'>"
        htmlStream.WriteLine "            <label for='filter-type'>文件类型筛选:</label>"
        htmlStream.WriteLine "            <select id='filter-type' onchange='filterFiles()'>"
        htmlStream.WriteLine "                <option value='all'>所有文件</option>"
        htmlStream.WriteLine "                <option value='image'>图片文件</option>"
        htmlStream.WriteLine "                <option value='document'>文档文件</option>"
        htmlStream.WriteLine "                <option value='video'>视频文件</option>"
        htmlStream.WriteLine "                <option value='other'>其他类型</option>"
        htmlStream.WriteLine "            </select>"
        htmlStream.WriteLine "        </div>"
        htmlStream.WriteLine "        <div class='control-group'>"
        htmlStream.WriteLine "            <input type='checkbox' id='select-all' onchange='toggleSelectAll()'>"
        htmlStream.WriteLine "            <label for='select-all'>全选</label>"
        htmlStream.WriteLine "        </div>"
        htmlStream.WriteLine "        <button class='delete-selected-btn' onclick='deleteSelectedFiles()'>删除选中文件</button>"
        htmlStream.WriteLine "    </div>"
    End Sub
    
    ' 写入MD5分组数据
    Private Sub WriteMD5Groups(htmlStream)
        htmlStream.WriteLine "    <div class='md5-groups'>"
        
        Dim md5Groups, md5Hash, fileList, file
        
        Set md5Groups = reportData("MD5Groups")
        For Each md5Hash In md5Groups
            Set fileList = md5Groups(md5Hash)
            
            ' 写入MD5分组标题
            htmlStream.WriteLine "        <div class='md5-group'>"
            htmlStream.WriteLine "            <div class='md5-hash'>MD5: " & md5Hash & " <span style='font-size: 14px; font-weight: normal; color: #7f8c8d;'>(" & fileList.Count & "个重复文件)</span></div>"
            htmlStream.WriteLine "            <ul class='file-list'>"
            
            ' 写入文件列表
            For Each file In fileList
                WriteFileItem htmlStream, file
            Next
            
            htmlStream.WriteLine "            </ul>"
            htmlStream.WriteLine "        </div>"
        Next
        
        htmlStream.WriteLine "    </div>"
    End Sub
    
    ' 写入文件项
    Private Sub WriteFileItem(htmlStream, file)
        Dim fileSize, fileDate, fileType, icon, actionText
        
        fileSize = fso.GetFile(file).Size
        fileDate = fso.GetFile(file).DateLastModified
        fileType = GetFileType(file)
        icon = GetFileIcon(fileType)
        actionText = GetActionText(fileType)
        
        htmlStream.WriteLine "                <li class='file-item' data-size='" & fileSize & "' data-type='" & fileType & "'>"
        htmlStream.WriteLine "                    <input type='checkbox' class='file-checkbox' onchange='updateSelection()' data-file=""" & file & """>"
        htmlStream.WriteLine "                    <div class='file-left'>"
        htmlStream.WriteLine "                        <span class='file-icon'>" & icon & "</span>"
        htmlStream.WriteLine "                        <div class='file-name-container'>"
        htmlStream.WriteLine "                            <span class='file-name'>" & fso.GetFileName(file) & "</span>"
        htmlStream.WriteLine "                            <span class='file-path'>" & file & "</span>"
        htmlStream.WriteLine "                        </div>"
        
        ' 如果是图片文件，添加预览按钮
        If fileType = "image" Then
            htmlStream.WriteLine "                        <span class='preview-btn'>"
            htmlStream.WriteLine "                            预览"
            htmlStream.WriteLine "                            <div class='preview-container'>"
            htmlStream.WriteLine "                                <img class='preview-image' src='file:///" & Replace(file, "\", "/") & "' alt='文件预览' onError='this.style.display=""none""; this.nextElementSibling.style.display=""block"";'>"
            htmlStream.WriteLine "                                <div class='preview-error hidden'>图片加载失败</div>"
            htmlStream.WriteLine "                            </div>"
            htmlStream.WriteLine "                        </span>"
        End If
        
        htmlStream.WriteLine "                    </div>"
        htmlStream.WriteLine "                    <div class='file-right'>"
        htmlStream.WriteLine "                        <span class='file-size'>" & FormatFileSize(fileSize) & "</span>"
        htmlStream.WriteLine "                        <span class='file-date'>" & FormatDateTime(fileDate, vbShortDate) & "</span>"
        htmlStream.WriteLine "                        <div class='file-actions'>"
        htmlStream.WriteLine "                            <a href='#' onclick='locateFile(""" & file & """)' class='action-btn btn-locate'>定位</a>"
        htmlStream.WriteLine "                            <a href='#' onclick='openFile(""" & file & """)' class='action-btn btn-open'>" & actionText & "</a>"
        htmlStream.WriteLine "                        </div>"
        htmlStream.WriteLine "                    </div>"
        htmlStream.WriteLine "                </li>"
    End Sub
    
    ' 写入HTML页脚
    Private Sub WriteHTMLFooter(htmlStream)
        htmlStream.WriteLine "    <div class='footer'>"
        htmlStream.WriteLine "        <p id='copyright'>© 批量文件处理工具</p>"
        htmlStream.WriteLine "        <p>本报告自动生成，请勿手动修改</p>"
        htmlStream.WriteLine "    </div>"
        
        ' 写入JavaScript代码
        htmlStream.WriteLine "    <script>"
        htmlStream.WriteLine "        // 文件操作功能"
        htmlStream.WriteLine "        function locateFile(filePath) {"
        htmlStream.WriteLine "            alert(""定位文件："" + filePath);"
        htmlStream.WriteLine "            // 实际执行时启用以下代码"
        htmlStream.WriteLine "            // var shell = new ActiveXObject('WScript.Shell');"
        htmlStream.WriteLine "            // shell.Run('explorer.exe /select,"" + filePath + ""');"
        htmlStream.WriteLine "        }"
        htmlStream.WriteLine ""
        htmlStream.WriteLine "        function openFile(filePath) {"
        htmlStream.WriteLine "            alert(""打开文件："" + filePath);"
        htmlStream.WriteLine "            // 实际执行时启用以下代码"
        htmlStream.WriteLine "            // var shell = new ActiveXObject('WScript.Shell');"
        htmlStream.WriteLine "            // shell.Run('"" + filePath + ""');"
        htmlStream.WriteLine "        }"
        htmlStream.WriteLine ""
        htmlStream.WriteLine "        // 排序功能"
        htmlStream.WriteLine "        function sortFiles() {"
        htmlStream.WriteLine "            const sortBy = document.getElementById('sort-by').value;"
        htmlStream.WriteLine "            if (sortBy === 'none') {"
        htmlStream.WriteLine "                // 恢复原始顺序"
        htmlStream.WriteLine "                location.reload();"
        htmlStream.WriteLine "                return;"
        htmlStream.WriteLine "            }"
        htmlStream.WriteLine ""
        htmlStream.WriteLine "            const fileLists = document.querySelectorAll('.file-list');"
        htmlStream.WriteLine "            fileLists.forEach(fileList => {"
        htmlStream.WriteLine "                const items = Array.from(fileList.querySelectorAll('.file-item'));"
        htmlStream.WriteLine "                items.sort((a, b) => {"
        htmlStream.WriteLine "                    const sizeA = parseInt(a.dataset.size);"
        htmlStream.WriteLine "                    const sizeB = parseInt(b.dataset.size);"
        htmlStream.WriteLine "                    return sortBy === 'asc' ? sizeA - sizeB : sizeB - sizeA;"
        htmlStream.WriteLine "                });"
        htmlStream.WriteLine ""
        htmlStream.WriteLine "                // 重新排序DOM元素"
        htmlStream.WriteLine "                items.forEach(item => fileList.appendChild(item));"
        htmlStream.WriteLine "            });"
        htmlStream.WriteLine "        }"
        htmlStream.WriteLine ""
        htmlStream.WriteLine "        // 筛选功能"
        htmlStream.WriteLine "        function filterFiles() {"
        htmlStream.WriteLine "            const filterType = document.getElementById('filter-type').value;"
        htmlStream.WriteLine "            const fileItems = document.querySelectorAll('.file-item');"
        htmlStream.WriteLine "            let visibleCount = 0;"
        htmlStream.WriteLine ""
        htmlStream.WriteLine "            fileItems.forEach(item => {"
        htmlStream.WriteLine "                const fileType = item.dataset.type;"
        htmlStream.WriteLine "                if (filterType === 'all' || fileType === filterType) {"
        htmlStream.WriteLine "                    item.classList.remove('hidden');"
        htmlStream.WriteLine "                    visibleCount++;"
        htmlStream.WriteLine "                } else {"
        htmlStream.WriteLine "                    item.classList.add('hidden');"
        htmlStream.WriteLine "                }"
        htmlStream.WriteLine "            });"
        htmlStream.WriteLine ""
        htmlStream.WriteLine "            // 处理空状态"
        htmlStream.WriteLine "            const md5Groups = document.querySelectorAll('.md5-group');"
        htmlStream.WriteLine "            md5Groups.forEach(group => {"
        htmlStream.WriteLine "                const visibleItems = group.querySelectorAll('.file-item:not(.hidden)');"
        htmlStream.WriteLine "                const fileList = group.querySelector('.file-list');"
        htmlStream.WriteLine "                "
        htmlStream.WriteLine "                if (visibleItems.length === 0) {"
        htmlStream.WriteLine "                    // 添加空状态"
        htmlStream.WriteLine "                    if (!fileList.querySelector('.empty-state')) {"
        htmlStream.WriteLine "                        const emptyState = document.createElement('div');"
        htmlStream.WriteLine "                        emptyState.className = 'empty-state';"
        htmlStream.WriteLine "                        emptyState.textContent = '没有符合条件的文件';"
        htmlStream.WriteLine "                        fileList.appendChild(emptyState);"
        htmlStream.WriteLine "                    }"
        htmlStream.WriteLine "                } else {"
        htmlStream.WriteLine "                    // 移除空状态"
        htmlStream.WriteLine "                    const emptyState = fileList.querySelector('.empty-state');"
        htmlStream.WriteLine "                    if (emptyState) {"
        htmlStream.WriteLine "                        emptyState.remove();"
        htmlStream.WriteLine "                    }"
        htmlStream.WriteLine "                }"
        htmlStream.WriteLine "            });"
        htmlStream.WriteLine "        }"
        htmlStream.WriteLine ""
        htmlStream.WriteLine "        // 全选功能"
        htmlStream.WriteLine "        function toggleSelectAll() {"
        htmlStream.WriteLine "            const selectAll = document.getElementById('select-all');"
        htmlStream.WriteLine "            const checkboxes = document.querySelectorAll('.file-checkbox');"
        htmlStream.WriteLine "            checkboxes.forEach(checkbox => {"
        htmlStream.WriteLine "                // 只操作可见文件"
        htmlStream.WriteLine "                const fileItem = checkbox.closest('.file-item');"
        htmlStream.WriteLine "                if (!fileItem.classList.contains('hidden')) {"
        htmlStream.WriteLine "                    checkbox.checked = selectAll.checked;"
        htmlStream.WriteLine "                }"
        htmlStream.WriteLine "            });"
        htmlStream.WriteLine "            updateSelection();"
        htmlStream.WriteLine "        }"
        htmlStream.WriteLine ""
        htmlStream.WriteLine "        // 更新选择状态"
        htmlStream.WriteLine "        function updateSelection() {"
        htmlStream.WriteLine "            const checkboxes = document.querySelectorAll('.file-checkbox');"
        htmlStream.WriteLine "            let selectedCount = 0;"
        htmlStream.WriteLine "            "
        htmlStream.WriteLine "            checkboxes.forEach(checkbox => {"
        htmlStream.WriteLine "                if (checkbox.checked) {"
        htmlStream.WriteLine "                    selectedCount++;"
        htmlStream.WriteLine "                }"
        htmlStream.WriteLine "            });"
        htmlStream.WriteLine ""
        htmlStream.WriteLine "            const deleteBtn = document.querySelector('.delete-selected-btn');"
        htmlStream.WriteLine "            deleteBtn.textContent = selectedCount > 0 ? `删除选中文件 (${selectedCount})` : '删除选中文件';"
        htmlStream.WriteLine "        }"
        htmlStream.WriteLine ""
        htmlStream.WriteLine "        // 删除选中文件"
        htmlStream.WriteLine "        function deleteSelectedFiles() {"
        htmlStream.WriteLine "            const checkedBoxes = document.querySelectorAll('.file-checkbox:checked');"
        htmlStream.WriteLine "            if (checkedBoxes.length === 0) {"
        htmlStream.WriteLine "                alert('请先选择要删除的文件');"
        htmlStream.WriteLine "                return;"
        htmlStream.WriteLine "            }"
        htmlStream.WriteLine ""
        htmlStream.WriteLine "            // 获取选中的文件列表"
        htmlStream.WriteLine "            const filesToDelete = Array.from(checkedBoxes).map(checkbox => checkbox.dataset.file);"
        htmlStream.WriteLine "            "
        htmlStream.WriteLine "            // 二次确认"
        htmlStream.WriteLine "            if (confirm(`确定要删除选中的 ${checkedBoxes.length} 个文件吗？此操作不可恢复！`)) {"
        htmlStream.WriteLine "                // 实际删除逻辑"
        htmlStream.WriteLine "                filesToDelete.forEach(filePath => {"
        htmlStream.WriteLine "                    // 实际执行时启用以下代码"
        htmlStream.WriteLine "                    // const fso = new ActiveXObject('Scripting.FileSystemObject');"
        htmlStream.WriteLine "                    // if (fso.FileExists(filePath)) {"
        htmlStream.WriteLine "                    //     fso.DeleteFile(filePath, true);"
        htmlStream.WriteLine "                    //     // 从DOM中移除"
        htmlStream.WriteLine "                    //     const checkbox = document.querySelector(`.file-checkbox[data-file=""${filePath}""]`);"
        htmlStream.WriteLine "                    //     if (checkbox) {"
        htmlStream.WriteLine "                    //         checkbox.closest('.file-item').remove();"
        htmlStream.WriteLine "                    //     }"
        htmlStream.WriteLine "                    // }"
        htmlStream.WriteLine "                    "
        htmlStream.WriteLine "                    // 模拟删除"
        htmlStream.WriteLine "                    const checkbox = document.querySelector(`.file-checkbox[data-file=""${filePath}""]`);"
        htmlStream.WriteLine "                    if (checkbox) {"
        htmlStream.WriteLine "                        checkbox.closest('.file-item').remove();"
        htmlStream.WriteLine "                    }"
        htmlStream.WriteLine "                });"
        htmlStream.WriteLine "                "
        htmlStream.WriteLine "                alert('文件删除成功');"
        htmlStream.WriteLine "                updateSelection();"
        htmlStream.WriteLine "                // 检查是否有文件组为空"
        htmlStream.WriteLine "                checkEmptyGroups();"
        htmlStream.WriteLine "            }"
        htmlStream.WriteLine "        }"
        htmlStream.WriteLine ""
        htmlStream.WriteLine "        // 检查空的文件组"
        htmlStream.WriteLine "        function checkEmptyGroups() {"
        htmlStream.WriteLine "            const fileLists = document.querySelectorAll('.file-list');"
        htmlStream.WriteLine "            fileLists.forEach(fileList => {"
        htmlStream.WriteLine "                if (fileList.querySelectorAll('.file-item:not(.hidden)').length === 0) {"
        htmlStream.WriteLine "                    fileList.closest('.md5-group').remove();"
        htmlStream.WriteLine "                }"
        htmlStream.WriteLine "            });"
        htmlStream.WriteLine ""
        htmlStream.WriteLine "            // 如果没有文件组了，显示空状态"
        htmlStream.WriteLine "            if (document.querySelectorAll('.md5-group').length === 0) {"
        htmlStream.WriteLine "                const emptyState = document.createElement('div');"
        htmlStream.WriteLine "                emptyState.className = 'empty-state';"
        htmlStream.WriteLine "                emptyState.textContent = '没有重复文件';"
        htmlStream.WriteLine "                document.querySelector('.md5-groups').appendChild(emptyState);"
        htmlStream.WriteLine "            }"
        htmlStream.WriteLine "        }"
        htmlStream.WriteLine ""
        htmlStream.WriteLine "        // 鼠标追踪功能"
        htmlStream.WriteLine "        let currentPreview = null;"
        htmlStream.WriteLine "        "
        htmlStream.WriteLine "        // 鼠标移动事件监听"
        htmlStream.WriteLine "        document.addEventListener('mousemove', function(event) {"
        htmlStream.WriteLine "            if (currentPreview) {"
        htmlStream.WriteLine "                // 设置预览图左下角对齐鼠标（即图片在鼠标右上角显示）"
        htmlStream.WriteLine "                const previewRect = currentPreview.getBoundingClientRect();"
        htmlStream.WriteLine "                "
        htmlStream.WriteLine "                // 初始位置：预览图左下角对齐鼠标"
        htmlStream.WriteLine "                let left = event.clientX;"
        htmlStream.WriteLine "                let top = event.clientY;"
        htmlStream.WriteLine "                "
        htmlStream.WriteLine "                // 检查右侧边界：如果预览图超出屏幕右侧，调整到鼠标左侧显示"
        htmlStream.WriteLine "                if (left + previewRect.width > window.innerWidth) {"
        htmlStream.WriteLine "                    left = event.clientX - previewRect.width;"
        htmlStream.WriteLine "                }"
        htmlStream.WriteLine "                "
        htmlStream.WriteLine "                // 检查底部边界：如果预览图超出屏幕底部，调整到鼠标上方显示"
        htmlStream.WriteLine "                if (top + previewRect.height > window.innerHeight) {"
        htmlStream.WriteLine "                    top = event.clientY - previewRect.height;"
        htmlStream.WriteLine "                }"
        htmlStream.WriteLine "                "
        htmlStream.WriteLine "                // 检查左侧边界：确保预览图不会超出屏幕左侧"
        htmlStream.WriteLine "                if (left < 0) {"
        htmlStream.WriteLine "                    left = 0;"
        htmlStream.WriteLine "                }"
        htmlStream.WriteLine "                "
        htmlStream.WriteLine "                // 检查顶部边界：确保预览图不会超出屏幕顶部"
        htmlStream.WriteLine "                if (top < 0) {"
        htmlStream.WriteLine "                    top = 0;"
        htmlStream.WriteLine "                }"
        htmlStream.WriteLine "                "
        htmlStream.WriteLine "                // 应用位置"
        htmlStream.WriteLine "                currentPreview.style.left = left + 'px';"
        htmlStream.WriteLine "                currentPreview.style.top = top + 'px';"
        htmlStream.WriteLine "            }"
        htmlStream.WriteLine "        });"
        htmlStream.WriteLine "        "
        htmlStream.WriteLine "        // 预览图显示/隐藏事件"
        htmlStream.WriteLine "        document.querySelectorAll('.preview-btn').forEach(btn => {"
        htmlStream.WriteLine "            btn.addEventListener('mouseenter', function() {"
        htmlStream.WriteLine "                const previewImage = this.querySelector('.preview-image');"
        htmlStream.WriteLine "                if (previewImage) {"
        htmlStream.WriteLine "                    currentPreview = previewImage;"
        htmlStream.WriteLine "                    // 立即计算初始位置"
        htmlStream.WriteLine "                    const event = { clientX: window.event.clientX, clientY: window.event.clientY };"
        htmlStream.WriteLine "                    if (currentPreview) {"
        htmlStream.WriteLine "                        const previewRect = currentPreview.getBoundingClientRect();"
        htmlStream.WriteLine "                        let left = event.clientX;"
        htmlStream.WriteLine "                        let top = event.clientY;"
        htmlStream.WriteLine "                        "
        htmlStream.WriteLine "                        if (left + previewRect.width > window.innerWidth) {"
        htmlStream.WriteLine "                            left = event.clientX - previewRect.width;"
        htmlStream.WriteLine "                        }"
        htmlStream.WriteLine "                        "
        htmlStream.WriteLine "                        if (top + previewRect.height > window.innerHeight) {"
        htmlStream.WriteLine "                            top = event.clientY - previewRect.height;"
        htmlStream.WriteLine "                        }"
        htmlStream.WriteLine "                        "
        htmlStream.WriteLine "                        if (left < 0) {"
        htmlStream.WriteLine "                            left = 0;"
        htmlStream.WriteLine "                        }"
        htmlStream.WriteLine "                        "
        htmlStream.WriteLine "                        if (top < 0) {"
        htmlStream.WriteLine "                            top = 0;"
        htmlStream.WriteLine "                        }"
        htmlStream.WriteLine "                        "
        htmlStream.WriteLine "                        currentPreview.style.left = left + 'px';"
        htmlStream.WriteLine "                        currentPreview.style.top = top + 'px';"
        htmlStream.WriteLine "                    }"
        htmlStream.WriteLine "                }"
        htmlStream.WriteLine "            });"
        htmlStream.WriteLine "            "
        htmlStream.WriteLine "            btn.addEventListener('mouseleave', function() {"
        htmlStream.WriteLine "                currentPreview = null;"
        htmlStream.WriteLine "            });"
        htmlStream.WriteLine "        });"
        htmlStream.WriteLine "        "
        htmlStream.WriteLine "        // 动态显示版权年份"
        htmlStream.WriteLine "        document.addEventListener('DOMContentLoaded', function() {"
        htmlStream.WriteLine "            const copyrightElement = document.getElementById('copyright');"
        htmlStream.WriteLine "            const currentYear = new Date().getFullYear();"
        htmlStream.WriteLine "            copyrightElement.innerHTML = '© ' + currentYear + ' 批量文件处理工具';"
        htmlStream.WriteLine "        });"
        htmlStream.WriteLine ""
        htmlStream.WriteLine "        // 初始化"
        htmlStream.WriteLine "        document.addEventListener('DOMContentLoaded', function() {"
        htmlStream.WriteLine "            updateSelection();"
        htmlStream.WriteLine "        });"
        htmlStream.WriteLine "    </script>"
        
        htmlStream.WriteLine "</body>"
        htmlStream.WriteLine "</html>"
    End Sub
    
    ' 获取文件类型
    Private Function GetFileType(filePath)
        Dim ext
        ext = LCase(fso.GetExtensionName(filePath))
        Select Case ext
            ' 图片文件
            Case "jpg", "jpeg", "png", "gif", "bmp", "tiff", "webp", "svg", "raw", "psd"
                GetFileType = "image"
            ' 文档文件
            Case "doc", "docx", "pdf", "txt", "rtf", "md", "odt", "wpd"
                GetFileType = "document"
            ' 表格文件
            Case "xls", "xlsx", "csv", "ods", "numbers"
                GetFileType = "spreadsheet"
            ' 演示文稿
            Case "ppt", "pptx", "odp", "key"
                GetFileType = "presentation"
            ' 压缩包
            Case "zip", "rar", "7z", "tar", "gz", "bz2", "iso", "cab", "arj"
                GetFileType = "archive"
            ' 音频文件
            Case "mp3", "flac", "wma", "ape", "wav", "aac", "ogg", "m4a", "mid"
                GetFileType = "audio"
            ' 视频文件
            Case "mp4", "avi", "mov", "wmv", "flv", "mkv", "webm", "rm", "rmvb", "mpeg", "vob"
                GetFileType = "video"
            ' 代码文件
            Case "js", "html", "css", "java", "py", "c", "cpp", "cs", "php", "rb", "go", "ts"
                GetFileType = "code"
            ' 电子书
            Case "epub", "mobi", "azw", "azw3", "ibooks", "fb2"
                GetFileType = "ebook"
            ' 字体文件
            Case "ttf", "otf", "woff", "woff2", "eot", "fon"
                GetFileType = "font"
            ' 虚拟镜像
            Case "iso", "img", "vhd", "vmdk", "qcow2", "dmg"
                GetFileType = "diskimage"
            ' 数据库文件
            Case "db", "sqlite", "mdb", "accdb", "sql", "bak", "dump"
                GetFileType = "database"
            ' 邮件文件
            Case "eml", "msg", "pst", "ost"
                GetFileType = "email"
            ' 网页文件
            Case "html", "htm", "xhtml", "shtml", "php", "asp", "aspx"
                GetFileType = "webpage"
            ' 程序文件
            Case "exe", "bat", "cmd", "com", "vbs", "js", "ps1"
                GetFileType = "program"
            Case Else
                GetFileType = "other"
        End Select
    End Function
    
    ' 获取文件图标
    Private Function GetFileIcon(fileType)
        Select Case fileType
            Case "image"
                GetFileIcon = "🎨"
            Case "document"
                GetFileIcon = "📃"
            Case "spreadsheet"
                GetFileIcon = "📊"
            Case "presentation"
                GetFileIcon = "🎤"
            Case "archive"
                GetFileIcon = "📦"
            Case "audio"
                GetFileIcon = "🎧"
            Case "video"
                GetFileIcon = "🎬"
            Case "code"
                GetFileIcon = "💻"
            Case "ebook"
                GetFileIcon = "📚"
            Case "font"
                GetFileIcon = "🅰️"
            Case "diskimage"
                GetFileIcon = "📀"
            Case "database"
                GetFileIcon = "🔍"
            Case "email"
                GetFileIcon = "📧"
            Case "webpage"
                GetFileIcon = "🔗"
            Case "program"
                GetFileIcon = "⚙️"
            Case "other"
                GetFileIcon = "📄"
            'Case Else
                'GetFileIcon = "📁"
        End Select
    End Function
    
    ' 获取操作文本
    Private Function GetActionText(fileType)
        Select Case fileType
            Case "image"
                GetActionText = "查看"
            Case "video"
                GetActionText = "播放"
            Case "audio"
                GetActionText = "播放"
            Case "document"
                GetActionText = "打开"
            Case "spreadsheet"
                GetActionText = "打开"
            Case "presentation"
                GetActionText = "打开"
            Case "code"
                GetActionText = "编辑"
            Case "ebook"
                GetActionText = "阅读"
            Case "archive"
                GetActionText = "解压"
            Case "font"
                GetActionText = "安装"
            Case "diskimage"
                GetActionText = "挂载"
            Case "database"
                GetActionText = "打开"
            Case "email"
                GetActionText = "打开"
            Case "webpage"
                GetActionText = "浏览"
            Case "program"
                GetActionText = "运行"
            Case "other"
                GetActionText = "未知"
            'Case Else
                'GetActionText = "打开"
        End Select
    End Function
    
    ' 格式化文件大小
    Private Function FormatFileSize(size)
        Dim units, unitIndex, formattedSize
        
        units = Array("B", "KB", "MB", "GB", "TB")
        unitIndex = 0
        formattedSize = size
        
        Do While formattedSize >= 1024 And unitIndex < UBound(units)
            formattedSize = formattedSize / 1024
            unitIndex = unitIndex + 1
        Loop
        
        FormatFileSize = Round(formattedSize, 2) & " " & units(unitIndex)
    End Function
End Class
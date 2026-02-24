' 文件加密插件 - plugins\security_tools\file_encryptor.vbs
Option Explicit

Sub file_encryptor_Init(pluginObj)
    pluginObj("Name") = "文件加密插件"
    pluginObj("Version") = "1.0"
    pluginObj("ProcessFile") = GetRef("ProcessEncrypt")
    pluginObj("EncryptKey") = ""
    pluginObj("SuccessCount") = 0
    pluginObj("Cleanup") = GetRef("file_encryptor_Cleanup")
    
    ' 从命令行参数获取密钥
    Dim i, arg
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If Left(LCase(arg), 8) = "/encrypt:" Then
            pluginObj("EncryptKey") = Mid(arg, 9)
            LogDebug "设置加密密钥", 1
            Exit For
        End If
    Next
End Sub

Sub ProcessEncrypt(fileObj, pluginObj)
    If pluginObj("EncryptKey") = "" Then Exit Sub
    
    Dim stream, content, encrypted
    On Error Resume Next
    Set stream = fso.OpenTextFile(fileObj.Path, 1, False, True)
    content = stream.ReadAll()
    stream.Close
    
    ' 执行XOR加密
    encrypted = XOREncrypt(content, pluginObj("EncryptKey"))
    
    ' 写入加密内容
    Set stream = fso.CreateTextFile(fileObj.Path, True, True)
    stream.Write encrypted
    stream.Close
    
    If Err.Number = 0 Then
        pluginObj("SuccessCount") = pluginObj("SuccessCount") + 1
        pluginObj("Result") = "成功加密 " & pluginObj("SuccessCount") & " 个文件"
        LogDebug "加密成功: " & fileObj.Path, 3
    Else
        LogDebug "加密失败: " & fileObj.Path & " 错误: " & Err.Description, 1
    End If
    On Error GoTo 0
End Sub

Function XOREncrypt(content, key)
    Dim i, keyIndex, result, charCode, keyCode
    result = ""
    keyIndex = 0
    
    For i = 1 To Len(content)
        charCode = Asc(Mid(content, i, 1))
        keyCode = Asc(Mid(key, (keyIndex Mod Len(key)) + 1, 1))
        result = result & Chr(charCode Xor keyCode)
        keyIndex = keyIndex + 1
    Next
    
    XOREncrypt = result
End Function

Sub file_encryptor_Cleanup(pluginObj)
    LogDebug "文件加密插件清理完成", 3
End Sub
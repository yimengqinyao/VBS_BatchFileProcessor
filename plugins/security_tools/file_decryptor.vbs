' 文件解密插件 - plugins\security_tools\file_decryptor.vbs
Option Explicit

Sub file_decryptor_Init(pluginObj)
    pluginObj("Name") = "文件解密插件"
    pluginObj("Version") = "1.0"
    pluginObj("ProcessFile") = GetRef("ProcessDecrypt")
    pluginObj("DecryptKey") = ""
    pluginObj("SuccessCount") = 0
    pluginObj("Cleanup") = GetRef("file_decryptor_Cleanup")
    
    ' 从命令行参数获取密钥
    Dim i, arg
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If Left(LCase(arg), 9) = "/decrypt:" Then
            pluginObj("DecryptKey") = Mid(arg, 10)
            LogDebug "设置解密密钥", 1
            Exit For
        End If
    Next
End Sub

Sub ProcessDecrypt(fileObj, pluginObj)
    If pluginObj("DecryptKey") = "" Then Exit Sub
    
    ' XOR解密与加密使用相同的函数
    Dim stream, content, decrypted
    On Error Resume Next
    Set stream = fso.OpenTextFile(fileObj.Path, 1, False, True)
    content = stream.ReadAll()
    stream.Close
    
    decrypted = XOREncrypt(content, pluginObj("DecryptKey"))
    
    Set stream = fso.CreateTextFile(fileObj.Path, True, True)
    stream.Write decrypted
    stream.Close
    
    If Err.Number = 0 Then
        pluginObj("SuccessCount") = pluginObj("SuccessCount") + 1
        pluginObj("Result") = "成功解密 " & pluginObj("SuccessCount") & " 个文件"
        LogDebug "解密成功: " & fileObj.Path, 3
    Else
        LogDebug "解密失败: " & fileObj.Path & " 错误: " & Err.Description, 1
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

Sub file_decryptor_Cleanup(pluginObj)
    LogDebug "文件解密插件清理完成", 3
End Sub
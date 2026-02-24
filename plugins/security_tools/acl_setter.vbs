' 权限设置插件 - plugins\security_tools\acl_setter.vbs
Option Explicit

Sub acl_setter_Init(pluginObj)
    pluginObj("Name") = "权限设置插件"
    pluginObj("Version") = "1.0"
    pluginObj("ProcessFile") = GetRef("ProcessSetACL")
    pluginObj("ACLString") = ""
    pluginObj("SuccessCount") = 0
    pluginObj("Cleanup") = GetRef("acl_setter_Cleanup")
    
    ' 从命令行参数获取权限字符串
    Dim i, arg
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If Left(LCase(arg), 8) = "/acl:" Then
            pluginObj("ACLString") = Mid(arg, 9)
            LogDebug "设置文件权限: " & pluginObj("ACLString"), 1
            Exit For
        End If
    Next
End Sub

Sub ProcessSetACL(fileObj, pluginObj)
    If pluginObj("ACLString") = "" Then Exit Sub
    
    Dim parts, user, permission, wmi, objFile, objSec, objAce
    parts = Split(pluginObj("ACLString"), ":")
    If UBound(parts) <> 1 Then Exit Sub
    
    user = parts(0)
    permission = parts(1)
    
    ' 获取WMI对象
    Set wmi = CreateObject("WbemScripting.SWbemLocator")
    Set wmi = wmi.ConnectServer(".", "root\cimv2")
    
    ' 设置权限
    On Error Resume Next
    Set objFile = wmi.Get("Win32_LogicalFileSecuritySetting='" & Replace(fileObj.Path, "\", "\\") & "'")
    objFile.GetSecurityDescriptor objSec
    
    Set objAce = CreateObject("WbemScripting.SWbemObject")
    objAce.Path_.Class = "Win32_ACE"
    objAce.AccessMask = GetAccessMask(permission)
    objAce.AceFlags = 3
    objAce.AceType = 0
    
    Set objAce.Trustee = CreateObject("WbemScripting.SWbemObject")
    objAce.Trustee.Path_.Class = "Win32_Trustee"
    objAce.Trustee.Name = user
    
    objSec.DACL = objSec.DACL + Array(objAce)
    objFile.SetSecurityDescriptor objSec
    
    If Err.Number = 0 Then
        pluginObj("SuccessCount") = pluginObj("SuccessCount") + 1
        pluginObj("Result") = "成功设置权限 " & pluginObj("SuccessCount") & " 个文件"
        LogDebug "权限设置成功: " & fileObj.Path, 3
    Else
        LogDebug "权限设置失败: " & fileObj.Path & " 错误: " & Err.Description, 1
    End If
    On Error GoTo 0
End Sub

Function GetAccessMask(permission)
    Select Case UCase(permission)
        Case "FULLCONTROL"
            GetAccessMask = 2032127
        Case "MODIFY"
            GetAccessMask = 1245631
        Case "READ"
            GetAccessMask = 131209
        Case "WRITE"
            GetAccessMask = 278
        Case Else
            GetAccessMask = 0
    End Select
End Function

Sub acl_setter_Cleanup(pluginObj)
    LogDebug "权限设置插件清理完成", 3
End Sub
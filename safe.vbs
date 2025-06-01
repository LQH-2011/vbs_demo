Option Explicit
Dim WshShell, keyPaths, key, parentKey

Set WshShell = WScript.CreateObject("WScript.Shell")

' 定义需要禁用的注册表项路径（包含父键）
keyPaths = Array( _
    "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskMgr", _
    "HKCU\Software\Policies\Microsoft\Windows\System\DisableCMD", _
    "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools" _
)

' 提示确认
If MsgBox("即将强制禁用系统工具（实验用途），继续吗？", vbYesNo + vbCritical, "警告") = vbNo Then
    WScript.Quit
End If

' 循环处理每个注册表项
For Each key In keyPaths
    ' 提取父键路径（去掉最后一个反斜杠后的内容）
    parentKey = Left(key, InStrRev(key, "\") - 1)
    
    ' 先强制创建父键（即使已存在也不影响）
    On Error Resume Next ' 忽略可能的权限错误
    WshShell.RegWrite parentKey & "\", "" ' 创建父键
    On Error Goto 0
    
    ' 写入实际值（1为禁用）
    WshShell.RegWrite key, 1, "REG_DWORD"
Next

' 强制刷新组策略（需要管理员权限）
On Error Resume Next
WshShell.Run "gpupdate /force /target:user", 0, True
On Error Goto 0

MsgBox "系统工具已锁定！", vbInformation, "操作完成"
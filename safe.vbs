Option Explicit
Dim WshShell, keyPaths, key, parentKey

Set WshShell = WScript.CreateObject("WScript.Shell")

' ������Ҫ���õ�ע�����·��������������
keyPaths = Array( _
    "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskMgr", _
    "HKCU\Software\Policies\Microsoft\Windows\System\DisableCMD", _
    "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools" _
)

' ��ʾȷ��
If MsgBox("����ǿ�ƽ���ϵͳ���ߣ�ʵ����;����������", vbYesNo + vbCritical, "����") = vbNo Then
    WScript.Quit
End If

' ѭ������ÿ��ע�����
For Each key In keyPaths
    ' ��ȡ����·����ȥ�����һ����б�ܺ�����ݣ�
    parentKey = Left(key, InStrRev(key, "\") - 1)
    
    ' ��ǿ�ƴ�����������ʹ�Ѵ���Ҳ��Ӱ�죩
    On Error Resume Next ' ���Կ��ܵ�Ȩ�޴���
    WshShell.RegWrite parentKey & "\", "" ' ��������
    On Error Goto 0
    
    ' д��ʵ��ֵ��1Ϊ���ã�
    WshShell.RegWrite key, 1, "REG_DWORD"
Next

' ǿ��ˢ������ԣ���Ҫ����ԱȨ�ޣ�
On Error Resume Next
WshShell.Run "gpupdate /force /target:user", 0, True
On Error Goto 0

MsgBox "ϵͳ������������", vbInformation, "�������"
'create by Ramon 2017-02-05 ADS
'QQ:289370119

'���������� -Run ��ֻ����
Set objArgs = WScript.Arguments
If objArgs.count =1 Then
    if objArgs(0) = "-Run" Then
        Runfrpc 
    else
        RunAsAdmin
        command
    end if
Else
    sinfo = showOsInfo
    pos = InStr(sinfo,"Microsoft Windows XP")
    if(  pos>0 ) Then
        command
    else
        RunAsAdmin
    end if
end if

'�����û�����Ĳ�������ִ�в���
function command()
    dim m
    m=InputBox("���������Ĳ������룺"& chr(10)&chr(10)&" 0.��ȫ��װ����"& chr(10)&" 1.������"& chr(10)&" 2.����������"& chr(10)&" 3.�Ƴ������� "& chr(10)&" 4.�˳�")
    if m="0" Then
            AutoRun
            Runfrpc
        elseif m="1" Then
            Runfrpc
        elseif m="2" Then
            AutoRun
        elseif m="3" Then
            UnAutoRun
        elseif m="4" Then
            wscript.quit
        else
            Msgbox "����Ĳ������벻��ȷ��"
            command
    end if
End function

'����frpc ������Ϊ�ػ����̣����frpc�����û�н��̣��Զ�����  10��ɨ��һ�� ����ֹwscript���̲��ܽ���
function Runfrpc()
    do while 1=1
        if IsProcess("frpc.exe") = False then 
            Set ws = CreateObject("Wscript.Shell")
            'ws.run "cmd /c frpc",vbhide
            path = createobject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path
            exepath = path&"\frpc.exe"
            inipath = path&"\frpc.ini"
            logpath = path&"\runlog.log"
            'Msgbox exepath&" -c "&inipath
            ws.run("%comspec% /c "&exepath&" -c "&inipath&" >> "&logpath),vbhide
        end if
        wscript.sleep 10000
    loop
End function

'����������
function AutoRun() 
    Set ws = CreateObject("Wscript.Shell")
    currentpath = createobject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path
    path = currentpath&"\frpcRun.vbs -Run"
    ws.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\andsoft",path,"REG_SZ"
End function

'ɾ��������
function UnAutoRun()  
    Set ws = CreateObject("Wscript.Shell") 
    ws.RegDelete "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\andsoft" 
End function


' �Թ���Ա������иýű��ķ���
Sub RunAsAdmin()
  Dim objItems, objItem, strVer, nVer
  Set objItems = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
  For Each objItem In objItems
    strVer = objItem.Version
  Next
  nVer = Split(strVer, ".")(0) & Split(strVer, ".")(1)
  If nVer >= 60 Then
    Dim oShell, oArg, strArgs
    Set oShell = CreateObject("Shell.Application")
    If Not WScript.Arguments.Named.Exists("ADMIN") Then
      For Each oArg In WScript.Arguments
        strArgs = strArgs & " """ & oArg & """"
      Next
      strArgs = strArgs & " /ADMIN:1"
      Call oShell.ShellExecute("WScript.exe", """" & WScript.ScriptFullName & """" & strArgs, "", "runas", 1)
      Set oShell = Nothing
      WScript.Quit(0)
    End If
    Set oShell = Nothing
  End If
End Sub


'������
Function IsProcess(ExeName)
    Dim WMI, Obj, Objs,i
    IsProcess = False
    Set WMI = GetObject("WinMgmts:")
    Set Objs = WMI.InstancesOf("Win32_Process")
    For Each Obj In Objs
        If InStr(UCase(ExeName),UCase(Obj.Description)) <> 0 Then
            IsProcess = True
            Exit For
        End If
    Next
    Set Objs = Nothing
    Set WMI = Nothing
End Function

'���ϵͳx86 x64
Function X86orX64()     
    On Error Resume Next  
    strComputer = "."  
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")  
    Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48)      
    For Each objItem in colItems  
          
       If InStr(objItem.SystemType, "64") <> 0 Then  
            X86orX64 = "x64"         
        Else  
           X86orX64 = "x86"  
        End If  
    Next  
      
End Function  

'������ϵͳ�汾
Function showOsInfo()     
    Dim res  
    On Error Resume Next  
    strComputer = "."  
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")  
    Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)  
    For Each objItem in colItems  
        res =res & "_" &  objItem.Caption         
        'res =res & "_" &  objItem.SystemDrive  
        'res =res & "_" &  objItem.Version  
        'WScript.Echo objItem.OSArchitecture         
    Next
    'WScript.Echo res  
    showOsInfo=res
End Function
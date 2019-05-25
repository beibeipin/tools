'create by Ramon 2017-02-05 ADS
'QQ:289370119

'带参数启动 -Run 则只运行
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

'根据用户输入的操作代码执行操作
function command()
    dim m
    m=InputBox("请输入您的操作代码："& chr(10)&chr(10)&" 0.完全安装运行"& chr(10)&" 1.仅运行"& chr(10)&" 2.加入启动项"& chr(10)&" 3.移除启动项 "& chr(10)&" 4.退出")
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
            Msgbox "输入的操作代码不正确！"
            command
    end if
End function

'启动frpc 本身作为守护进程，监控frpc，如果没有进程，自动启动  10秒扫描一次 ，终止wscript进程才能结束
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

'设置自启动
function AutoRun() 
    Set ws = CreateObject("Wscript.Shell")
    currentpath = createobject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path
    path = currentpath&"\frpcRun.vbs -Run"
    ws.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\andsoft",path,"REG_SZ"
End function

'删除自启动
function UnAutoRun()  
    Set ws = CreateObject("Wscript.Shell") 
    ws.RegDelete "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\andsoft" 
End function


' 以管理员身份运行该脚本的方法
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


'检测进程
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

'检测系统x86 x64
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

'检查操作系统版本
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
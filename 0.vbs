' 该vbs首先判断xp系统，然后判断系统位数，以管理员执行0.bat
Set objShell = CreateObject("Shell.Application")

' 判断操作系统版本和系统位数
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")

For Each objOperatingSystem 在 colOperatingSystems
    strOSVersion = objOperatingSystem.Version
Next

' 检查是否为Windows XP或Windows Server 2003系统
If Left(strOSVersion, 3) = "5.1" Or Left(strOSVersion, 3) = "5.2" Then
    WScript.Quit
End If

' 判断系统位数
Set colComputerSystem = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")

For Each objComputerSystem 在 colComputerSystem
    strOSArch = objComputerSystem.SystemType
Next

' 在64位系统上，SystemType属性值包含"x64"
If InStr(strOSArch, "x64") > 0 Then
    ' 以下为64位系统执行，注意C:\Windows\sysnative
    ' CreateFolder("C:\Windows\sysnative\123q-B")
    RunAsAdmin("C:\Windows\sysnative\cmd.exe /c 1.bat")
Else
    ' 以下为32位系统执行
    ' CreateFolder("C:\Windows\System32\123q-32")
    RunAsAdmin("C:\Windows\system32\cmd.exe /c 1.bat")
End If

' 以下为创建文件夹函数，如果不需要，可以删除CreateFolder子程序
Sub CreateFolder(folderPath)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If Not objFSO.FolderExists(folderPath) Then
        objFSO.CreateFolder(folderPath)
    End If
End Sub

' 以下为以管理员权限运行指定命令的函数,0代表隐藏窗口
Sub RunAsAdmin(command)
    objShell.ShellExecute command, "", "", "runas", 0
End Sub

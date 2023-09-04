# vbsforsystem
vbs以隐藏的方式运行bat脚本，做好前期判断
该工具可以作为前置对工具脚本进行引导

0.任意可以调用vbs的分发工具或压缩工具

1.该脚本首先判断系统版本，区分XP系统【5.1】、2003【5.2】系统或其他系统
如果是5.1或5.2，则自动退出。如果需要针对该系统进行操作，则直接在0.vbs中14行进行编辑
2.如果是其他系统，则开始判断是否为32位或64位系统
其主要是，如果是64位系统，前期调用vbs的软件可能要考虑兼容性，调用的vbs进程为C:\Windows\SysWOW64\wscript.exe，是32位程序，后续调用cmd等可能均为32位，无法在64位系统中引导到C:\Windows\System32\cmd.exe
只会引导到32位的 C:\Windows\SysWOW64\cmd.exe，包括system32目录和注册表等均依赖于cmd的比特数
3.判断系统后，以管理员的方式，并且隐藏窗口的方式调用1.bat
-》objShell.ShellExecute command, "", "", "runas", 0 其中0改为1可以显示cmd的界面
1.bat无需再判断系统版本来进行if跳转，也无需再调用管理员。
-》bat部分C:\Windows\sysnative\cmd.exe /c 1.bat 可以在后面添加 >>999.log 来添加全局执行日志。建议bat添加@echo on来调试
4.创建文件夹相关vbs代码可被删除

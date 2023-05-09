# 仅供上传者Test！！！

- ### XMLHTTP请求测试

|命令|注释|
|:--:|:--:|
|RECALL/ALL|显示需要重新调用的命令|
|SHOW TERMINAL|显示终端信息|
|TYPE *.DAT;*|在当前目录只显示.DAT文件|
|CREATE|建立文本文件.txt|
|CREATE  [TEST.test]|TEST主目录.test子目录|
|SET H 0|重新登陆|
|IPMENU|菜单|
|SET DEF DNFS:[SERVER]|跳转到服务器名为server的目录|
|DELETE|删除|
|DIR|显示目录|
|||
|||
|||

Enable-NetFirewallRule -DisplayGroup "Windows Remote Management"

Enable-PSRemoting -Force

New-NetFirewallRule -DisplayName "Remote Management" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 5985


Start-BitsTransfer
$client = new-object System.Net.WebClient
$client.DownloadFile('https://xxxxxxxx.xxx','D:\xxxxx/xxxx.xxx')

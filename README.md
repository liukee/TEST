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
|Create/dir [.newdir]  在当前目录创建一个目录||
|||
|||
以下是使用cmd命令下载文件的示例代码：

Sub DownloadFile(url As String, destination As String, username As String, password As String)
    Dim cmd As String
    cmd = "net use \\server\share " & password & " /user:" & username
    Shell cmd, vbHide
    cmd = "copy \\server\share\file.txt " & destination
    Shell cmd, vbHide
End Sub
这个代码片段定义了一个名为DownloadFile的子过程，它接受四个参数：要下载的文件的URL、要保存的本地文件路径、用户名和密码。该过程使用net use命令连接到共享文件夹，并使用copy命令将文件复制到本地计算机。

要调用这个过程，可以使用以下代码：

DownloadFile "\\server\share\file.txt", "C:\Downloads\file.txt", "username", "password"
这将下载名为file.txt的文件并将其保存到C:\Downloads\文件夹中。请注意，这种方法需要在本地计算机上设置共享文件夹，并且需要提供正确的用户名和密码进行身份验证

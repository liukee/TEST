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
VBA可以使用cmd命令下载文件，但是使用net use命令下载文件并不是最好的选择。通常，我们使用VBA的内置函数和对象来下载文件，例如使用XMLHTTP对象和FileSystemObject对象。

以下是一个使用XMLHTTP对象下载文件的示例代码：

Sub DownloadFile(url As String, destination As String)
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.send
    If http.Status = 200 Then
        Dim stream As Object
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        stream.Type = 1
        stream.Write http.responseBody
        stream.SaveToFile destination, 2
        stream.Close
    End If
End Sub
这个代码片段定义了一个名为DownloadFile的子过程，它接受两个参数：要下载的文件的URL和要保存的本地文件路径。该过程使用XMLHTTP对象打开URL并下载文件，然后使用FileSystemObject对象将文件保存到本地计算机。

要调用这个过程，可以使用以下代码：

DownloadFile "http://example.com/file.txt", "C:\Downloads\file.txt"
这将下载名为file.txt的文件并将其保存到C:\Downloads\文件夹中

        stream.Type = 1
        stream.Write http.responseBody
        stream.SaveToFile destination, 2
        stream.Close
    End If
End Sub

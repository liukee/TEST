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
｜SET DEF..｜返回上层目录｜
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

//创建FTP连接
Sub FTPConnect()

    Dim FTP As Object
    Set FTP = CreateObject("InetCtls.Inet")

    ' 设置FTP服务器地址和端口号
    FTP.RemoteHost = "ftp.example.com"
    FTP.RemotePort = 21

    ' 设置FTP用户名和密码
    FTP.UserName = "ftp_username"
    FTP.Password = "ftp_password"

    ' 连接FTP服务器
    FTP.Execute "OPEN " & FTP.RemoteHost & " " & FTP.RemotePort

    ' 登录FTP服务器
    FTP.Execute "USER " & FTP.UserName & " " & FTP.Password

    ' 切换到FTP服务器上的目录
    FTP.Execute "CD /ftp_directory"

    ' 下载FTP服务器上的文件
    FTP.Execute "GET ftp_filename local_filename"

    ' 上传文件到FTP服务器
    FTP.Execute "PUT local_filename ftp_filename"

    ' 关闭FTP连接
    FTP.Execute "CLOSE"
    Set FTP = Nothing

End Sub

以上代码使用CreateObject函数创建了一个InetCtls.Inet对象，该对象提供了FTP连接的功能。通过设置RemoteHost、RemotePort、UserName和Password属性，可以指定FTP服务器的地址、端口号、用户名和密码。使用Execute方法执行FTP命令，如OPEN、USER、CD、GET和PUT等，来连接FTP服务器、登录FTP服务器、切换目录、下载文件和上传文件。最后使用CLOSE命令关闭FTP连接。

//扫描枪
Private Sub TextBox1_TextChanged()
    ' 假设 TextBox1 是扫码枪输入的文本框对象
    ' 在此处编写确认事件的代码
    ' 例如，当输入的文本达到一定长度时触发确认事件
    If Len(TextBox1.Value) = 10 Then
        ' 执行确认事件的代码
        MsgBox "确认事件已触发！"
    End If
End Sub

//循环遍历listbox
Dim i As Integer
Dim listBoxItem As Variant

For i = 0 To ListBox1.ListCount - 1
    listBoxItem = ListBox1.List(i)
    ' 在这里处理每个项目的内容
    ' 例如，将每个项目的内容打印到Immediate窗口
    Debug.Print listBoxItem
Next i


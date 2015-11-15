Module Module1

    Sub Main()
        Dim argParam = System.Environment.GetCommandLineArgs()

        Dim MySite, fos, outFile

        On Error Resume Next

        fos = CreateObject("scripting.filesystemobject")
        outFile = fos.OpenTextFile("\\10.7.68.111\Schedule\VBS_TRANS_LOG.txt", 8, True)

        outFile.WriteLine("==========================================================================================")
        outFile.WriteLine("Start Time: " & FormatDateTime(Now()))

        ' 创建cuteftp对象
        MySite = CreateObject("CuteFTPPro.TEConnection")

        ' 初始化参数
        MySite.Host = "10.7.68.100"
        MySite.Protocol = "SFTP"
        MySite.Port = 22
        MySite.Retries = 30
        MySite.Delay = 30
        MySite.MaxConnections = 2
        MySite.TransferType = "AUTO"
        MySite.DataChannel = "DEFAULT"
        MySite.AutoRename = "OFF"
        ' 用户名密码
        MySite.Login = "user"
        MySite.Password = "pwd"
        MySite.SocksInfo = ""
        MySite.ProxyInfo = ""
        ' 链接服务器
        MySite.Connect

        If CBool(MySite.IsConnected) Then
            outFile.WriteLine("Connected to server: " & MySite.Host)
        End If

        If Err.Number > 0 Then
            outFile.WriteLine("Error: " & Err.Description)
            Err.Clear()
        End If

        ' 备份和下载文件
        UploadAndBackupFiles(MySite, fos, outFile, "d:\files\", "d:\bakfiles\" & FormatCurrentDate() & "\")

        DownloadFilesFromSFTP(MySite, outFile, "d:\files", "/downloads/")

        ' Close
        outFile.WriteLine("End Time: " & FormatDateTime(Now()))
        outFile.WriteLine("==========================================================================================" & vbCrLf & vbCrLf)
        outFile.Close
        fos = Nothing
        MySite.Disconnect
        'MySite.Close

    End Sub

    ' 返回时间
    Function FormatCurrentDate()

        FormatCurrentDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
    End Function

    ' 上传和备份文件
    Sub UploadAndBackupFiles(MySite, fs, outFile, localFolder, bakLocalFolder)
        Dim oFolder, oFiles, fileName
        oFolder = fs.GetFolder(localFolder)
        oFiles = oFolder.Files
        For Each file In oFiles
            fileName = file.Name

            ' 检查备份文件夹
            If (Not (MySite.LocalExists(bakLocalFolder))) Then
                MySite.CreateLocalFolder(bakLocalFolder)
            End If

            If (MySite.LocalExists(localFolder & fileName)) Then
                ' Upload file
                If (Not (MySite.RemoteExists("/public/" & fileName))) Then
                    outFile.WriteLine(FormatCurrentDate() & ": [UPLOAD] " & fileName)
                    MySite.Upload(localFolder & fileName, "/public/" & fileName)
                End If
                ' Backup file to bak folder
                outFile.WriteLine(FormatCurrentDate() & ": [BACKUP] " & fileName)
                MySite.LocalRename(localFolder & fileName, bakLocalFolder & fileName)
            End If
        Next
    End Sub

    ' Download files from remote SFTP
    Sub DownloadFilesFromSFTP(MySite, outFile, localFolder, remoteFolder)
        Dim strFileList, strFileName, i, j
        MySite.LocalFolder = localFolder
        MySite.RemoteFolder = remoteFolder
        If CBool(MySite.RemoteExists(MySite.RemoteFolder)) Then
            If CBool(MySite.LocalExists(MySite.LocalFolder)) Then
                ' 获取远程下载目录的文件列表,以"|||"作为分隔符
                MySite.GetList("", "", "%NAME|||")
                strFileList = MySite.GetResult
                If Len(strFileList) <> 0 Then
                    i = 1
                    Do While True
                        j = InStr(i, strFileList, "|||")
                        If j <= 0 Then
                            Exit Do
                        End If
                        strFileName = Mid(strFileList, i, j - i)
                        outFile.WriteLine(FormatCurrentDate() & ": [DOWNLOAD] " & strFileName)
                        MySite.Download(strFileName)
                        outFile.WriteLine(FormatCurrentDate() & ": [*REMOVE*] " & strFileName)
                        MySite.RemoteRemove(strFileName)
                        '加5，因为分隔符"|||"的长度为3个字符，另外还有回车和换行2个字符
                        i = j + 5
                    Loop
                Else
                    outFile.WriteLine("Message! There is no file in remote sftp")
                End If
            Else
                outFile.WriteLine("Error! Local directory doesn't existing")
            End If
        Else
            outFile.WriteLine("Error! Remote directory doesn't existing")
        End If
    End Sub


End Module

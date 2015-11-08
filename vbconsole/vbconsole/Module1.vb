

Module Module1
    ''' <summary>
    ''' 程序主函数
    ''' </summary>
    Sub Main()
        Dim appDir As String
        'appDir = Application.StartupPath
        'appDir = Application.ExecutablePath
        appDir = System.Environment.CurrentDirectory
        appDir = My.Computer.FileSystem.CurrentDirectory
        appDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase
        Console.WriteLine(appDir)

    End Sub

End Module

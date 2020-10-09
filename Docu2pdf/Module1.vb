Imports System.IO
Imports System.Security.AccessControl
Imports FujiXerox.DocuWorks.Toolkit


Module Module1

    Sub Main()

        Dim cmds As String() = System.Environment.GetCommandLineArgs()






        'Dim file1 As String = "C:\Users\toshikanyama\Documents\3X02001.XDW"
        'Dim file2 As String = "C:\Users\toshikanyama\Documents\3X02001.PDF"
        Dim file1 As String = cmds(1)
        Dim file2 As String = cmds(2)
        Dim Ok As Integer = 0

        Dim path1 = System.IO.Path.GetDirectoryName(file1)
        Dim path2 = System.IO.Path.GetDirectoryName(file2)

        Console.WriteLine(file1)
        Console.WriteLine(file2)

        Dim fileInfo1 As New FileInfo(path1)
        Dim fileSec1 As FileSecurity = fileInfo1.GetAccessControl()

        ' アクセス権限をEveryoneに対しフルコントロール許可
        Dim accessRule1 As New FileSystemAccessRule("Everyone", FileSystemRights.FullControl, AccessControlType.Allow)
        fileSec1.AddAccessRule(accessRule1)
        fileInfo1.SetAccessControl(fileSec1)

        ' ファイルの読み取り専用属性を削除
        If (fileInfo1.Attributes And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
            fileInfo1.Attributes = FileAttributes.Normal
        End If

        Dim fileInfo2 As New FileInfo(path2)
        Dim fileSec2 As FileSecurity = fileInfo2.GetAccessControl()

        ' アクセス権限をEveryoneに対しフルコントロール許可
        Dim accessRule2 As New FileSystemAccessRule("Everyone", FileSystemRights.FullControl, AccessControlType.Allow)
        fileSec2.AddAccessRule(accessRule2)
        fileInfo2.SetAccessControl(fileSec2)

        ' ファイルの読み取り専用属性を削除
        If (fileInfo2.Attributes And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
            fileInfo2.Attributes = FileAttributes.Normal
        End If

        Ok = DocuToPdf(file1, file2, 600)




    End Sub


    Public Function DocuToPdf(ByVal file1 As String, ByVal file2 As String, ByVal Dpi As Integer) As Integer

        Dim Handle As Xdwapi.XDW_DOCUMENT_HANDLE = New Xdwapi.XDW_DOCUMENT_HANDLE()
        Dim mode As Xdwapi.XDW_OPEN_MODE_EX = New Xdwapi.XDW_OPEN_MODE_EX()
        With mode
            .Option = Xdwapi.XDW_OPEN_READONLY
            .AuthMode = Xdwapi.XDW_AUTH_NODIALOGUE
        End With

        Dim api_result As Integer = Xdwapi.XDW_OpenDocumentHandle(file1, Handle, mode)

        Dim info As Xdwapi.XDW_DOCUMENT_INFO = New Xdwapi.XDW_DOCUMENT_INFO()
        Xdwapi.XDW_GetDocumentInformation(Handle, info)
        Dim end_page As Integer = info.Pages
        Dim start_page As Integer = 1

        Dim pdf1 As Xdwapi.XDW_IMAGE_OPTION_PDF = New Xdwapi.XDW_IMAGE_OPTION_PDF()

        With pdf1
            .Compress = Xdwapi.XDW_COMPRESS_MRC_NORMAL
            .ConvertMethod = Xdwapi.XDW_CONVERT_MRC_OS
            .EndOfMultiPages = end_page
        End With

        Dim Dpi1 As Integer = Dpi
        Dim Color1 As Integer = Xdwapi.XDW_IMAGE_COLOR
        Dim ImageType1 As Integer = Xdwapi.XDW_IMAGE_PDF
        Dim ex1 As Xdwapi.XDW_IMAGE_OPTION_EX = New Xdwapi.XDW_IMAGE_OPTION_EX()
        With ex1
            .Dpi = Dpi1
            .Color = Color1
            .ImageType = ImageType1
            .DetailOption = pdf1
        End With
        Dim api_result2 As Integer = Xdwapi.XDW_ConvertPageToImageFile(Handle, start_page, file2, ex1)

        'Me.TextBox1.Text = api_result2.ToString

        Xdwapi.XDW_CloseDocumentHandle(Handle)

        DocuToPdf = api_result2
    End Function


End Module

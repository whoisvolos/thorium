Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Public Class FTP

    Public Const FTP_TRANSFER_TYPE_UNKNOWN As Integer = &H0
    Public Const FTP_TRANSFER_TYPE_ASCII As Integer = &H1
    Public Const FTP_TRANSFER_TYPE_BINARY As Integer = &H2
    Public Const INTERNET_DEFAULT_FTP_PORT As Short = 21 ' default for FTP servers
    Public Const INTERNET_SERVICE_FTP As Short = 1
    Public Const INTERNET_FLAG_PASSIVE As Integer = &H8000000 ' used for FTP connections
    Public Const INTERNET_OPEN_TYPE_PRECONFIG As Short = 0 ' use registry configuration
    Public Const INTERNET_OPEN_TYPE_DIRECT As Short = 1 ' direct to net
    Public Const INTERNET_OPEN_TYPE_PROXY As Short = 3 ' via named proxy
    Public Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY As Short = 4 ' prevent using java/script/INS
    Public Const MAX_PATH As Short = 260
    Public Structure FILETIME
        Dim dwLowDateTime As Integer
        Dim dwHighDateTime As Integer
    End Structure
    Public Structure WIN32_FIND_DATA
        Dim dwFileAttributes As Integer
        Dim ftCreationTime As FILETIME
        Dim ftLastAccessTime As FILETIME
        Dim ftLastWriteTime As FILETIME
        Dim nFileSizeHigh As Integer
        Dim nFileSizeLow As Integer
        Dim dwReserved0 As Integer
        Dim dwReserved1 As Integer
        <VBFixedString(MAX_PATH), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=MAX_PATH)> Public cFileName() As Char
        <VBFixedString(14), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=14)> Public cAlternate() As Char
    End Structure
    Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Integer) As Short
    Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Integer, ByVal sServerName As String, ByVal nServerPort As Short, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Integer, ByVal lFlags As Integer, ByVal lContext As Integer) As Integer
    Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Integer, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Integer) As Integer
    Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Integer, ByVal lpszDirectory As String) As Boolean
    Public Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Integer, ByVal lpszCurrentDirectory As String, ByRef lpdwCurrentDirectory As Integer) As Integer
    Public Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Integer, ByVal lpszDirectory As String) As Boolean
    Public Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Integer, ByVal lpszDirectory As String) As Boolean
    Public Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Integer, ByVal lpszFileName As String) As Boolean
    Public Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Integer, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
    Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hConnect As Integer, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Integer, ByVal dwFlagsAndAttributes As Integer, ByVal dwFlags As Integer, ByRef dwContext As Integer) As Boolean
    Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Integer, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Integer, ByVal dwContext As Integer) As Boolean
    Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (ByRef lpdwError As Integer, ByVal lpszBuffer As String, ByRef lpdwBufferLength As Integer) As Boolean
    Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Integer, ByVal lpszSearchFile As String, ByRef lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Integer, ByVal dwContent As Integer) As Integer
    Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Integer, ByRef lpvFindData As WIN32_FIND_DATA) As Integer
    Public Const PassiveConnection As Boolean = True
    Public Sub FTP_test(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        'KPD-Team 2000
        'URL: [url=http://www.allapi.net]http://www.allapi.net[/url]
        'E-Mail: KPDTeam@allapi.net
        Dim hConnection, hOpen As Integer
        Dim sOrgPath As String
        'open an internet connection
        hOpen = InternetOpen("T.H.O.R.I.U.M.", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
        'connect to the FTP server
        hConnection = InternetConnect(hOpen, "ftp.narod.ru", INTERNET_DEFAULT_FTP_PORT, "ark4110", "tasoth", INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
        'create a buffer to store the original directory
        sOrgPath = New String(Chr(0), MAX_PATH)
        'get the directory
        FtpGetCurrentDirectory(hConnection, sOrgPath, Len(sOrgPath))
        'create a new directory 'testing'
        FtpCreateDirectory(hConnection, "testing")
        'set the current directory to 'root/testing'
        FtpSetCurrentDirectory(hConnection, "testing")
        'upload the file 'test.htm'
        FtpPutFile(hConnection, "C:\Data\4110\Tao_TOP\Work_dir\1_wedge_DC.txt", "1_wedge_DC.txt", FTP_TRANSFER_TYPE_UNKNOWN, 0)
        'rename 'test.htm' to 'apiguide.htm'
        FtpRenameFile(hConnection, "1_wedge_DC.txt", "apiguide.htm")
        'enumerate the file list from the current directory ('root/testing')
        FtpGetCurrentDirectoryFileList(hConnection, "*.*")

        'back to initial directory
        FtpSetCurrentDirectory(hConnection, sOrgPath)

        'retrieve the file from the FTP server
        FtpGetFile(hConnection, "SHAENKO_Angular_measurement_system.pdf", "C:\Data\4110\Tao_TOP\Work_dir\SHAENKO_Angular_measurement_system.pdf", False, 0, FTP_TRANSFER_TYPE_UNKNOWN, 0)

        'set the current directory to 'root/testing'
        FtpSetCurrentDirectory(hConnection, "testing")
        'delete the file from the FTP server
        FtpDeleteFile(hConnection, "apiguide.htm")
        'set the current directory back to the root
        FtpSetCurrentDirectory(hConnection, sOrgPath)
        'remove the direcrtory 'testing'
        FtpRemoveDirectory(hConnection, "testing")
        'close the FTP connection
        InternetCloseHandle(hConnection)
        'close the internet connection
        InternetCloseHandle(hOpen)
    End Sub
    Public Function FtpGetCurrentDirectoryFileList(ByRef hConnection As Integer, ByRef search_pattern As String) As String()

        Application.DoEvents()

        Dim pData As WIN32_FIND_DATA = New WIN32_FIND_DATA
        Dim hFind, lRet As Integer

        ' list of files in directory
        Dim file As New Collection

        Dim tmp_Str() As String

        ReDim tmp_Str(-1)

        FtpGetCurrentDirectoryFileList = tmp_Str

        'create a buffer
        pData.cFileName = New String(Chr(0), MAX_PATH)
        'find the first file
        hFind = FtpFindFirstFile(hConnection, search_pattern, pData, 0, 0)
        'if there's no file, then exit sub
        If hFind = 0 Then Exit Function

        file.Add(VB.Left(pData.cFileName, InStr(1, pData.cFileName, New String(Chr(0), 1), CompareMethod.Binary) - 1))
        Do
            'create a buffer
            pData.cFileName = New String(Chr(0), MAX_PATH)
            'find the next file
            lRet = InternetFindNextFile(hFind, pData)
            'if there's no next file, exit do
            If lRet = 0 Then Exit Do

            file.Add(VB.Left(pData.cFileName, InStr(1, pData.cFileName, New String(Chr(0), 1), CompareMethod.Binary) - 1))
        Loop
        'close the search handle
        InternetCloseHandle(hFind)

        ' return array creation

        ReDim tmp_Str(file.Count - 1)

        For i As Integer = 0 To file.Count - 1

            tmp_Str(i) = file.Item(i + 1)

        Next i

        Return tmp_Str

    End Function

    Public Function FtpCurrentDirectoryFileExists(ByRef hConnection As Integer, ByRef Filename As String) As Boolean

        Dim FileList() As String = FtpGetCurrentDirectoryFileList(hConnection, Filename)

        If FileList Is Nothing Then Return False

        If FileList.Length = 0 Then Return False Else Return True

    End Function


    Public Sub FtpCurrentDirectoryFileListDelete(ByRef hConnection As Integer, ByRef File_pattern As String)

        Dim FileList() As String = FtpGetCurrentDirectoryFileList(hConnection, File_pattern)

        For i As Integer = 0 To UBound(FileList)

            FtpDeleteFile(hConnection, FileList(i))

        Next i

    End Sub

    Public Sub FtpSetDirHierarchy(ByRef hConnection As Integer, ByRef Dir_Hierarchy() As String)

        If Dir_Hierarchy Is Nothing Then Exit Sub

        For i As Integer = 0 To UBound(Dir_Hierarchy)

            FtpSetCurrentDirectory(hConnection, Dir_Hierarchy(i))

        Next i

    End Sub
    Public Sub ShowError()
        Dim lErr, lenBuf As Integer
        Dim sErr As String = ""
        'get the required buffer size
        InternetGetLastResponseInfo(lErr, sErr, lenBuf)
        'create a buffer
        sErr = New String(Chr(0), lenBuf)
        'retrieve the last respons info
        InternetGetLastResponseInfo(lErr, sErr, lenBuf)
        'show the last response info
        MsgBox("Error " & CStr(lErr) & ": " & sErr, MsgBoxStyle.OkOnly + MsgBoxStyle.Critical)
    End Sub
End Class
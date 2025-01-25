Attribute VB_Name = "mMain"
Option Explicit

Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Private Const MAX_PATH = 260

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Public Const MODE_INIT As Integer = 1
Public Const MODE_USER_INI As Integer = 2
Public Const MODE_USER As Integer = 3
Public Const MODE_EXEC As Integer = 4
Public Const MODE_WIPE As Integer = 5
Public Const MODE_DELE As Integer = 6
Public Const MODE_VERS_INI As Integer = 7
Public Const MODE_VERS As Integer = 8
Public Const MODE_QUIT As Integer = 9
Public Const MODE_TAKE_INI As Integer = 10
Public Const MODE_TAKE As Integer = 11
Public Const MODE_TAKE_OFF As Integer = 12
Public Const MODE_GIVE_INI_1 As Integer = 13
Public Const MODE_GIVE_INI_2 As Integer = 14
Public Const MODE_GIVE As Integer = 15

Public MyFolder As String
Public PortNo As Integer
Public CurrentMode As Integer
Public InitCount As Integer
Public Buffer As String
Public TransferError As Boolean
Public QuitMode As Boolean
Public FileName As String
Public FileSize As Long

Public Function FileExists(fn As String) As Boolean
    Dim retval As Long, FindData As WIN32_FIND_DATA, retval2 As Long
    retval = FindFirstFile(fn, FindData)
    If retval <> -1 Then FileExists = True
    retval2 = FindClose(retval)
End Function


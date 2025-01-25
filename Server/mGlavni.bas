Attribute VB_Name = "mGlavni"
Option Explicit

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

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public MyFolder As String
Public DistantFolder As String
Public NetMode As Integer

Public PosljednjiDatum As Date
Public dbData As Database
Public rsModeli As Recordset
Public rsKlijenti As Recordset
Public rsKolekcija As Recordset
Public rsNarudzbe As Recordset
Public rsNarudzbeBCKP As Recordset
Public rsTasks As Recordset
Public PortNo As Integer
Public CurrentMode As Integer
Public Buffer As String
Public FileSize As Long
Public MainFormHeight As Single
Public TrenutniFajl As Integer
Public TransferError As Boolean
Public QuitMode As Boolean

Public cFajloviZaSlanje As New Collection

Public Const MODE_INIT As Integer = 1
Public Const MODE_USER_INI As Integer = 2
Public Const MODE_USER As Integer = 3
Public Const MODE_NARUDZBA_INI_1 As Integer = 4
Public Const MODE_NARUDZBA_INI_2 As Integer = 5
Public Const MODE_NARUDZBA As Integer = 6
Public Const MODE_NARUDZBA_OFF As Integer = 7
Public Const MODE_TAKE_INI_1 As Integer = 8
Public Const MODE_TAKE_INI_2 As Integer = 9
Public Const MODE_TAKE As Integer = 11
Public Const MODE_TAKE_OFF As Integer = 12
Public Const MODE_EXEC As Integer = 13
Public Const MODE_QUIT As Integer = 14

Sub Main()

If App.PrevInstance Then End

frmSplash.Show
frmSplash.Refresh

Dim sTMP As String * 255, strlen As Integer

MyFolder = CurDir
'MyFolder = "C:\Poruèivanje obuæe\"
If Right(MyFolder, 1) <> "\" Then MyFolder = MyFolder & "\"
strlen = GetPrivateProfileString("Main", "DistantFolder", "", sTMP, 255, MyFolder & "settings.ini")
DistantFolder = Left(sTMP, strlen)
PortNo = GetPrivateProfileInt("Main", "PortNo", 0, MyFolder & "settings.ini")
If PortNo = 0 Then
    MsgBox "Ne mogu proèitati osnovne postavke!", 16, "Poruèivanje obuæe"
    End
End If
If DistantFolder <> "" Then
    MyFolder = DistantFolder
    If Right(MyFolder, 1) <> "\" Then MyFolder = MyFolder & "\"
End If
PosljednjiDatum = Date

Set dbData = OpenDatabase(MyFolder & "data.mdb", False, False, ";pwd=external")
Set rsModeli = dbData.OpenRecordset("Modeli", dbOpenDynaset, dbSeeChanges)
Set rsKlijenti = dbData.OpenRecordset("Klijenti", dbOpenDynaset, dbSeeChanges)
Set rsKolekcija = dbData.OpenRecordset("Kolekcija", dbOpenDynaset, dbSeeChanges)
Set rsNarudzbe = dbData.OpenRecordset("Narudzbe", dbOpenDynaset, dbSeeChanges)
Set rsNarudzbeBCKP = dbData.OpenRecordset("NarudzbeBCKP", dbOpenDynaset, dbSeeChanges)
Set rsTasks = dbData.OpenRecordset("Tasks", dbOpenDynaset, dbSeeChanges)
frmGlavni.Display
Unload frmSplash
End Sub


Public Function FileExists(fn As String) As Boolean
    Dim retval As Long, FindData As WIN32_FIND_DATA, retval2 As Long
    retval = FindFirstFile(fn, FindData)
    If retval <> -1 Then FileExists = True
    retval2 = FindClose(retval)
End Function

Public Function IsString(sChr As String) As Boolean
Dim lChr As Integer
If Len(sChr) <> 1 Then Exit Function
lChr = Asc(sChr)
If (lChr > 64 And lChr < 91) Or (lChr > 96 And lChr < 123) Then
    IsString = True
Else
    Select Case sChr
        Case "š", "ð", "è", "æ", "ž", "Š", "Ð", "È", "Æ", "Ž"
            IsString = True
        Case Else
            IsString = False
    End Select
End If
End Function

Attribute VB_Name = "Glavni"
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

Public cModeli As New Collection
Public cNarudzbe As New Collection

Public MyFolder As String

Sub Main()
On Error GoTo greška:

If App.PrevInstance Then End

Dim m_lModelID As Long, m_sModel As String, m_sSlika As String, m_iTip As Integer, m_sMatLica As String, m_sMatDjona As String, m_sBoja As String, m_sSortiment As String, m_sCijena As Single, m_dRok As Date

Dim cTMP As New cModel
Dim m_sNB As String, m_sNC As String, c As New cNarudzba

MyFolder = CurDir
'MyFolder = "D:\Poruèivanje obuæe\"
If Right(MyFolder, 1) <> "\" Then MyFolder = MyFolder & "\"

If FileExists(MyFolder & "ConnAgent.new") Then
    Kill MyFolder & "ConnAgent.exe"
    FileCopy MyFolder & "ConnAgent.new", MyFolder & "ConnAgent.exe"
    Kill MyFolder & "ConnAgent.new"
End If

e:

If FileExists(MyFolder & "modeli.dat") Then
    Open MyFolder & "modeli.dat" For Input As #1
        Do Until EOF(1)
            Input #1, m_lModelID, m_sModel, m_sSlika, m_iTip, m_sMatLica, m_sMatDjona, m_sBoja, m_sSortiment, m_sCijena, m_dRok
                cTMP.ModelID = m_lModelID
                cTMP.Model = m_sModel
                cTMP.Slika = m_sSlika
                cTMP.Tip = m_iTip
                cTMP.MatLica = m_sMatLica
                cTMP.MatDjona = m_sMatDjona
                cTMP.Boja = m_sBoja
                cTMP.Sortiment = m_sSortiment
                cTMP.Cijena = m_sCijena
                cTMP.Rok = m_dRok
                
                cModeli.Add cTMP, "ID" & m_lModelID
                Set cTMP = Nothing
        Loop
    Close #1
End If

If FileExists(MyFolder & "narudzbe.dat") Then
    Open MyFolder & "narudzbe.dat" For Input As #1
        Do Until EOF(1)
            Input #1, m_lModelID, m_sNB, m_sNC
                c.ModelID = m_lModelID
                c.NarudzbaBroj = m_sNB
                c.NarudzbaComment = m_sNC
                
                cNarudzbe.Add c, "ID" & m_lModelID
                Set c = Nothing
        Loop
    Close #1
End If
frmGlavni.Display
Exit Sub

greška:
If Err.Number = 75 Then
    GoTo e
Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End If
End Sub

Public Function FileExists(fn As String) As Boolean
    Dim retval As Long, FindData As WIN32_FIND_DATA, retval2 As Long
    retval = FindFirstFile(fn, FindData)
    If retval <> -1 Then FileExists = True
    retval2 = FindClose(retval)
End Function




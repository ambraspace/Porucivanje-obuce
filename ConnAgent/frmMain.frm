VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Poruèivanje obuæe - povezivanje"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm ctlComm 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      ParityReplace   =   0
      RThreshold      =   1
      RTSEnable       =   -1  'True
      BaudRate        =   56000
   End
   Begin VB.Shape ctlProgressBorder 
      Height          =   255
      Left            =   120
      Top             =   720
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Shape ctlProgress 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   720
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Shape ctlProgressBack 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   120
      Top             =   720
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label lblPercent 
      Alignment       =   1  'Right Justify
      Caption         =   "0%"
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOpis 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6135
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' API za FileExists funkciju
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

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

' API za èitanje .INI fajlova
Private Declare Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private MyFolder As String  'varijabla koja sadrži ime foldera u kojem se ConnAgent.exe nalazi
Private User As String      'varijabla za ime korisnika
Private PortNo As Integer   'broj porta za komunikaciju

Private CurrentMode As Integer  'indikator trenutnog naèina rada (moda), tj. faze prenosa
Private FileName As String      'naziv fajla koji se prima ili šalje
Private FileSize As Long        'velièina fajla koji se prima ili šalje
Private Buffer As String        'bafer za primanje komandi od servera (server je druga komponenta cijelog paketa pomoæu koje se upravlja ConnAgent.exe-om
Private QuitMode As Boolean     'indikator prekida
Private TransferError As Boolean    'indikator greške u prenosu

'konstante - indikatori naèina rada (faza prenosa)
Private Const MODE_LISTEN As Integer = 100          'ConnAgent osluškoje liniju i èeka komandu
Private Const MODE_USER As Integer = 101
Private Const MODE_GIVE_INI As Integer = 102        'uvodna faza primanja fajla (komanda GIVE)
Private Const MODE_GIVE As Integer = 103            'glavna faza primanja fajla (komanda GIVE)
Private Const MODE_NOFILE As Integer = 104          'ukoliko zahtjevani fajl ne postoji ConnAgent prelazi u ovu fazu
Private Const MODE_TAKE_INI As Integer = 105        'uvodna faza za slanje fajla (komanda TAKE)
Private Const MODE_TAKE As Integer = 106            'glavna faza za slanje fajla (komanda TAKE)
Private Const MODE_VERS As Integer = 107
Private Const MODE_WAITFORCALL As Integer = 108     'ukoliko pukne veza ConnAgent u ovoj fazi oèekuje ponovni poziv

Private Const VERSION As Single = 1.3

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = vbKeyEnd Then
    If Me.ctlComm.PortOpen Then Me.ctlComm.PortOpen = False
    End
End If
End Sub

Private Sub Form_Load()
If FindWindow(vbNullString, "Poruèivanje obuæe") <> 0 Then
    MsgBox "Ugasite prozor ""Poruèivanje obuæe""," & vbCrLf & _
        "pa ponovo pokrenite ""Povezivanje""!", 16, "Greška"
    End
End If

MyFolder = CurDir
'MyFolder = "D:\Poruèivanje obuæe"
If Right(MyFolder, 1) <> "\" Then MyFolder = MyFolder & "\"

Dim strlen As Integer, tmpUser As String * 50
strlen = GetPrivateProfileString("Main", "User", "Nepoznat", tmpUser, 50, MyFolder & "settings.ini")
User = Left(tmpUser, strlen)

PortNo = GetPrivateProfileInt("Main", "PortNo", 0, MyFolder & "settings.ini")

If PortNo = 0 Or User = "Nepoznat" Then
    MsgBox "Ne mogu proèitati osnovne postavke!", 16, "Greška"
    End
End If

Me.Show
Me.ctlComm.CommPort = PortNo
Me.ctlComm.PortOpen = True

DoConnect

End Sub

Private Sub DoConnect()
Dim k As Date

Me.lblTitle = "Povezivanje u toku..."
Me.ctlComm.Output = "AT&FV1E0X3A" & vbCr

k = Now + #12:00:50 AM#
Do Until k < Now
    DoEvents
    If Me.ctlComm.CDHolding Then Exit Do
Loop

If Me.ctlComm.CDHolding Then
    Me.lblTitle = "Povezivanje je uspjelo! Prenos podataka u toku..."
    Me.ctlComm.InBufferCount = 0
    Me.ctlComm.OutBufferCount = 0
    Me.lblOpis = "Èekam komandu..."
    CurrentMode = MODE_LISTEN
    CheckConnection
Else
    Me.lblTitle = "Povezivanje nije uspjelo!"
    Me.ctlComm.PortOpen = False
    
    k = Now + #12:00:03 AM#
    Do Until k < Now
        DoEvents
    Loop
    Me.ctlComm.PortOpen = True
    CurrentMode = MODE_WAITFORCALL
    Me.lblTitle = "Èekam poziv..."
End If
End Sub

Private Sub ctlComm_OnComm()
Dim sTMP As String, bKomad() As Byte, bKomadSize As Long, k As Date

If Me.ctlComm.CommEvent = comEvReceive Then
    Select Case CurrentMode
        Case 0
            sTMP = Me.ctlComm.Input
        Case MODE_WAITFORCALL
            Buffer = Buffer & Me.ctlComm.Input
            If InStr(Buffer, "RING") > 0 Then
                CurrentMode = 0
                Buffer = ""
                DoConnect
            End If
        Case MODE_LISTEN
            Buffer = Buffer & Me.ctlComm.Input
            If Right(Buffer, 1) = "*" Then
                sTMP = Left(Buffer, Len(Buffer) - 1)
                Buffer = ""
                If DoCommand(sTMP) Then
                    Me.ctlComm.Output = "*"
                Else
                    Me.ctlComm.Output = "|"
                End If
                If QuitMode Then Me.ctlComm.PortOpen = False
            End If
        Case MODE_USER
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                Me.lblOpis = "Èekam komandu..."
                CurrentMode = MODE_LISTEN
                Me.ctlComm.Output = User & "*"
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_NOFILE
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                CurrentMode = MODE_LISTEN
                Me.ctlComm.Output = "ERR*"
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_GIVE_INI
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                Open FileName For Binary As #1
                Me.lblOpis = "Šaljem: " & FileName & " (" & Round(FileSize / 1024, 1) & "KB)..."
                Me.ctlProgress.Width = 0
                Me.ctlProgress.Visible = True
                Me.ctlProgressBack.Visible = True
                Me.ctlProgressBorder.Visible = True
                Me.lblPercent = "0%"
                Me.lblPercent.Visible = True
                CurrentMode = MODE_GIVE
                Me.ctlComm.Output = FileSize & "OK*"
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_GIVE
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                Me.ctlProgress.Width = (Seek(1) - 1) / FileSize * Me.ctlProgressBack.Width
                Me.lblPercent = Int((Seek(1) - 1) / FileSize * 100) & "%"
                bKomadSize = FileSize - Seek(1) + 1
                If bKomadSize > 512 Then
                    ReDim bKomad(511)
                    Get #1, , bKomad
                    Me.ctlComm.Output = bKomad
                Else
                    ReDim bKomad(bKomadSize - 1)
                    Get #1, , bKomad
                    Close #1
                    Me.lblOpis = "Èekam komandu..."
                    Me.ctlProgress.Visible = False
                    Me.ctlProgressBack.Visible = False
                    Me.ctlProgressBorder.Visible = False
                    Me.lblPercent.Visible = False
                    CurrentMode = MODE_LISTEN
                    Me.ctlComm.Output = bKomad
                End If
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_TAKE_INI
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                If FileExists(FileName) Then Kill FileName
                Me.ctlComm.InputMode = comInputModeBinary
                Open FileName For Binary As #1
                CheckReceiveSize
                Me.lblOpis = "Primam: " & FileName & " (" & Round(FileSize / 1024, 1) & "KB)..."
                Me.lblPercent = "0%"
                Me.lblPercent.Visible = True
                Me.ctlProgress.Width = 0
                Me.ctlProgress.Visible = True
                Me.ctlProgressBack.Visible = True
                Me.ctlProgressBorder.Visible = True
                CurrentMode = MODE_TAKE
                Me.ctlComm.Output = "*"
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_TAKE
            ReDim bKomad(Me.ctlComm.RThreshold - 1)
            bKomad = Me.ctlComm.Input
            Put #1, , bKomad
            Me.ctlProgress.Width = LOF(1) / FileSize * Me.ctlProgressBack.Width
            Me.lblPercent = Int(LOF(1) / FileSize * 100) & "%"
            If LOF(1) = FileSize Then
                Close #1
                Me.ctlProgress.Visible = False
                Me.ctlProgressBack.Visible = False
                Me.ctlProgressBorder.Visible = False
                Me.lblPercent.Visible = False
                Me.lblOpis = "Èekam komandu..."
                Me.ctlComm.RThreshold = 1
                Me.ctlComm.InputLen = 0
                Me.ctlComm.InputMode = comInputModeText
                CurrentMode = MODE_LISTEN
                Me.ctlComm.Output = "*"
            Else
                CheckReceiveSize
                Me.ctlComm.Output = "*"
            End If
        Case MODE_VERS
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                CurrentMode = MODE_LISTEN
                Me.ctlComm.Output = VERSION & "*"
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
    End Select
End If
End Sub

Private Function DoCommand(sCommand As String) As Boolean
Dim m_lModelID As Long, m_sModel As String, m_sSlika As String, m_iTip As Integer, m_sMatLica As String, m_sMatDjona As String, m_sBoja As String, m_sSortiment As String, m_sCijena As Single, m_dRok As Date

DoCommand = True

Select Case Left(sCommand, 4)
    Case "USER"
        Me.lblOpis = "Identifikacija u toku..."
        CurrentMode = MODE_USER
    Case "DELE"
        FileName = MyFolder & Mid(sCommand, 6)
        If FileExists(FileName) Then
            Me.lblOpis = "Brišem fajl: " & FileName & "..."
            Kill FileName
            Me.lblOpis = "Èekam komandu..."
        End If
    Case "QUIT"
        Me.lblOpis = "Prekidam vezu..."
        QuitMode = True
    Case "WIPE"
        If FileExists(MyFolder & "modeli.dat") Then
            Me.lblOpis = "Brišem kolekciju..."
            Open MyFolder & "modeli.dat" For Input As #1
                Do Until EOF(1)
                    Input #1, m_lModelID, m_sModel, m_sSlika, m_iTip, m_sMatLica, m_sMatDjona, m_sBoja, m_sSortiment, m_sCijena, m_dRok
                    If FileExists(MyFolder & "pic\" & m_sSlika) Then Kill MyFolder & "pic\" & m_sSlika
                Loop
            Close #1
            Kill MyFolder & "modeli.dat"
            Me.lblOpis = "Èekam komandu..."
        End If
    Case "EXEC"
        If FileExists(MyFolder & "Porucivanje obuce.exe") Then
            Me.lblOpis = "Pokreæem ""Poruèivanje obuæe""..."
            Shell MyFolder & "Porucivanje obuce.exe", vbNormalFocus
            Me.lblOpis = "Èekam komandu..."
        ElseIf FileExists(MyFolder & "Poruèivanje obuæe.exe") Then
            Me.lblOpis = "Pokreæem ""Poruèivanje obuæe""..."
            Shell MyFolder & "Poruèivanje obuæe.exe", vbNormalFocus
            Me.lblOpis = "Èekam komandu..."
        End If
    Case "GIVE"
        FileName = MyFolder & Mid(sCommand, 6)
        If FileExists(FileName) Then
            FileSize = FileLen(FileName)
            If FileSize > 0 Then
                CurrentMode = MODE_GIVE_INI
            Else
                CurrentMode = MODE_NOFILE
            End If
        Else
            CurrentMode = MODE_NOFILE
        End If
    Case "TAKE"
        FileName = MyFolder & MyMid(sCommand, 6, InStr(sCommand, ":") - 1)
        m_sModel = Mid(sCommand, InStr(sCommand, ":") + 1)
        FileSize = Val(m_sModel)
        If CStr(FileSize) <> m_sModel Or FileSize = 0 Then GoTo greška
        If FileExists(Left(FileName, InStrRev(FileName, "\")) & "*") Then
            CurrentMode = MODE_TAKE_INI
        Else
            If MakeFolder(Left(FileName, InStrRev(FileName, "\") - 1)) Then
                CurrentMode = MODE_TAKE_INI
            Else
                GoTo greška
            End If
        End If
    Case "VERS"
        CurrentMode = MODE_VERS
    Case Else
        GoTo greška
End Select
Exit Function

greška:
DoCommand = False
End Function

Private Function MakeFolder(sFolder As String) As Boolean
On Error GoTo greška

If Right(sFolder, 1) = "\" Then sFolder = Left(sFolder, Len(sFolder) - 1)

MakeFolder = True
If FileExists(Left(sFolder, InStrRev(sFolder, "\")) & "*") Then
    MkDir sFolder
Else
    If MakeFolder(Left(sFolder, InStrRev(sFolder, "\") - 1)) Then
        MkDir sFolder
    Else
        GoTo greška
    End If
End If
Exit Function

greška:
MakeFolder = False
End Function

Private Sub CheckReceiveSize()
Me.ctlComm.InputLen = 512
If FileSize - LOF(1) < 512 Then Me.ctlComm.InputLen = FileSize - LOF(1)
Me.ctlComm.RThreshold = Me.ctlComm.InputLen
End Sub

Private Sub CheckConnection()
Dim k As Date
Do While Me.ctlComm.CDHolding
    DoEvents
Loop
If QuitMode Then End
If TransferError Then
    Me.lblTitle = "Došlo je do greške u komunikaciji!"
Else
    Me.lblTitle = "Došlo je do prekida veze!"
End If
Me.lblOpis = ""
If Me.ctlComm.PortOpen Then Me.ctlComm.PortOpen = False
k = Now + #12:00:03 AM#
Do Until k < Now
    DoEvents
Loop
Close #1
If CurrentMode = MODE_TAKE Or CurrentMode = MODE_TAKE_INI Then
    If FileExists(FileName) Then Kill FileName
End If
Me.ctlComm.RThreshold = 1
Me.ctlComm.InputLen = 0
Me.ctlComm.InputMode = comInputModeText
FileName = ""
FileSize = 0
Buffer = ""
TransferError = False
Me.lblOpis = ""
Me.lblPercent.Visible = False
Me.ctlProgress.Visible = False
Me.ctlProgressBack.Visible = False
Me.ctlProgressBorder.Visible = False
Me.ctlComm.PortOpen = True
CurrentMode = MODE_WAITFORCALL
Me.lblTitle = "Èekam poziv..."
End Sub

Private Function FileExists(fn As String) As Boolean
    Dim retval As Long, FindData As WIN32_FIND_DATA, retval2 As Long
    retval = FindFirstFile(fn, FindData)
    If retval <> -1 Then FileExists = True
    retval2 = FindClose(retval)
End Function



Private Function MyMid(sString As String, iStart As Long, iStop As Long) As String
If iStart <= iStop Then
    MyMid = Mid(sString, iStart, iStop - iStart + 1)
End If
End Function


' Tok komunikacije - opis komandi:

' ---> USER*
' <--- *
' ---> *
' <--- User(varijabla)

' ---> VERS*
' <--- *
' ---> *
' <--- 1*

' ---> EXEC*
' <--- *

' ---> DELE filename.ext*
' <--- *

' ---> WIPE*
' <--- *

' ---> QUIT*
' <--- *

' ---> GIVE filename.ext*
' <--- *
' ---> *
' <--- FileSize(varijabla)OK*
' ---> *
' <--- blok1(512 bajta)
' ---> *
' <--- blok2(512 bajta)
' ---> *
' ...
' <--- blokN(<=512 bajta)

' ---> TAKE filename.ext:size*
' <--- *
' ---> *
' <--- *
' ---> blok1(512 bajta)
' <--- *
' ---> blok2(512 bajta)
' <--- *
' ...
' ---> blokN(<=512 bajta)
' <--- *




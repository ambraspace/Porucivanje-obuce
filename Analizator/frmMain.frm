VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Poruèivanje obuæe - analizator"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog ctlComDlg 
      Left            =   2640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm ctlComm 
      Left            =   1200
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
   Begin VB.CommandButton cmdTake 
      Caption         =   "TAKE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdGive 
      Caption         =   "GIVE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "QUIT"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdUser 
      Caption         =   "USER"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdVersion 
      Caption         =   "VERS"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdWipe 
      Caption         =   "WIPE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "EXEC"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Shape ctlProgress 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   120
      Top             =   720
      Visible         =   0   'False
      Width           =   5655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub cmdDelete_Click()
Dim sTMP As String
ButtonsEnabled False
sTMP = InputBox("Upišite naziv fajla:", "Brisanje fajla", "")
If sTMP = "" Then
    ButtonsEnabled True
    Exit Sub
Else
    CurrentMode = MODE_DELE
    Me.ctlComm.Output = "DELE " & sTMP & "*"
End If
End Sub

Private Sub cmdExec_Click()
ButtonsEnabled False
CurrentMode = MODE_EXEC
Me.ctlComm.Output = "EXEC*"
End Sub

Private Sub cmdGive_Click()
ButtonsEnabled False
FileName = InputBox("Upišite naziv fajla:", "GIVE", "")
If FileName = "" Then Exit Sub
FileName = MyFolder & FileName
CurrentMode = MODE_GIVE_INI_1
Me.ctlComm.Output = "GIVE " & Mid(FileName, InStrRev(FileName, "\") + 1) & "*"
End Sub

Private Sub cmdQuit_Click()
ButtonsEnabled False
CurrentMode = MODE_QUIT
Me.ctlComm.Output = "QUIT*"
End Sub

Private Sub cmdTake_Click()
On Error GoTo greška

ButtonsEnabled False
With Me.ctlComDlg
    .CancelError = True
    .DialogTitle = "TAKE"
    .FileName = ""
    .Filter = "All files (*.*)|*.*"
    .Flags = cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNPathMustExist
    .InitDir = MyFolder
    .ShowOpen
End With
FileName = Me.ctlComDlg.FileName
If FileExists(FileName) Then
    FileSize = FileLen(FileName)
    If FileSize = 0 Then
        ButtonsEnabled True
        Exit Sub
    Else
        CurrentMode = MODE_TAKE_INI
        Me.ctlComm.Output = "TAKE " & Me.ctlComDlg.FileTitle & ":" & FileSize & "*"
    End If
Else
    ButtonsEnabled True
    Exit Sub
End If

greška:
End Sub

Private Sub cmdUser_Click()
ButtonsEnabled False
CurrentMode = MODE_USER_INI
Me.ctlComm.Output = "USER*"
End Sub

Private Sub cmdVersion_Click()
ButtonsEnabled False
CurrentMode = MODE_VERS_INI
Me.ctlComm.Output = "VERS*"
End Sub

Private Sub cmdWipe_Click()
ButtonsEnabled False
CurrentMode = MODE_WIPE
Me.ctlComm.Output = "WIPE*"
End Sub

Private Sub ctlComm_OnComm()
Dim sTMP As String, bKomadSize As Long, bKomad() As Byte
If Me.ctlComm.CommEvent = comEvReceive Then
    Select Case CurrentMode
        Case 0
            sTMP = Me.ctlComm.Input
        Case MODE_INIT
            sTMP = Me.ctlComm.Input
            If sTMP = "|" Then
                CurrentMode = 0
                ButtonsEnabled True
                CheckConnection
            Else
                If InitCount = 3 Then
                    Me.ctlComm.PortOpen = False
                    MsgBox "Greška u komunikaciji!", 16, "Greška"
                    End
                Else
                    InitCount = InitCount + 1
                    Me.ctlComm.Output = "*"
                End If
            End If
        Case MODE_USER_INI
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                CurrentMode = MODE_USER
                Me.ctlComm.Output = "*"
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_USER
            Buffer = Buffer & Me.ctlComm.Input
            If Right(Buffer, 1) = "*" Then
                sTMP = Left(Buffer, Len(Buffer) - 1)
                Buffer = ""
                CurrentMode = 0
                MsgBox "Povezani ste sa " & sTMP, vbOKOnly, "USER"
                ButtonsEnabled True
            End If
        Case MODE_EXEC
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                CurrentMode = 0
                MsgBox "Komanda EXEC poslata!", vbInformation + vbOKOnly, "EXEC"
                ButtonsEnabled True
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_WIPE
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                CurrentMode = 0
                MsgBox "Komanda WIPE poslata!", vbInformation + vbOKOnly, "WIPE"
                ButtonsEnabled True
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_DELE
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                CurrentMode = 0
                MsgBox "Komanda DELE poslata!", vbInformation + vbOKOnly, "DELE"
                ButtonsEnabled True
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_VERS_INI
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                CurrentMode = MODE_VERS
                Me.ctlComm.Output = "*"
            ElseIf sTMP = "|" Then
                CurrentMode = 0
                MsgBox "Verzija nepoznata!", vbInformation + vbOKOnly, "VERS"
                ButtonsEnabled True
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_VERS
            Buffer = Buffer & Me.ctlComm.Input
            If Right(Buffer, 1) = "*" Then
                sTMP = Left(Buffer, Len(Buffer) - 1)
                Buffer = ""
                CurrentMode = 0
                MsgBox "Verzija ""ConnAgent""-a je " & sTMP, vbOKOnly, "VERS"
                ButtonsEnabled True
            End If
        Case MODE_QUIT
            sTMP = Me.ctlComm.Input
            CurrentMode = 0
            sTMP = Left(sTMP, 1)
            If sTMP = "*" Then
                QuitMode = True
            Else
                TransferError = True
            End If
            Me.ctlComm.PortOpen = False
        Case MODE_TAKE_INI
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                Open FileName For Binary As #1
                Me.ctlProgress.Width = 0
                Me.ctlProgress.Visible = True
                CurrentMode = MODE_TAKE
                Me.ctlComm.Output = "*"
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_TAKE
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                Me.ctlProgress.Width = (Seek(1) - 1) / FileSize * 5655
                bKomadSize = FileSize - Seek(1) + 1
                If bKomadSize > 512 Then
                    ReDim bKomad(511)
                    Get #1, , bKomad
                    Me.ctlComm.Output = bKomad
                Else
                    ReDim bKomad(bKomadSize - 1)
                    Get #1, , bKomad
                    Close #1
                    CurrentMode = MODE_TAKE_OFF
                    Me.ctlComm.Output = bKomad
                End If
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_TAKE_OFF
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                CurrentMode = 0
                Me.ctlProgress.Width = 5655
                ButtonsEnabled True
                Me.ctlProgress.Visible = False
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_GIVE_INI_1
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                CurrentMode = MODE_GIVE_INI_2
                Me.ctlComm.Output = "*"
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_GIVE_INI_2
            Buffer = Buffer & Me.ctlComm.Input
            If Right(Buffer, 1) = "*" Then
                sTMP = Left(Buffer, Len(Buffer) - 1)
                Buffer = ""
                If sTMP = "ERR" Then
                    CurrentMode = 0
                    MsgBox "Fajl " & Mid(FileName, InStrRev(FileName, "\") + 1) & " ne postoji!", vbExclamation + vbOKOnly, "GIVE"
                    ButtonsEnabled True
                ElseIf Right(sTMP, 2) = "OK" Then
                    sTMP = Left(sTMP, Len(sTMP) - 2)
                    FileSize = Val(sTMP)
                    If FileSize <= 0 Or CStr(FileSize) <> sTMP Then GoTo g1
                    If FileExists(FileName) Then Kill FileName
                    Open FileName For Binary As #1
                    Me.ctlProgress.Width = 0
                    Me.ctlProgress.Visible = True
                    Me.ctlComm.InputMode = comInputModeBinary
                    CheckReceiveSize
                    CurrentMode = MODE_GIVE
                    Me.ctlComm.Output = "*"
                Else
g1:
                    TransferError = True
                    Me.ctlComm.PortOpen = False
                End If
            End If
        Case MODE_GIVE
            ReDim bKomad(Me.ctlComm.RThreshold - 1)
            bKomad = Me.ctlComm.Input
            Put #1, , bKomad
            Me.ctlProgress.Width = LOF(1) / FileSize * 5655
            If LOF(1) = FileSize Then
                Close #1
                CurrentMode = 0
                Me.ctlComm.InputMode = comInputModeText
                Me.ctlComm.RThreshold = 1
                Me.ctlComm.InputLen = 0
                Me.ctlProgress.Visible = False
                ButtonsEnabled True
            Else
                CheckReceiveSize
                Me.ctlComm.Output = "*"
            End If
    End Select
End If
End Sub

Private Sub CheckReceiveSize()
Me.ctlComm.InputLen = 512
If FileSize - LOF(1) < 512 Then Me.ctlComm.InputLen = FileSize - LOF(1)
Me.ctlComm.RThreshold = Me.ctlComm.InputLen
End Sub

Private Sub Form_Load()
Dim a As Integer

If App.PrevInstance Then End

MyFolder = CurDir
'MyFolder = "C:\Poruèivanje obuæe"
If Right(MyFolder, 1) <> "\" Then MyFolder = MyFolder & "\"

PortNo = GetPrivateProfileInt("Main", "PortNo", 0, MyFolder & "settings.ini")
If PortNo = 0 Then
    MsgBox "Ne mogu proèitati osnovne postavke!", 16, "Greška"
    End
End If
Me.ctlComm.CommPort = PortNo

Me.Show
a = MsgBox("Za povezivanje pitisnite dugme ""OK"".", vbExclamation + vbOKCancel, "Povezivanje")
Select Case a
    Case vbCancel
        End
    Case vbOK
        DoPovezivanje
End Select
End Sub

Private Sub DoPovezivanje()
Dim k As Date
Me.ctlComm.PortOpen = True
Me.ctlComm.Output = "AT&FV1E0X3D" & vbCr
k = Now + #12:00:50 AM#
Do Until k < Now
    DoEvents
    If Me.ctlComm.CDHolding Then Exit Do
Loop
If Me.ctlComm.CDHolding Then
    Me.ctlComm.InBufferCount = 0
    Me.ctlComm.OutBufferCount = 0
    CurrentMode = MODE_INIT
    Me.ctlComm.Output = "*"
Else
    Me.ctlComm.PortOpen = False
    MsgBox "Povezivanje nije uspjelo!", vbExclamation + vbOKOnly, "Povezivanje"
    End
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub CheckConnection()
Do
    DoEvents
Loop While Me.ctlComm.CDHolding

If QuitMode Then End
If TransferError Then
    MsgBox "Greška u komunikaciji!", 16, "Greška"
Else
    MsgBox "Došlo je do prekida veze!", 16, "Greška"
End If
Close #1
End
End Sub

Private Sub ButtonsEnabled(b As Boolean)
Me.cmdDelete.Enabled = b
Me.cmdExec.Enabled = Me.cmdDelete.Enabled
Me.cmdGive.Enabled = Me.cmdDelete.Enabled
Me.cmdQuit.Enabled = Me.cmdDelete.Enabled
Me.cmdTake.Enabled = Me.cmdDelete.Enabled
Me.cmdUser.Enabled = Me.cmdDelete.Enabled
Me.cmdVersion.Enabled = Me.cmdDelete.Enabled
Me.cmdWipe.Enabled = Me.cmdDelete.Enabled
End Sub

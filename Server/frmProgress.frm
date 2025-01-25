VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ctlProgressBar 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Center
      Caption         =   "Operacija..."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Property Let Value(sInput As Single)
Me.ctlProgressBar.Value = sInput
End Property

Public Sub Display()
Me.Show
End Sub

''Public Sub DajNarudzbu()
''Me.ctlProgressBar.Min = 0
''Me.ctlProgressBar.Max = 100
''Me.ctlProgressBar.Value = 0
''Me.lblText = "Primanje narudžbe..."
''Me.Show 1, frmGlavni
''CommandAccepted = False
''WrongCommand = False
''CurrentMode = WAIT_MODE
''frmGlavni.ctlComm.Output = "GIVE narudzbe.dat*"
''Do Until CommandAccepted
''    DoEvents
''Loop
''If WrongCommand Then
''    MsgBox "Greška u komunikaciji!", 16, "Poruèivanje obuæe"
''    Me.Hide
''Else
''    CurrentMode = MODE_NARUDZBA_INI
''    frmGlavni.ctlComm.Output = "*"
''End If
''End Sub
''
''Public Sub PosaljiKolekciju()
''Dim i As Long, cTMP As New cFileToSend
''
''Me.ctlProgressBar.Min = 0
''Me.ctlProgressBar.Max = 100
''Me.ctlProgressBar.Value = 0
''Me.lblText = "Slanje kolekcije..."
''Me.Show 1, frmGlavni
''
''Open MyFolder & "modeli.dat" For Output As #1
''rsKolekcija.MoveFirst
''Do Until rsKolekcija.EOF
''    rsModeli.FindFirst "ID=" & rsKolekcija("ModeliID")
''    Write #1, rsKolekcija("ModelID"), rsModeli("Model"), rsKolekcija("Photo"), _
''        rsModeli("MatLica"), rsModeli("MatDjona"), rsModeli("Boja"), _
''        rsModeli("Sortiment"), rsModeli("Cijena"), rsModeli("Rok")
''    cTMP.FileName = MyFolder & "pic\" & rsKolekcija("Photo")
''    cTMP.FileSendString = "\pic" & rsKolekcija("Photo")
''    cTMP.FileIsSent = False
''    cFajloviZaSlanje.Add cTMP
''    Set cTMP = Nothing
''    rsKolekcija.MoveNext
''Loop
''Close #1
''cTMP.FileName = MyFolder & "modeli.dat"
''cTMP.FileSendString = "modeli.dat"
''cTMP.FileIsSent = False
''cFajloviZaSlanje.Add cTMP
''Set cTMP = Nothing
''
''CommandAccepted = False
''WrongCommand = False
''CurrentMode = WAIT_MODE
''frmGlavni.ctlComm.Output = "WIPE*"
''Do Until CommandAccepted
''    DoEvents
''Loop
''If WrongCommand Then
''    MsgBox "Greška u komunikaciji!", 16, "Poruèivanje obuæe"
''    Me.Hide
''    Exit Sub
''End If
''
''For i = 1 To cFajloviZaSlanje.Count
''    If FileExists(cFajloviZaSlanje(i).FileName) And FileLen(cFajloviZaSlanje(i).FileName) Then
''        frmGlavni.PošaljiFajl i
''        Do Until cFajloviZaSlanje(i).FileIsSent
''            DoEvents
''        Loop
''    End If
''Next
''
''Kill MyFolder & "modeli.dat"
''frmGlavni.LocirajKlijenta
''rsKlijenti.Edit
''rsKlijenti("KolekcijaDate") = Date
''rsKlijenti.Update
''frmGlavni.cboKlijenti_Click
''Me.Hide
''frmGlavni.ValidateMe
''End Sub
''

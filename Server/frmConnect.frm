VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Povezivanje"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   ControlBox      =   0   'False
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Poveži se"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      ToolTipText     =   "Pritisnite da biste se povezali sa klijentom"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Odustani"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Pritisnite da biste odustali od povezivanja"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmConnect.frx":030A
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblText 
      Caption         =   "Na telefonu birajte 000/000-000. Za povezivanje pritisnite dugme 'Poveži se' i nakon toga spustite telefonsku slušalicu!"
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub PovežiSe(sTel As String)
Me.lblText = "Na telefonu birajte 9, " & sTel & ". Za povezivanje pritisnite dugme 'Poveži se' i nakon toga spustite telefonsku slušalicu!"
Me.Show 1, frmGlavni
End Sub

Private Sub cmdCancel_Click()
If frmGlavni.ctlComm.PortOpen Then
    frmGlavni.ctlComm.PortOpen = False
    Me.cmdConnect.Enabled = True
    Me.cmdConnect.SetFocus
    Me.cmdCancel.Caption = "Odustani"
    Me.cmdCancel.ToolTipText = "Pritisnite da biste odustali od povezivanja"
Else
    Me.cmdConnect.SetFocus
    Me.Hide
End If
End Sub

Private Sub cmdConnect_Click()
Dim k As Date
Me.cmdConnect.Enabled = False
Me.cmdCancel.Caption = "Prekini"
Me.cmdCancel.ToolTipText = "Pritisnite da biste prekinuli proceduru povezivanja"
frmGlavni.ctlComm.PortOpen = True
frmGlavni.ctlComm.Output = "AT&FE0V1X3D" & vbCr
k = Now + #12:00:50 AM#
Do Until k < Now
    DoEvents
    If frmGlavni.ctlComm.PortOpen = False Then Exit Sub
    If frmGlavni.ctlComm.CDHolding Then Exit Do
Loop
If frmGlavni.ctlComm.CDHolding Then
    frmGlavni.ctlComm.InBufferCount = 0
    frmGlavni.ctlComm.OutBufferCount = 0
    CurrentMode = MODE_INIT
    frmGlavni.ctlComm.Output = "*"
Else
    frmGlavni.ctlComm.PortOpen = False
    MsgBox "Povezivanje nije uspjelo!" & vbCrLf & "Probajte ponovo.", 16, "Poruèivanje obuæe"
    Me.cmdConnect.Enabled = True
    Me.cmdConnect.SetFocus
    Me.cmdCancel.Caption = "Odustani"
    Me.cmdCancel.ToolTipText = "Pritisnite da biste odustali od povezivanja"
End If
End Sub


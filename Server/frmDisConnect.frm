VERSION 5.00
Begin VB.Form frmDisConnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prekid veze - opcije"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Odustani"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   "Povratak na prethodni korak (prozor)"
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdHangUp 
      Caption         =   "Prekini vezu"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "Pritisnite da biste prekinuli vezu sa klijentom"
      Top             =   840
      Width           =   1335
   End
   Begin VB.CheckBox chkExec 
      Caption         =   "Pokreni ""Poruèivanje obuæe"" na udaljenom raèunaru!"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Value           =   1  'Checked
      Width           =   4455
   End
End
Attribute VB_Name = "frmDisConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
Me.Hide
End Sub


Private Sub cmdHangUp_Click()
Me.cmdHangUp.Enabled = False
Me.cmdCancel.Enabled = False
Me.chkExec.Enabled = False
If Me.chkExec.Value = 1 Then
    CurrentMode = MODE_EXEC
    frmGlavni.ctlComm.Output = "EXEC*"
Else
    CurrentMode = MODE_QUIT
    frmGlavni.ctlComm.Output = "QUIT*"
End If
End Sub


Private Sub Form_Load()
Me.chkExec.ToolTipText = "Izaberite ovu opciju ako želite da pokrenete " & _
                    "program za poruèivanje obuæe na udaljenom raèunaru"
End Sub

Public Sub Display(f As Form)
Me.Show 1, f
End Sub


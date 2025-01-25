VERSION 5.00
Begin VB.Form frmBrisati 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Poruèivanje obuæe"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "Ne"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Da"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblPitanje 
      Caption         =   "Da li ste sigurni da želite obrisati sve odabrane modele?"
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmBrisati.frx":0000
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmBrisati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bOdgovor As Boolean

Public Property Get Odgovor()
Odgovor = m_bOdgovor
End Property

Public Sub Display(sPitanje As String, fForm As Form)
Me.lblPitanje = sPitanje
Me.Show 1, fForm
End Sub

Private Sub cmdNo_Click()
m_bOdgovor = False
Me.Hide
End Sub

Private Sub cmdYes_Click()
m_bOdgovor = True
Me.Hide
End Sub


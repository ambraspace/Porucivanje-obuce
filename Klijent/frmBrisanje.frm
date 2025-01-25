VERSION 5.00
Begin VB.Form frmBrisanje 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Brisati narudžbe?"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "Ne"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Da"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Da li ste sigurni da želite obrisati sve narudžbe za ovaj model?"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   300
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmBrisanje.frx":0000
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmBrisanje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_bBrisati As Boolean

Public Property Get Brisati() As Boolean
Brisati = m_bBrisati
End Property

Public Sub Display(f As Form)
Me.Show 1, f
End Sub

Private Sub cmdNo_Click()
m_bBrisati = False
Me.Hide
End Sub

Private Sub cmdYes_Click()
m_bBrisati = True
Me.Hide
End Sub

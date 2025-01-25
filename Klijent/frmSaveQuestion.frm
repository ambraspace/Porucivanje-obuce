VERSION 5.00
Begin VB.Form frmSaveQuestion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Poruèivanje obuæe"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Povratak"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      ToolTipText     =   "Pritisnite 'Povratak' da se vratite na prethodni korak (prozor)!"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "Ne"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      ToolTipText     =   "Pritisnite 'Ne' ako želite kasnije poslati narudžbu!"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Da"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Pritisnite 'Da' ako želite sada poslati narudžbu!"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Da li želite da sada pošaljete narudžbu?"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmSaveQuestion.frx":0000
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmSaveQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_iOdgovor As Integer

Public Property Get Odgovor() As Integer
Odgovor = m_iOdgovor
End Property

Private Sub cmdCancel_Click()
m_iOdgovor = 3
Me.Hide
End Sub

Private Sub cmdNo_Click()
m_iOdgovor = 2
Me.Hide
End Sub

Private Sub cmdYes_Click()
m_iOdgovor = 1
Me.Hide
End Sub

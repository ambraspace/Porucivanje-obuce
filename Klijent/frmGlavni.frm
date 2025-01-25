VERSION 5.00
Begin VB.Form frmGlavni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Poruèivanje obuæe"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "frmGlavni.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7815
   Begin VB.CommandButton cmdPrevious 
      Height          =   375
      Left            =   120
      Picture         =   "frmGlavni.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Prethodni model"
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "Kraj rada"
      Height          =   375
      Left            =   6000
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Kraj rada"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdNaruci 
      Caption         =   "Narudžba"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      ToolTipText     =   "Narudžba za ovaj model"
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Frame fraDetails 
      Height          =   2295
      Left            =   3840
      TabIndex        =   0
      Top             =   480
      Width           =   3855
      Begin VB.TextBox txtTip 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Tip modela"
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtMatLica 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Materijal lica"
         Top             =   520
         Width           =   2175
      End
      Begin VB.TextBox txtRok 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Rok isporuke"
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtCijena 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Cijena modela u prodaji"
         Top             =   1640
         Width           =   2175
      End
      Begin VB.TextBox txtSortiment 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Sortiment"
         Top             =   1360
         Width           =   2175
      End
      Begin VB.TextBox txtBoja 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Boja"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtMatDjona 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Materijal ðona"
         Top             =   800
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Tip:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Materijal lica:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   520
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Rok isporuke:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Cijena:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1640
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Sortiment:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Boja:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Materijal ðona:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   800
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdNext 
      Height          =   375
      Left            =   3000
      Picture         =   "frmGlavni.frx":0C54
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Sljedeæi model"
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblHelp 
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   20
      ToolTipText     =   "Uputstvo"
      Top             =   3480
      Width           =   5775
   End
   Begin VB.Label lblCounter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   960
      TabIndex        =   19
      ToolTipText     =   "Pozicija modela / Ukupno modela"
      Top             =   2902
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   7680
      X2              =   120
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label lblModel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3840
      TabIndex        =   12
      ToolTipText     =   "Naziv modela"
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblNema 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      Height          =   2655
      Left            =   120
      Top             =   120
      Width           =   3615
   End
   Begin VB.Image ctlImageHolder 
      Height          =   2655
      Left            =   120
      Stretch         =   -1  'True
      ToolTipText     =   "Fotografija modela"
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmGlavni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Display()
If cModeli.Count > 0 Then
    Me.Tag = 1
    LoadModel
    ValidateArrows
Else
    Me.cmdNaruci.Enabled = False
    Me.cmdPrevious.Enabled = False
    Me.cmdNext.Enabled = False
    Me.lblNema = "Nema modela!"
End If
Me.Show
End Sub

Public Sub LoadModel()
Dim iModel As Long, cImgSize As ImageSize
iModel = Val(Me.Tag)
Me.lblCounter = iModel & " / " & cModeli.Count
Me.lblModel = cModeli(iModel).Model
Select Case cModeli(iModel).Tip
    Case 7
        Me.txtTip = "muška"
    Case 5
        Me.txtTip = "ženska"
    Case 4
        Me.txtTip = "mladalaèka"
    Case 2
        Me.txtTip = "djeèija"
    Case 0
        Me.txtTip = "bebi"
    Case Else
        Me.txtTip = "?"
End Select
Me.txtMatLica = cModeli(iModel).MatLica
Me.txtMatDjona = cModeli(iModel).MatDjona
Me.txtBoja = cModeli(iModel).Boja
Me.txtSortiment = cModeli(iModel).Sortiment
Me.txtCijena = Format(cModeli(iModel).Cijena, "0.00") & " KM"
Me.txtRok = Format(cModeli(iModel).Rok, "d. M. yyyy.")
PrikažiNarudžbe
Me.ctlImageHolder.Visible = False
Me.ctlImageHolder.Move 120, 120, 3615, 2655
If FileExists(MyFolder & "pic\" & cModeli(iModel).Slika) Then
    cImgSize = GetImageSize(MyFolder & "pic\" & cModeli(iModel).Slika)
    If (cImgSize.Width / cImgSize.Height) > (3615 / 2655) Then
        Me.ctlImageHolder.Height = cImgSize.Height / cImgSize.Width * 3615
        Me.ctlImageHolder.Top = (2655 - Me.ctlImageHolder.Height) / 2 + 120
    ElseIf (cImgSize.Height / cImgSize.Width) > (2655 / 3615) Then
        Me.ctlImageHolder.Width = cImgSize.Width / cImgSize.Height * 2655
        Me.ctlImageHolder.Left = (3615 - Me.ctlImageHolder.Width) / 2 + 120
    End If
    Me.lblNema = ""
    Me.ctlImageHolder.Picture = LoadPicture(MyFolder & "pic\" & cModeli(iModel).Slika)
Else
    Me.lblNema = "Nema fotografije!"
    Me.ctlImageHolder.Picture = LoadPicture("")
End If
Me.ctlImageHolder.Visible = True
ValidateArrows
End Sub

Private Sub PrikažiNarudžbe()
Dim bNarudzbaPostoji As Boolean, i As Integer

If cNarudzbe.Count = 0 Then
    Me.lblHelp.ForeColor = RGB(255, 0, 0)
    Me.lblHelp.Caption = "Ovaj model niste naruèili." & vbCrLf & "Možete ga naruèiti pritiskom na dugme ""Narudžba""."
    Exit Sub
End If

For i = 1 To cNarudzbe.Count
    If cNarudzbe(i).ModelID = cModeli(Val(Me.Tag)).ModelID Then
        bNarudzbaPostoji = True
        Exit For
    End If
Next
    
If bNarudzbaPostoji Then
    Me.lblHelp.ForeColor = RGB(0, 0, 0)
    Me.lblHelp.Caption = "Ovaj model ste naruèili." & vbCrLf & "Da biste poništili ili promijenili narudžbu pritisnite dugme ""Narudžba""."
Else
    Me.lblHelp.ForeColor = RGB(255, 0, 0)
    Me.lblHelp.Caption = "Ovaj model niste naruèili." & vbCrLf & "Možete ga naruèiti pritiskom na dugme ""Narudžba""."
End If
End Sub

Private Sub ValidateArrows()
Me.cmdPrevious.Enabled = Val(Me.Tag) > 1
Me.cmdNext.Enabled = Val(Me.Tag) < cModeli.Count
End Sub

Private Sub cmdEnd_Click()
Unload Me
End Sub


Private Sub cmdNaruci_Click()
frmNaruci.Display Me.lblHelp.ForeColor = RGB(0, 0, 0)
End Sub

Private Sub cmdNext_Click()
Me.Tag = Val(Me.Tag) + 1
LoadModel
Me.cmdNaruci.SetFocus
End Sub

Private Sub cmdPrevious_Click()
Me.Tag = Val(Me.Tag) - 1
LoadModel
Me.cmdNaruci.SetFocus
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim cTMP As cNarudzba
If FileExists(MyFolder & "narudzbe.dat") Then Kill MyFolder & "narudzbe.dat"
If cNarudzbe.Count > 0 Then
    Open MyFolder & "narudzbe.dat" For Output As #1
        For Each cTMP In cNarudzbe
            Write #1, cTMP.ModelID, cTMP.NarudzbaBroj, cTMP.NarudzbaComment
        Next
    Close #1
End If
End Sub


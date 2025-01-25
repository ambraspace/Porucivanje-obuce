VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Štampanje narudžbi"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraSortiranje 
      Caption         =   "Sortiranje"
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      Begin VB.OptionButton optSortProdavnice 
         Caption         =   "po prodavnici"
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optSortModeli 
         Caption         =   "po modelu"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Odustani"
      Height          =   495
      Left            =   4080
      TabIndex        =   20
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Štampaj"
      Default         =   -1  'True
      Height          =   495
      Left            =   5520
      TabIndex        =   19
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtKopija 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6480
      MaxLength       =   2
      TabIndex        =   16
      Text            =   "1"
      Top             =   2130
      Width           =   375
   End
   Begin VB.CheckBox chkComments 
      Caption         =   "Štampaj komentare"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.Frame fraKlijenti 
      Caption         =   "Klijenti"
      Height          =   1215
      Left            =   5400
      TabIndex        =   13
      Top             =   840
      Width           =   1455
      Begin VB.OptionButton optKlijentiTrenutni 
         Caption         =   "trenutni"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optKlijentiSvi 
         Caption         =   "svi"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame fraNarudzbeRange 
      Caption         =   "Štampaj narudžbe"
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2535
      Begin VB.TextBox txtNarudzbeDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton optNarudzbeDate 
         Caption         =   "od datuma"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optNarudzbeLast 
         Caption         =   "posljednje"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optNarudzbeAll 
         Caption         =   "sve"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame fraNaslov 
      Caption         =   "Naslov"
      Height          =   1575
      Left            =   2760
      TabIndex        =   9
      Top             =   840
      Width           =   2535
      Begin VB.OptionButton optNaslovSveStrane 
         Caption         =   "na svim stranicama"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton optNaslovPrvaStrana 
         Caption         =   "na prvoj stranici"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   720
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CheckBox chkNaslovStampaj 
         Caption         =   "Štampaj naslov izvještaja"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Value           =   1  'Checked
         Width           =   2175
      End
   End
   Begin VB.CheckBox chkPageNums 
      Caption         =   "Štampaj brojeve stranica"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.ComboBox cboPrinterSelect 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Kopija:"
      Height          =   255
      Left            =   5760
      TabIndex        =   22
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "A4, uspravno"
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
      TabIndex        =   21
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Štampaè As Printer

Private Sub cboPrinterSelect_Click()
Set Štampaè = Printers(Me.cboPrinterSelect.ListIndex)
End Sub

Private Sub chkNaslovStampaj_Click()
Select Case Me.chkNaslovStampaj.Value
    Case 0
        Me.optNaslovPrvaStrana.Enabled = False
    Case 1
        Me.optNaslovPrvaStrana.Enabled = True
End Select
Me.optNaslovSveStrane.Enabled = Me.optNaslovPrvaStrana.Enabled
End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Public Sub Display()
Me.Show 1, frmGlavni
End Sub

Private Sub cmdPrint_Click()
'Me.Enabled = False
'Set Printer = Štampaè
'Printer.FontName = "Arial CE"
'Printer.Orientation = 1
'Printer.PaperSize = vbPRPSA4
'Printer.ScaleMode = vbMillimeters
'Printer.CurrentY = 10
'If Me.chkNaslovStampaj.Value = 1 Then
    
MsgBox "U pripremi!"

End Sub

Private Sub Form_Load()
Dim iDefPrint As Integer, i As Integer
For i = 0 To Printers.Count - 1
    If Printers(i).DeviceName = Printer.DeviceName Then iDefPrint = i
    Me.cboPrinterSelect.AddItem Printers(i).DeviceName
Next
Me.cboPrinterSelect.ListIndex = iDefPrint
Me.txtNarudzbeDate = Format(Date, "dd-MM-yy")
End Sub

Private Sub optKlijentiSvi_Click()
optKlijentiTrenutni_Click
End Sub

Private Sub optKlijentiTrenutni_Click()
Select Case Me.optKlijentiTrenutni.Value
    Case True
        Me.optSortModeli.Enabled = False
        Me.optSortProdavnice.Value = True
    Case Else
        Me.optSortModeli.Enabled = True
End Select
End Sub

Private Sub optNarudzbeAll_Click()
optNarudzbeDate_Click
End Sub

Private Sub optNarudzbeDate_Click()
Select Case Me.optNarudzbeDate.Value
    Case True
        Me.txtNarudzbeDate.Enabled = True
    Case Else
        Me.txtNarudzbeDate.Enabled = False
End Select
End Sub

Private Sub optNarudzbeLast_Click()
optNarudzbeDate_Click
End Sub

Private Sub txtKopija_GotFocus()
Me.txtKopija.SelStart = 0
Me.txtKopija.SelLength = Len(Me.txtKopija)
End Sub

Private Sub txtNarudzbeDate_Change()
Me.txtNarudzbeDate.SelStart = 0
Me.txtNarudzbeDate.SelLength = Len(Me.txtNarudzbeDate)
End Sub

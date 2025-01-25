VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmModel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Izmijeni model"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   ControlBox      =   0   'False
   Icon            =   "frmModel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog ctlComDlg 
      Left            =   120
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Izmijeni model"
      Default         =   -1  'True
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      ToolTipText     =   "Pritisnite da biste snimili izmjene"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Odustani"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      ToolTipText     =   "Pritisnite da biste se vratili na prethodni korak (prozor)"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdPicture 
      Caption         =   "Izmijeni fotografiju"
      Height          =   375
      Left            =   960
      TabIndex        =   8
      ToolTipText     =   "Pritisnite da biste izmijenili fotografiju"
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Frame fraDetails 
      Height          =   2775
      Left            =   3840
      TabIndex        =   11
      Top             =   0
      Width           =   3855
      Begin VB.TextBox txtModel 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   0
         ToolTipText     =   "Naziv modela"
         Top             =   210
         Width           =   2175
      End
      Begin VB.ComboBox cboTip 
         Height          =   315
         ItemData        =   "frmModel.frx":08CA
         Left            =   1560
         List            =   "frmModel.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   555
         Width           =   2175
      End
      Begin VB.TextBox txtCijena 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   6
         ToolTipText     =   "Cijena modela u prodaji"
         Top             =   2070
         Width           =   615
      End
      Begin VB.TextBox txtRok 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   7
         ToolTipText     =   "Rok isporuke"
         Top             =   2370
         Width           =   735
      End
      Begin VB.TextBox txtSortiment 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   29
         TabIndex        =   5
         ToolTipText     =   "Sortiment"
         Top             =   1770
         Width           =   2175
      End
      Begin VB.TextBox txtBoja 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Boja"
         Top             =   1470
         Width           =   2175
      End
      Begin VB.TextBox txtMatLica 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "Materijal lica"
         Top             =   870
         Width           =   2175
      End
      Begin VB.TextBox txtMatDjona 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   3
         ToolTipText     =   "Materijal ðona"
         Top             =   1170
         Width           =   2175
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Naziv modela:"
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
         TabIndex        =   20
         Top             =   240
         Width           =   1335
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
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1215
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
         TabIndex        =   17
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Cijena (KM):"
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
         TabIndex        =   16
         Top             =   2100
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
         TabIndex        =   15
         Top             =   1800
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
         TabIndex        =   14
         Top             =   1500
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
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
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
         TabIndex        =   12
         Top             =   900
         Width           =   1335
      End
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
      TabIndex        =   18
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
Attribute VB_Name = "frmModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AddNewModel()
Me.Caption = "Dodaj novi model"
Me.cmdAdd.Caption = "Dodaj model"
Me.cmdPicture.Caption = "Dodaj fotografiju"
'Me.txtModel = ""
'Me.cboTip.ListIndex = 0
'Me.txtMatLica = ""
'Me.txtMatDjona = ""
'Me.txtBoja = ""
'Me.txtSortiment = ""
'Me.txtCijena = ""
'Me.txtRok = Format(PosljednjiDatum, "dd-MM-yy")
Me.ctlImageHolder.Picture = LoadPicture("")
Me.ctlComDlg.FileName = ""
Me.Show 1, frmGlavni
End Sub

Public Sub EditModel()
Me.Caption = "Izmijeni model"
Me.cmdAdd.Caption = "Izmijeni model"
Me.cmdPicture.Caption = "Izmijeni fotografiju"
rsKolekcija.FindFirst "ID=" & Mid(frmGlavni.ctlListKolekcija.SelectedItem.Key, 3)
rsModeli.FindFirst "ID=" & rsKolekcija("ModelID")
Me.ctlComDlg.FileName = MyFolder & "pic\" & rsKolekcija("Photo")
If FileExists(Me.ctlComDlg.FileName) Then
    UèitajSliku Me.ctlComDlg.FileName
Else
    Me.ctlImageHolder.Picture = LoadPicture("")
End If
Me.txtModel = rsModeli("Model")
Select Case rsModeli("Tip")
    Case 7
        Me.cboTip.ListIndex = 0
    Case 5
        Me.cboTip.ListIndex = 1
    Case 4
        Me.cboTip.ListIndex = 2
    Case 2
        Me.cboTip.ListIndex = 3
    Case 0
        Me.cboTip.ListIndex = 4
End Select
Me.txtMatLica = rsModeli("MatLica")
Me.txtMatDjona = rsModeli("MatDjona")
Me.txtBoja = rsModeli("Boja")
Me.txtSortiment = rsModeli("Sortiment")
Me.txtCijena = Format(rsModeli("Cijena"), "0.00")
Me.txtRok = Format(rsModeli("Rok"), "dd-MM-yy")
Me.Show 1, frmGlavni
End Sub

Private Sub cboTip_KeyPress(KeyAscii As Integer)
Select Case Chr(KeyAscii)
    Case "0", "1"
        Me.cboTip.ListIndex = 4
    Case "2", "3"
        Me.cboTip.ListIndex = 3
    Case "4"
        Me.cboTip.ListIndex = 2
    Case "5"
        Me.cboTip.ListIndex = 1
    Case "7"
        Me.cboTip.ListIndex = 0
End Select
End Sub

Private Sub cmdAdd_Click()

If Trim(Me.txtModel) = "" Then
    MsgBox "Polje ""Model"" nije pravilno zadato!", vbCritical + vbOKOnly, "Poruèivanje obuæe"
    Exit Sub
End If
If Trim(Me.txtMatLica) = "" Then
    MsgBox "Polje ""Materijal lica"" nije pravilno zadato!", vbCritical + vbOKOnly, "Poruèivanje obuæe"
    Exit Sub
End If
If Trim(Me.txtMatDjona) = "" Then
    MsgBox "Polje ""Materijal ðona"" nije pravilno zadato!", vbCritical + vbOKOnly, "Poruèivanje obuæe"
    Exit Sub
End If
If Not BojaUredu Then
    MsgBox "Polje ""Boja"" nije pravilno zadato!", vbCritical + vbOKOnly, "Poruèivanje obuæe"
    Exit Sub
End If
If Not SortimentUredu Then
    MsgBox "Polje ""Sortiment"" nije pravilno zadato!", vbCritical + vbOKOnly, "Poruèivanje obuæe"
    Exit Sub
End If
If Not CijenaUredu Then
    MsgBox "Polje ""Cijena"" nije pravilno zadato!", vbCritical + vbOKOnly, "Poruèivanje obuæe"
    Exit Sub
End If
If Not RokUredu Then
    MsgBox "Polje ""Rok isporuke"" nije pravilno zadato!", vbCritical + vbOKOnly, "Poruèivanje obuæe"
    Exit Sub
End If
If Not FileExists(Me.ctlComDlg.FileName) Then
    MsgBox "Niste odabrali fotografiju modela!", vbCritical + vbOKOnly, "Poruèivanje obuæe"
    Exit Sub
End If

Select Case Me.cmdAdd.Caption
    Case "Dodaj model"
        rsModeli.AddNew
        rsModeli("Model") = Me.txtModel
        Select Case Me.cboTip.ListIndex
            Case 0
                rsModeli("Tip") = 7
            Case 1
                rsModeli("Tip") = 5
            Case 2
                rsModeli("Tip") = 4
            Case 3
                rsModeli("Tip") = 2
            Case 4
                rsModeli("Tip") = 0
        End Select
        rsModeli("MatLica") = Me.txtMatLica
        rsModeli("MatDjona") = Me.txtMatDjona
        rsModeli("Boja") = Me.txtBoja
        rsModeli("Sortiment") = Me.txtSortiment
        rsModeli("Cijena") = CSng(Me.txtCijena)
        rsModeli("Rok") = CDate(Me.txtRok)
        PosljednjiDatum = CDate(Me.txtRok)
        rsKolekcija.AddNew
        rsKolekcija("ModelID") = rsModeli("ID")
        rsKolekcija("Photo") = "p" & rsKolekcija("ID") & ".jpg"
        FileCopy Me.ctlComDlg.FileName, MyFolder & "pic\" & rsKolekcija("Photo")
        frmGlavni.ctlListKolekcija.ListItems.Add , "ID" & rsKolekcija("ID"), rsModeli("Model")
        frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(1) = frmGlavni.ctlListKolekcija.ListItems.Count
        frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(2) = Me.cboTip.Text
        frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(3) = rsModeli("MatLica")
        frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(4) = rsModeli("MatDjona")
        frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(5) = rsModeli("Boja")
        frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(6) = rsModeli("Sortiment")
        frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(7) = Format(rsModeli("Cijena"), "0.00")
        frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(8) = Format(rsModeli("Rok"), "d. M. yyyy.")
        rsModeli.Update
        rsKolekcija.Update
        frmGlavni.cmdDelModel.Enabled = frmGlavni.ModelSelected
        frmGlavni.cmdEditModel.Enabled = frmGlavni.cmdDelModel.Enabled
        Me.Hide
    Case "Izmijeni model"
        rsModeli.Edit
        rsModeli("Model") = Me.txtModel
        Select Case Me.cboTip.ListIndex
            Case 0
                rsModeli("Tip") = 7
            Case 1
                rsModeli("Tip") = 5
            Case 2
                rsModeli("Tip") = 4
            Case 3
                rsModeli("Tip") = 2
            Case 4
                rsModeli("Tip") = 0
        End Select
        rsModeli("MatLica") = Me.txtMatLica
        rsModeli("MatDjona") = Me.txtMatDjona
        rsModeli("Boja") = Me.txtBoja
        rsModeli("Sortiment") = Me.txtSortiment
        rsModeli("Cijena") = CSng(Me.txtCijena)
        rsModeli("Rok") = CDate(Me.txtRok)
        If Me.ctlComDlg.FileName <> MyFolder & "pic\" & rsKolekcija("Photo") Then
            If FileExists(MyFolder & "pic\" & rsKolekcija("Photo")) Then Kill MyFolder & "pic\" & rsKolekcija("Photo")
            FileCopy Me.ctlComDlg.FileName, MyFolder & "pic\" & rsKolekcija("Photo")
        End If
        frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).Text = rsModeli("Model")
        'frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(1) = frmGlavni.ctlListKolekcija.ListItems.Count
        frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(2) = Me.cboTip.Text
        frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(3) = rsModeli("MatLica")
        frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(4) = rsModeli("MatDjona")
        frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(5) = rsModeli("Boja")
        frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(6) = rsModeli("Sortiment")
        frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(7) = Format(rsModeli("Cijena"), "0.00")
        frmGlavni.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(8) = Format(rsModeli("Rok"), "d. M. yyyy.")
        rsModeli.Update
        Me.Hide
End Select
frmGlavni.ctlListKolekcija_ItemCheck frmGlavni.ctlListKolekcija.ListItems(1)
frmGlavni.cmdDelModel.Enabled = frmGlavni.ModelSelected
frmGlavni.cmdEditModel.Enabled = frmGlavni.cmdDelModel.Enabled
End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdPicture_Click()
With Me.ctlComDlg
    .CancelError = False
    .DialogTitle = "Izaberite fotografiju modela"
    .Filter = "JPEG slike (*.jpg, *.jpeg, *.jpe)|*.jpg;*.jpeg;*.jpe)"
    .Flags = cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNPathMustExist
    .ShowOpen
End With
Me.ctlImageHolder.Visible = False
Me.ctlImageHolder.Move 120, 120, 3615, 2655
If FileExists(Me.ctlComDlg.FileName) Then
    UèitajSliku Me.ctlComDlg.FileName
Else
    Me.ctlImageHolder.Picture = LoadPicture("")
End If
Me.ctlImageHolder.Visible = True
End Sub

Private Sub UèitajSliku(sFileName As String)
    Dim t As ImageSize
        t = GetImageSize(sFileName)
        If (t.Width / t.Height) > (3615 / 2655) Then
            Me.ctlImageHolder.Height = t.Height / t.Width * 3615
            Me.ctlImageHolder.Top = (2655 - Me.ctlImageHolder.Height) / 2 + 120
        ElseIf (t.Width / t.Height) < (3615 / 2655) Then
            Me.ctlImageHolder.Width = t.Width / t.Height * 2655
            Me.ctlImageHolder.Left = (3615 - Me.ctlImageHolder.Width) / 2 + 120
        End If
        Me.ctlImageHolder.Picture = LoadPicture(sFileName)
End Sub

Private Sub Form_Load()
Me.cboTip.AddItem "muška"
Me.cboTip.AddItem "ženska"
Me.cboTip.AddItem "mladalaèka"
Me.cboTip.AddItem "djeèija"
Me.cboTip.AddItem "bebi"
Me.cboTip.ListIndex = 0
End Sub


Private Sub SelectAllText(a As Control)
a.SelStart = 0
a.SelLength = Len(a.Text)
End Sub

Private Sub txtModel_GotFocus()
SelectAllText Me.txtModel
End Sub

Private Sub txtRok_GotFocus()
SelectAllText Me.txtRok
End Sub

Private Sub txtSortiment_GotFocus()
SelectAllText Me.txtSortiment
End Sub

Private Sub txtBoja_GotFocus()
SelectAllText Me.txtBoja
End Sub

Private Sub txtCijena_GotFocus()
SelectAllText Me.txtCijena
End Sub

Private Sub txtMatDjona_GotFocus()
SelectAllText Me.txtMatDjona
End Sub

Private Sub txtMatLica_GotFocus()
SelectAllText Me.txtMatLica
End Sub


Private Sub txtBoja_KeyPress(KeyAscii As Integer)
If Not (IsString(Chr(KeyAscii)) Or KeyAscii = 44 Or KeyAscii = 8) Then
    KeyAscii = 0
    Beep
    Exit Sub
End If
If KeyAscii = 44 Then
    If Right(Me.txtBoja, 1) = "," Or Me.txtBoja = "" Then KeyAscii = 0
End If
End Sub

Private Sub txtSortiment_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 45) Then
    KeyAscii = 0
    Beep
    Exit Sub
End If
If (KeyAscii = 44 Or KeyAscii = 45) And Me.txtSortiment = "" Then KeyAscii = 0
If KeyAscii = 45 And ((InStrRev(Me.txtSortiment, "-") > InStrRev(Me.txtSortiment, ",")) Or Right(Me.txtSortiment, 1) = ",") Then KeyAscii = 0
If KeyAscii = 44 And (Right(Me.txtSortiment, 1) = "," Or Right(Me.txtSortiment, 1) = "-") Then KeyAscii = 0

End Sub

Private Sub txtCijena_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 44) Then
    KeyAscii = 0
    Beep
End If
If KeyAscii = 44 And Me.txtCijena = "" Then KeyAscii = 0
If KeyAscii = 44 And InStr(Me.txtCijena, ",") > 0 Then KeyAscii = 0
End Sub

Private Sub txtRok_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8) Then
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub txtRok_Change()
If Len(Me.txtRok) = 2 Or Len(Me.txtRok) = 5 Then Me.txtRok = Me.txtRok & "-"
Me.txtRok.SelStart = Len(Me.txtRok)
End Sub

Private Function BojaUredu() As Boolean
Dim i As Long

o1:
If Left(Me.txtBoja, 1) = "," Then
    Me.txtBoja = Mid(Me.txtBoja, 2)
    GoTo o1
Else
o2:
    If Right(Me.txtBoja, 1) = "," Then
        Me.txtBoja = Left(Me.txtBoja, Len(Me.txtBoja) - 1)
        GoTo o2
    End If
End If

Me.txtBoja = Trim(Me.txtBoja)

BojaUredu = True

Select Case Len(Me.txtBoja)
    Case Is > 1
        For i = 1 To Len(Me.txtBoja) - 1
            If Mid(Me.txtBoja, i, 1) = "," And Mid(Me.txtBoja, i + 1, 1) = "," Then BojaUredu = False
        Next
        If fnNumOfChars(Me.txtBoja, ",") > 3 Then BojaUredu = False
    Case 1
        BojaUredu = True
    Case 0
        BojaUredu = False
End Select

End Function

Private Function fnNumOfChars(sString As String, sChar As String) As Long
Dim i As Long
For i = 1 To Len(sString)
    If Mid(sString, i, 1) = sChar Then fnNumOfChars = fnNumOfChars + 1
Next
End Function

Private Function SortimentUredu() As Boolean
Dim sBrojevi() As Integer, sZnakovi() As String, i As Long, bZnakoviOK As Boolean, bBrojeviOK As Boolean
' ako dužina polja nije 2,5,8 itd. izlazi iz funkcije
If (Len(Me.txtSortiment) + 1) Mod 3 <> 0 Then Exit Function
ReDim sBrojevi(((Len(Me.txtSortiment) + 1) / 3) - 1)
' popunjava brojeve
For i = 0 To UBound(sBrojevi)
    If Mid(Me.txtSortiment, 3 * i + 1, 2) <> CStr(Val(Mid(Me.txtSortiment, 3 * i + 1, 2))) Then
        Exit Function
    Else
        sBrojevi(i) = Val(Mid(Me.txtSortiment, 3 * i + 1, 2))
    End If
Next
'popunjava znakove
If Len(Me.txtSortiment) > 4 Then
    ReDim sZnakovi(UBound(sBrojevi) - 1)
    For i = 0 To UBound(sZnakovi)
        sZnakovi(i) = Mid(Me.txtSortiment, (i + 1) * 3, 1)
        If sZnakovi(i) <> "-" And sZnakovi(i) <> "," Then Exit Function
    Next
    
    If UBound(sZnakovi) = 0 Then
        bZnakoviOK = True
        GoTo n1
    End If
    bZnakoviOK = True
    For i = 0 To UBound(sZnakovi) - 1
        If sZnakovi(i) = "-" And sZnakovi(i + 1) = "-" Then
            Exit Function
        End If
    Next
Else
    bZnakoviOK = True
End If

n1:
If UBound(sBrojevi) = 0 Then
    bBrojeviOK = True
Else
    bBrojeviOK = True
    For i = 0 To UBound(sBrojevi) - 1
        If sBrojevi(i) >= sBrojevi(i + 1) Then Exit Function
    Next
End If

If bBrojeviOK And bZnakoviOK Then SortimentUredu = True
End Function

Private Function CijenaUredu() As Boolean
Dim sTMP As String, i As Integer
o1:
If Left(Me.txtCijena, 1) = "," Then
    Me.txtCijena = Mid(Me.txtCijena, 2)
    GoTo o1
Else
o2:
    If Right(Me.txtCijena, 1) = "," Then
        Me.txtCijena = Left(Me.txtCijena, Len(Me.txtCijena) - 1)
        GoTo o2
    End If
End If

Me.txtCijena = Trim(Me.txtCijena)

If Me.txtCijena = "" Then Exit Function
For i = 1 To Len(Me.txtCijena)
    Select Case Mid(Me.txtCijena, i, 1)
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            sTMP = sTMP & "0"
        Case ","
            sTMP = sTMP & "."
        Case Else
            Exit Function
    End Select
Next
If Format(CSng(Me.txtCijena.Text), sTMP) = Me.txtCijena.Text Then CijenaUredu = True
End Function

Private Function RokUredu() As Boolean
If IsDate(Me.txtRok) Then
    If Format(Me.txtRok, "dd-MM-yy") = Me.txtRok Then RokUredu = True
End If
End Function





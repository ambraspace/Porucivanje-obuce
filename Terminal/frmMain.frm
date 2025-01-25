VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PO Terminal"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin MSCommLib.MSComm ctlComm 
      Left            =   3720
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      ParityReplace   =   0
      RThreshold      =   1
      RTSEnable       =   -1  'True
      BaudRate        =   56000
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   4455
   End
   Begin VB.TextBox txtIn 
      BackColor       =   &H80000004&
      Height          =   2775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PortNo As Integer


Private Sub ctlComm_OnComm()
If Me.ctlComm.CommEvent = comEvReceive Then
    Me.txtIn = Me.txtIn & Me.ctlComm.Input
    Osvježi
End If
End Sub

Private Sub Form_Load()
PortNo = Val(Left(InputBox("Upišite broj porta:", "PO Terminal", ""), 1))
If PortNo < 1 Or PortNo > 6 Then
    MsgBox "Greška!", 16, "PO Terminal"
    End
End If
Me.ctlComm.CommPort = PortNo
Me.ctlComm.PortOpen = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.ctlComm.PortOpen = False
End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Trim(Me.txtSend) <> "" Then
    Select Case Me.ctlComm.CDHolding
        Case True
            Me.ctlComm.Output = CStr(Me.txtSend)
        Case False
            Me.ctlComm.Output = CStr(Me.txtSend) & vbCr
    End Select
    Me.txtIn = Me.txtIn & Me.txtSend & vbCrLf
    Me.txtSend = ""
    Osvježi
End If
End Sub

Private Sub Osvježi()
If Len(Me.txtIn) > 3072 Then Me.txtIn = Right(Me.txtIn, 2048)
Me.txtIn.SelStart = Len(Me.txtIn)
End Sub


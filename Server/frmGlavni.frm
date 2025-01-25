VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGlavni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Poruèivanje obuæe"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   Icon            =   "frmGlavni.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm ctlComm 
      Left            =   6240
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
   Begin VB.Frame Frame3 
      Caption         =   "Porudžbine"
      Height          =   2060
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   9135
      Begin VB.TextBox txtKomentar 
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   1650
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   58
         Text            =   "00"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   5160
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   57
         Text            =   "00"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   56
         Text            =   "00"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   5880
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   55
         Text            =   "00"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   54
         Text            =   "00"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   6600
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   53
         Text            =   "00"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   52
         Text            =   "00"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   7320
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   51
         Text            =   "00"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   50
         Text            =   "00"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   8040
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   49
         Text            =   "00"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   48
         Text            =   "00"
         Top             =   750
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   11
         Left            =   5160
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   47
         Text            =   "00"
         Top             =   750
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   12
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   46
         Text            =   "00"
         Top             =   750
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   13
         Left            =   5880
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   45
         Text            =   "00"
         Top             =   750
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   14
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   44
         Text            =   "00"
         Top             =   750
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   15
         Left            =   6600
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   43
         Text            =   "00"
         Top             =   750
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   16
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   42
         Text            =   "00"
         Top             =   750
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   17
         Left            =   7320
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   41
         Text            =   "00"
         Top             =   750
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   18
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   40
         Text            =   "00"
         Top             =   750
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   19
         Left            =   8040
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   39
         Text            =   "00"
         Top             =   750
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   20
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   38
         Text            =   "00"
         Top             =   1020
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   21
         Left            =   5160
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   37
         Text            =   "00"
         Top             =   1020
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   22
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   36
         Text            =   "00"
         Top             =   1020
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   23
         Left            =   5880
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   35
         Text            =   "00"
         Top             =   1020
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   24
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   34
         Text            =   "00"
         Top             =   1020
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   25
         Left            =   6600
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   33
         Text            =   "00"
         Top             =   1020
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   26
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   32
         Text            =   "00"
         Top             =   1020
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   27
         Left            =   7320
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   31
         Text            =   "00"
         Top             =   1020
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   28
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   30
         Text            =   "00"
         Top             =   1020
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   29
         Left            =   8040
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   29
         Text            =   "00"
         Top             =   1020
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   30
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   28
         Text            =   "00"
         Top             =   1290
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   31
         Left            =   5160
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   27
         Text            =   "00"
         Top             =   1290
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   32
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   26
         Text            =   "00"
         Top             =   1290
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   33
         Left            =   5880
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   25
         Text            =   "00"
         Top             =   1290
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   34
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   24
         Text            =   "00"
         Top             =   1290
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   35
         Left            =   6600
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "00"
         Top             =   1290
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   36
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   22
         Text            =   "00"
         Top             =   1290
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   37
         Left            =   7320
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   21
         Text            =   "00"
         Top             =   1290
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   38
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   20
         Text            =   "00"
         Top             =   1290
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   39
         Left            =   8040
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   19
         Text            =   "00"
         Top             =   1290
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.ListView ctlListNarudzbe 
         Height          =   1695
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Model"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Datum"
            Object.Width           =   2381
         EndProperty
      End
      Begin VB.CommandButton cmdPrintNarudzbe 
         Height          =   375
         Left            =   8640
         Picture         =   "frmGlavni.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Štampaj porudžbine"
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cmdDelNarudzba 
         Height          =   375
         Left            =   8640
         Picture         =   "frmGlavni.frx":0C54
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Izbriši porudžbinu"
         Top             =   900
         Width           =   375
      End
      Begin VB.CommandButton cmdDownload 
         Enabled         =   0   'False
         Height          =   375
         Left            =   8640
         Picture         =   "frmGlavni.frx":0FDE
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Uzmi porudžbinu"
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Komentar:"
         Height          =   255
         Left            =   3840
         TabIndex        =   74
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblBroj 
         Alignment       =   2  'Center
         Caption         =   "00"
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
         Index           =   0
         Left            =   4800
         TabIndex        =   73
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblBroj 
         Alignment       =   2  'Center
         Caption         =   "00"
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
         Index           =   1
         Left            =   5160
         TabIndex        =   72
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblBroj 
         Alignment       =   2  'Center
         Caption         =   "00"
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
         Index           =   2
         Left            =   5520
         TabIndex        =   71
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblBroj 
         Alignment       =   2  'Center
         Caption         =   "00"
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
         Index           =   3
         Left            =   5880
         TabIndex        =   70
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblBroj 
         Alignment       =   2  'Center
         Caption         =   "00"
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
         Index           =   4
         Left            =   6240
         TabIndex        =   69
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblBroj 
         Alignment       =   2  'Center
         Caption         =   "00"
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
         Index           =   5
         Left            =   6600
         TabIndex        =   68
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblBroj 
         Alignment       =   2  'Center
         Caption         =   "00"
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
         Index           =   6
         Left            =   6960
         TabIndex        =   67
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblBroj 
         Alignment       =   2  'Center
         Caption         =   "00"
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
         Index           =   7
         Left            =   7320
         TabIndex        =   66
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblBroj 
         Alignment       =   2  'Center
         Caption         =   "00"
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
         Index           =   8
         Left            =   7680
         TabIndex        =   65
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblBroj 
         Alignment       =   2  'Center
         Caption         =   "00"
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
         Index           =   9
         Left            =   8040
         TabIndex        =   64
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Brojevi (fr.):"
         Height          =   255
         Left            =   3840
         TabIndex        =   63
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblBoja 
         Alignment       =   1  'Right Justify
         Caption         =   "boja1:"
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   62
         Top             =   495
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblBoja 
         Alignment       =   1  'Right Justify
         Caption         =   "boja1:"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   61
         Top             =   765
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblBoja 
         Alignment       =   1  'Right Justify
         Caption         =   "boja1:"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   60
         Top             =   1035
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblBoja 
         Alignment       =   1  'Right Justify
         Caption         =   "boja1:"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   59
         Top             =   1305
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Kolekcija za slanje"
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   9135
      Begin VB.CheckBox chkSelect 
         Height          =   255
         Left            =   8640
         TabIndex        =   76
         Top             =   2160
         Value           =   1  'Checked
         Width           =   255
      End
      Begin MSComctlLib.ListView ctlListKolekcija 
         Height          =   2175
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Model"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "#"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tip"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Materijal lica"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Materijal ðona"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Boja"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Sortiment"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Cijena"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Rok isporuke"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.CommandButton cmdEditModel 
         Height          =   375
         Left            =   8640
         Picture         =   "frmGlavni.frx":1368
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Izmijeni model"
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton cmdDelModel 
         Height          =   375
         Left            =   8640
         Picture         =   "frmGlavni.frx":16F2
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Izbriši model"
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton cmdAddModel 
         Height          =   375
         Left            =   8640
         Picture         =   "frmGlavni.frx":1A7C
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Dodaj model"
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdUpload 
         Enabled         =   0   'False
         Height          =   375
         Left            =   8640
         Picture         =   "frmGlavni.frx":1E06
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Pošalji kolekciju modela"
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Podaci o klijentu"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.ComboBox cboKlijenti 
         Height          =   315
         ItemData        =   "frmGlavni.frx":2190
         Left            =   120
         List            =   "frmGlavni.frx":2192
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Ime klijenta - prodavnice"
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Poveži se"
         Default         =   -1  'True
         Height          =   375
         Left            =   7560
         TabIndex        =   5
         ToolTipText     =   "Pritisnite da biste se povezali sa klijentom"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblOrderDate 
         Caption         =   "Porudžbina primljena: 00. 00. 0000."
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         ToolTipText     =   "Datum posljednje porudžbine"
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblKolekcijaDate 
         Caption         =   "Kolekcija poslata: 00. 00. 0000."
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         ToolTipText     =   "Datum posljednjeg slanja kolekcije"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Klijent:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Image imgOff 
         Height          =   240
         Left            =   8760
         Picture         =   "frmGlavni.frx":2194
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgOn 
         Height          =   240
         Left            =   8760
         Picture         =   "frmGlavni.frx":251E
         Top             =   180
         Width           =   240
      End
   End
   Begin MSComctlLib.ProgressBar ctlProgressBar 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   6120
      Visible         =   0   'False
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblAkcija 
      Caption         =   "Operacija..."
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   5880
      Visible         =   0   'False
      Width           =   9135
   End
End
Attribute VB_Name = "frmGlavni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Display()

Me.ctlComm.CommPort = PortNo
rsKlijenti.MoveLast
rsKlijenti.MoveFirst
Do Until rsKlijenti.EOF
    Me.cboKlijenti.AddItem rsKlijenti("Klijent")
    rsKlijenti.MoveNext
Loop
Me.cboKlijenti.ListIndex = 0
If rsKolekcija.RecordCount > 0 Then
    rsKolekcija.MoveLast
    rsKolekcija.MoveFirst
    Do Until rsKolekcija.EOF
        rsModeli.FindFirst "ID=" & rsKolekcija("ModelID")
        Me.ctlListKolekcija.ListItems.Add , "ID" & rsKolekcija("ID"), rsModeli("Model")
        Me.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(1) = Me.ctlListKolekcija.ListItems.Count
        Select Case rsModeli("Tip")
            Case 0
                Me.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(2) = "bebi"
            Case 2
                Me.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(2) = "djeèija"
            Case 4
                Me.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(2) = "mladalaèka"
            Case 5
                Me.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(2) = "ženska"
            Case 7
                Me.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(2) = "muška"
            Case Else
                Me.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(2) = "?"
        End Select
        Me.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(3) = rsModeli("MatLica")
        Me.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(4) = rsModeli("MatDjona")
        Me.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(5) = rsModeli("Boja")
        Me.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(6) = rsModeli("Sortiment")
        Me.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(7) = Format(rsModeli("Cijena"), "0.00")
        Me.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).SubItems(8) = Format(rsModeli("Rok"), "d. M. yyyy.")
        rsKolekcija.MoveNext
    Loop
End If
cboKlijenti_Click
chkSelect_Click
Me.cmdDelModel.Enabled = Me.ModelSelected
Me.cmdEditModel.Enabled = Me.cmdDelModel.Enabled
Me.cmdDelNarudzba.Enabled = Me.ctlListNarudzbe.ListItems.Count > 0
Me.cmdPrintNarudzbe.Enabled = rsNarudzbe.RecordCount > 0
MainFormHeight = Me.Height
Me.Show
End Sub

Public Property Get ModelSelected() As Boolean
Dim c As MSComctlLib.ListItem
If Me.ctlListKolekcija.ListItems.Count = 0 Then Exit Property
For Each c In Me.ctlListKolekcija.ListItems
    If c.Selected Then
        ModelSelected = True
        Exit Property
    End If
Next
End Property


Public Property Get NarudzbaSelected() As Boolean
Dim c As MSComctlLib.ListItem
If Me.ctlListNarudzbe.ListItems.Count = 0 Then Exit Property
For Each c In Me.ctlListNarudzbe.ListItems
    If c.Selected Then
        NarudzbaSelected = True
        Exit Property
    End If
Next
End Property

Private Sub cboKlijenti_Click()

LocirajKlijenta

Me.cmdConnect.ToolTipText = "Pritisnite da biste se povezali sa " & Me.cboKlijenti.Text

If IsDate(rsKlijenti("KolekcijaDate")) Then
    Me.lblKolekcijaDate = "Kolekcija poslata: " & Format(rsKlijenti("KolekcijaDate"), "d. M. yyyy.")
Else
    Me.lblKolekcijaDate = "Kolekcija poslata: nikad"
End If
If IsDate(rsKlijenti("NarudzbaDate")) Then
    Me.lblOrderDate = "Porudžbina primljena: " & Format(rsKlijenti("NarudzbaDate"), "d. M. yyyy.")
Else
    Me.lblOrderDate = "Porudžbina primljena: nikad"
End If

Me.ctlListNarudzbe.ListItems.Clear
If rsNarudzbe.RecordCount = 0 Then GoTo kraj
rsNarudzbe.MoveLast
rsNarudzbe.MoveFirst
rsNarudzbe.FindFirst "KlijentID=" & rsKlijenti("ID")
If rsNarudzbe.NoMatch Then GoTo kraj
GoTo e
Opet:
rsNarudzbe.FindNext "KlijentID=" & rsKlijenti("ID")
If rsNarudzbe.NoMatch Then GoTo kraj
e:
rsModeli.FindFirst "ID=" & rsNarudzbe("ModelID")
Me.ctlListNarudzbe.ListItems.Add , "ID" & rsNarudzbe("ID"), rsModeli("Model")
Me.ctlListNarudzbe.ListItems("ID" & rsNarudzbe("ID")).SubItems(1) = Format(rsNarudzbe("Datum"), "d. M. yyyy.")
GoTo Opet

kraj:
Me.cmdDelNarudzba.Enabled = Me.ctlListNarudzbe.ListItems.Count > 0
ctlListNarudzbe_ItemClick Me.ctlListNarudzbe.SelectedItem
End Sub



Private Sub chkSelect_Click()
Dim cTMP As MSComctlLib.ListItem
If Me.ctlListKolekcija.ListItems.Count = 0 Then
    Me.chkSelect.Value = 0
    GoTo k1
End If
For Each cTMP In Me.ctlListKolekcija.ListItems
    If Me.chkSelect.Value = 0 Then
        cTMP.Checked = False
    ElseIf Me.chkSelect.Value = 1 Then
        cTMP.Checked = True
    End If
Next
k1:
Me.cmdUpload.Enabled = Me.cmdDownload.Enabled And Me.chkSelect.Value > 0
End Sub

Private Sub cmdAddModel_Click()
frmModel.AddNewModel
End Sub


Private Sub cmdConnect_Click()
Select Case Me.cmdConnect.Caption
    Case "Poveži se"
        frmConnect.PovežiSe rsKlijenti("Telefon")
    Case "Prekini vezu"
        frmDisConnect.Display Me
End Select
End Sub

Private Sub cmdDelModel_Click()
Dim i As Long, c As MSComctlLib.ListItem
frmBrisati.Display "Da li ste sigurni da želite izbrisati sve izabrane modele?", Me
If frmBrisati.Odgovor = True Then
    rsKolekcija.MoveLast
    For i = rsKolekcija.RecordCount To 1 Step -1
        Set c = Me.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID"))
        If c.Selected Then
            If FileExists(MyFolder & "pic\" & rsKolekcija("Photo")) Then Kill MyFolder & "pic\" & rsKolekcija("Photo")
            Me.ctlListKolekcija.ListItems.Remove "ID" & rsKolekcija("ID")
            rsKolekcija.Delete
        End If
        rsKolekcija.MovePrevious
    Next
    Me.cmdDelModel.Enabled = Me.ModelSelected
    Me.cmdEditModel.Enabled = Me.cmdDelModel.Enabled
    If Me.ctlListKolekcija.ListItems.Count > 0 Then
        Dim cTMP As MSComctlLib.ListItem
        For Each cTMP In Me.ctlListKolekcija.ListItems
            cTMP.SubItems(1) = cTMP.Index
        Next
        Me.ctlListKolekcija_ItemCheck Me.ctlListKolekcija.ListItems(1)
    Else
        Me.chkSelect.Value = 0
    End If
    Me.cmdUpload.Enabled = Me.cmdDownload.Enabled And Me.chkSelect.Value > 0
End If
End Sub


Private Sub cmdDelNarudzba_Click()
Dim i As Long
frmBrisati.Display "Da li ste sigurni da želite obrisati sve narudžbe klijenta " & Me.cboKlijenti & "?", Me
If frmBrisati.Odgovor = True Then
    For i = Me.ctlListNarudzbe.ListItems.Count To 1 Step -1
        rsNarudzbe.FindFirst "ID=" & Mid(Me.ctlListNarudzbe.ListItems(i).Key, 3)
        Me.ctlListNarudzbe.ListItems.Remove "ID" & rsNarudzbe("ID")
        rsNarudzbeBCKP.AddNew
        rsNarudzbeBCKP("KlijentID") = rsNarudzbe("KlijentID")
        rsNarudzbeBCKP("ModelID") = rsNarudzbe("ModelID")
        rsNarudzbeBCKP("Datum") = rsNarudzbe("Datum")
        rsNarudzbeBCKP("NarudzbaBroj") = rsNarudzbe("NarudzbaBroj")
        rsNarudzbeBCKP("NarudzbaComment") = rsNarudzbe("NarudzbaComment")
        rsNarudzbeBCKP.Update
        rsNarudzbe.Delete
    Next
End If
Me.cmdDelNarudzba.Enabled = False
Me.cmdPrintNarudzbe.Enabled = rsNarudzbe.RecordCount > 0
ctlListNarudzbe_ItemClick Me.ctlListNarudzbe.SelectedItem
End Sub


Private Sub cmdDownload_Click()
With Me
    .cmdConnect.Enabled = False
    .cmdUpload.Enabled = False
    .cmdAddModel.Enabled = False
    .cmdDelModel.Enabled = False
    .cmdEditModel.Enabled = False
    .chkSelect.Enabled = False
    .ctlListKolekcija.Enabled = False
    .cmdDownload.Enabled = False
    .cmdDelNarudzba.Enabled = False
    .cmdPrintNarudzbe.Enabled = False
End With
CurrentMode = MODE_NARUDZBA_INI_1
Me.ctlComm.Output = "GIVE narudzbe.dat*"
End Sub

Private Sub cmdEditModel_Click()
frmModel.EditModel
End Sub

Private Sub cmdPrintNarudzbe_Click()
frmPrint.Display
End Sub


Public Sub CollectFilesToSend()
Dim cTMP As New cFileToSend

If rsTasks.RecordCount = 0 Then GoTo w1
rsTasks.FindFirst "KlijentID=" & rsKlijenti("ID")
If rsTasks.NoMatch Then
w1:
    rsKolekcija.MoveFirst
    Do Until rsKolekcija.EOF
        If Me.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).Checked Then
            rsTasks.AddNew
            rsTasks("KlijentID") = rsKlijenti("ID")
            rsTasks("KolekcijaID") = rsKolekcija("ID")
            rsTasks.Update
        End If
        rsKolekcija.MoveNext
    Loop
    rsTasks.AddNew
    rsTasks("KlijentID") = rsKlijenti("ID")
    rsTasks("KolekcijaID") = 0
    rsTasks.Update
End If

Set cFajloviZaSlanje = Nothing
rsTasks.FindFirst "KlijentID=" & rsKlijenti("ID")
q1:
If rsTasks("KolekcijaID") <> 0 Then
    rsKolekcija.FindFirst "ID=" & rsTasks("KolekcijaID")
    cTMP.FileName = MyFolder & "pic\" & rsKolekcija("Photo")
    cTMP.FileSendString = "pic\" & rsKolekcija("Photo")
    cTMP.TaskID = rsTasks("ID")
    cFajloviZaSlanje.Add cTMP
    Set cTMP = Nothing
Else
    cTMP.FileName = MyFolder & "modeli.dat"
    cTMP.FileSendString = "modeli.dat"
    cTMP.TaskID = rsTasks("ID")
    cFajloviZaSlanje.Add cTMP
    Set cTMP = Nothing
End If

rsTasks.FindNext "KlijentID=" & rsKlijenti("ID")
If Not rsTasks.NoMatch Then GoTo q1

End Sub


Public Sub CheckReceiveSize()
Me.ctlComm.InputLen = 512
If FileSize - LOF(1) < 512 Then Me.ctlComm.InputLen = FileSize - LOF(1)
Me.ctlComm.RThreshold = Me.ctlComm.InputLen
End Sub


Private Sub cmdUpload_Click()
With Me
    .cmdConnect.Enabled = False
    .cmdUpload.Enabled = False
    .cmdAddModel.Enabled = False
    .cmdDelModel.Enabled = False
    .cmdEditModel.Enabled = False
    .chkSelect.Enabled = False
    .ctlListKolekcija.Enabled = False
    .cmdDownload.Enabled = False
    .cmdDelNarudzba.Enabled = False
    .cmdPrintNarudzbe.Enabled = False
    .ctlProgressBar.Value = 0
    .lblAkcija.Caption = "Šaljem kolekciju..."
    .Height = MainFormHeight + .ctlProgressBar.Height + .lblAkcija.Height + .lblAkcija.Left
    .ctlProgressBar.Visible = True
    .lblAkcija.Visible = True
End With
CurrentMode = MODE_TAKE_INI_1
Me.ctlComm.Output = "WIPE*"
End Sub

Private Sub ctlComm_OnComm()
Dim sTMP As String, bKomad() As Byte, bKomadSize As Long
If Me.ctlComm.CommEvent = comEvReceive Then
    Select Case CurrentMode
        Case MODE_INIT
            sTMP = Me.ctlComm.Input
            If sTMP = "|" Then
                CurrentMode = MODE_USER_INI
                Me.ctlComm.Output = "USER*"
            Else
                Me.ctlComm.Output = "*"
            End If
        Case MODE_USER_INI
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                CurrentMode = MODE_USER
                Me.ctlComm.Output = "*"
            Else
                CurrentMode = 0
                Me.ctlComm.PortOpen = False
                MsgBox "Došlo je do greške u komunikaciji!" & vbCrLf & "Pokušajte ponovo.", 16, "Poruèivanje obuæe"
                frmConnect.cmdConnect.Enabled = True
                frmConnect.cmdConnect.SetFocus
                frmConnect.cmdCancel.Caption = "Odustani"
                frmConnect.cmdCancel.ToolTipText = "Pritisnite da biste odustali od povezivanja"
            End If
        Case MODE_USER
            Buffer = Buffer & Me.ctlComm.Input
            If Right(Buffer, 1) = "*" Then
                sTMP = Left(Buffer, Len(Buffer) - 1)
                Buffer = ""
                CurrentMode = 0
                If sTMP = Me.cboKlijenti.Text Then
                    With Me
                        .cboKlijenti.Locked = True
                        .imgOn.ZOrder
                        .cmdConnect.Caption = "Prekini vezu"
                        .cmdConnect.ToolTipText = "Pritisnite da biste prekinuli vezu sa " & .cboKlijenti.Text
                        .cmdUpload.Enabled = .ctlListKolekcija.ListItems.Count > 0 And .chkSelect.Value > 0
                        .cmdDownload.Enabled = True
                    End With
                    frmConnect.cmdConnect.Enabled = True
                    frmConnect.cmdConnect.SetFocus
                    frmConnect.cmdCancel.Caption = "Odustani"
                    frmConnect.cmdCancel.ToolTipText = "Pritisnite da biste odustali od povezivanja"
                    frmConnect.Hide
                    CheckConnection
                Else
                    MsgBox "Povezali ste se sa pogrešnim klijentom!" & vbCrLf & "Izaberite drugog klijenta.", vbExclamation + vbOKOnly, "Poruèivanje obuæe"
                    Me.ctlComm.PortOpen = False
                    frmConnect.cmdConnect.Enabled = True
                    frmConnect.cmdConnect.SetFocus
                    frmConnect.cmdCancel.Caption = "Odustani"
                    frmConnect.cmdCancel.ToolTipText = "Pritisnite da biste odustali od povezivanja"
                    frmConnect.Hide
                End If
            End If
        Case MODE_NARUDZBA_INI_1
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                CurrentMode = MODE_NARUDZBA_INI_2
                Me.ctlComm.Output = "*"
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_NARUDZBA_INI_2
            Buffer = Buffer & Me.ctlComm.Input
            If Right(Buffer, 1) = "*" Then
                sTMP = Left(Buffer, Len(Buffer) - 1)
                Buffer = ""
                If sTMP = "ERR" Then
                    CurrentMode = 0
                    MsgBox "Narudžba kod klijenta " & Me.cboKlijenti.Text & " ne postoji!", vbInformation + vbOKOnly, "Poruèivanje obuæe"
                    With Me
                        .cmdConnect.Enabled = True
                        .cmdUpload.Enabled = .ctlListKolekcija.ListItems.Count > 0 And .chkSelect.Value > 0
                        .cmdAddModel.Enabled = True
                        .cmdDelModel.Enabled = .ModelSelected
                        .cmdEditModel.Enabled = .cmdDelModel.Enabled
                        .chkSelect.Enabled = True
                        .ctlListKolekcija.Enabled = True
                        .cmdDownload.Enabled = True
                        .cmdDelNarudzba.Enabled = .ctlListNarudzbe.ListItems.Count > 0
                        .cmdPrintNarudzbe.Enabled = rsNarudzbe.RecordCount > 0
                    End With
                ElseIf Right(sTMP, 2) = "OK" Then
                    sTMP = Left(sTMP, Len(sTMP) - 2)
                    FileSize = Val(sTMP)
                    If FileSize = 0 Then
                        TransferError = True
                        Me.ctlComm.PortOpen = False
                        Exit Sub
                    End If
                    If CStr(FileSize) = sTMP Then
                        Me.Height = MainFormHeight + Me.ctlProgressBar.Height + Me.lblAkcija.Height + Me.lblAkcija.Left
                        Me.lblAkcija.Caption = "Primam narudžbu..."
                        Me.ctlProgressBar.Value = 0
                        Me.lblAkcija.Visible = True
                        Me.ctlProgressBar.Visible = True
                        If FileExists(MyFolder & "narudzbe.dat") Then Kill MyFolder & "narudzbe.dat"
                        Open MyFolder & "narudzbe.dat" For Binary As #1
                        Me.ctlComm.InputMode = comInputModeBinary
                        CheckReceiveSize
                        CurrentMode = MODE_NARUDZBA
                        Me.ctlComm.Output = "*"
                    Else
                        TransferError = True
                        Me.ctlComm.PortOpen = False
                    End If
                Else
                    TransferError = True
                    Me.ctlComm.PortOpen = False
                End If
            End If
        Case MODE_NARUDZBA
            ReDim bKomad(Me.ctlComm.RThreshold - 1)
            bKomad = Me.ctlComm.Input
            Put #1, , bKomad
            Me.ctlProgressBar.Value = LOF(1) / FileSize * 100
            If LOF(1) = FileSize Then
                CurrentMode = 0
                Close #1
                Me.ctlComm.RThreshold = 1
                Me.ctlComm.InputLen = 0
                Me.ctlComm.InputMode = comInputModeText
                Dim m_lModelID As Long, m_sNarudzbaBroj As String, m_sNarudzbaComment As String
                Open MyFolder & "narudzbe.dat" For Input As #1
                    Do Until EOF(1)
                        Input #1, m_lModelID, m_sNarudzbaBroj, m_sNarudzbaComment
                        rsNarudzbe.AddNew
                        rsNarudzbe("KlijentID") = rsKlijenti("ID")
                        rsNarudzbe("ModelID") = m_lModelID
                        rsNarudzbe("NarudzbaBroj") = m_sNarudzbaBroj
                        rsNarudzbe("NarudzbaComment") = m_sNarudzbaComment
                        rsNarudzbe.Update
                    Loop
                Close #1
                Kill MyFolder & "narudzbe.dat"
                rsKlijenti.Edit
                rsKlijenti("NarudzbaDate") = Date
                rsKlijenti.Update
                cboKlijenti_Click
                CurrentMode = MODE_NARUDZBA_OFF
                Me.ctlComm.Output = "DELE narudzbe.dat*"
            Else
                CheckReceiveSize
                Me.ctlComm.Output = "*"
            End If
        Case MODE_NARUDZBA_OFF
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                CurrentMode = 0
                Me.ctlProgressBar.Visible = False
                Me.lblAkcija.Visible = False
                Me.Height = MainFormHeight
                With Me
                    .cmdConnect.Enabled = True
                    .cmdUpload.Enabled = .ctlListKolekcija.ListItems.Count > 0 And .chkSelect.Value > 0
                    .cmdAddModel.Enabled = True
                    .cmdDelModel.Enabled = .ModelSelected
                    .cmdEditModel.Enabled = .cmdDelModel.Enabled
                    .chkSelect.Enabled = True
                    .ctlListKolekcija.Enabled = True
                    .cmdDownload.Enabled = True
                    .cmdDelNarudzba.Enabled = .ctlListNarudzbe.ListItems.Count > 0
                    .cmdPrintNarudzbe.Enabled = rsNarudzbe.RecordCount > 0
                End With
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_TAKE_INI_1
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                CollectFilesToSend
                TrenutniFajl = 1
skok1:
                If cFajloviZaSlanje(TrenutniFajl).FileSendString = "modeli.dat" Then
                    Open MyFolder & "modeli.dat" For Output As #1
                        rsKolekcija.MoveFirst
                        Do Until rsKolekcija.EOF
                            If Me.ctlListKolekcija.ListItems("ID" & rsKolekcija("ID")).Checked Then
                                rsModeli.FindFirst "ID=" & rsKolekcija("ModelID")
                                Write #1, CLng(rsKolekcija("ModelID")), rsModeli("Model"), rsKolekcija("Photo"), CInt(rsModeli("Tip")), rsModeli("MatLica"), rsModeli("MatDjona"), rsModeli("Boja"), rsModeli("Sortiment"), CSng(rsModeli("Cijena")), CDate(rsModeli("Rok"))
                            End If
                            rsKolekcija.MoveNext
                        Loop
                    Close #1
                End If
                If Not FileExists(cFajloviZaSlanje(TrenutniFajl).FileName) Then
skok2:
                    rsTasks.FindFirst "ID=" & cFajloviZaSlanje(TrenutniFajl).TaskID
                    rsTasks.Delete
                    If TrenutniFajl < cFajloviZaSlanje.Count Then
                        TrenutniFajl = TrenutniFajl + 1
                        GoTo skok1
                    Else
skok3:
                        CurrentMode = 0
                        rsKlijenti.Edit
                        rsKlijenti("KolekcijaDate") = Date
                        rsKlijenti.Update
                        cboKlijenti_Click
                        Set cFajloviZaSlanje = Nothing
                        TrenutniFajl = 0
                        Me.ctlProgressBar.Visible = False
                        Me.lblAkcija.Visible = False
                        Me.Height = MainFormHeight
                        With Me
                            .cmdConnect.Enabled = True
                            .cmdUpload.Enabled = .ctlListKolekcija.ListItems.Count > 0 And .chkSelect.Value > 0
                            .cmdAddModel.Enabled = True
                            .cmdDelModel.Enabled = .ModelSelected
                            .cmdEditModel.Enabled = .cmdDelModel.Enabled
                            .chkSelect.Enabled = True
                            .ctlListKolekcija.Enabled = True
                            .cmdDownload.Enabled = True
                            .cmdDelNarudzba.Enabled = .ctlListNarudzbe.ListItems.Count > 0
                            .cmdPrintNarudzbe.Enabled = rsNarudzbe.RecordCount > 0
                        End With
                    End If
                Else
                    FileSize = FileLen(cFajloviZaSlanje(TrenutniFajl).FileName)
                    If FileSize = 0 Then
                        GoTo skok2
                    Else
                        CurrentMode = MODE_TAKE_INI_2
                        Me.ctlComm.Output = "TAKE " & cFajloviZaSlanje(TrenutniFajl).FileSendString & ":" & FileSize & "*"
                    End If
                End If
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_TAKE_INI_2
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                Me.lblAkcija.Caption = "Šaljem: " & cFajloviZaSlanje(TrenutniFajl).FileName & " (" & Round(FileSize / 1024, 1) & "KB)..."
                Me.ctlProgressBar.Value = 0
                Open cFajloviZaSlanje(TrenutniFajl).FileName For Binary As #1
                CurrentMode = MODE_TAKE
                Me.ctlComm.Output = "*"
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_TAKE
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                Me.ctlProgressBar.Value = (Seek(1) - 1) / FileSize * 100
                bKomadSize = FileSize - Seek(1) + 1
                If bKomadSize > 512 Then
                    ReDim bKomad(511)
                    Get #1, , bKomad
                    Me.ctlComm.Output = bKomad
                Else
                    ReDim bKomad(bKomadSize - 1)
                    Get #1, , bKomad
                    Close #1
                    CurrentMode = MODE_TAKE_OFF
                    Me.ctlComm.Output = bKomad
                End If
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_TAKE_OFF
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                Me.ctlProgressBar.Value = 100
                rsTasks.FindFirst "ID=" & cFajloviZaSlanje(TrenutniFajl).TaskID
                rsTasks.Delete
                If cFajloviZaSlanje(TrenutniFajl).FileSendString = "modeli.dat" Then Kill MyFolder & "modeli.dat"
                If TrenutniFajl < cFajloviZaSlanje.Count Then
                    TrenutniFajl = TrenutniFajl + 1
                    GoTo skok1
                Else
                    GoTo skok3
                End If
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_EXEC
            sTMP = Me.ctlComm.Input
            If sTMP = "*" Then
                CurrentMode = MODE_QUIT
                Me.ctlComm.Output = "QUIT*"
            Else
                TransferError = True
                Me.ctlComm.PortOpen = False
            End If
        Case MODE_QUIT
            sTMP = Me.ctlComm.Input
            CurrentMode = 0
            sTMP = Left(sTMP, 1)
            If sTMP = "*" Then
                QuitMode = True
            Else
                TransferError = True
            End If
            frmDisConnect.cmdHangUp.Enabled = True
            frmDisConnect.cmdCancel.Enabled = True
            frmDisConnect.chkExec.Enabled = True
            frmDisConnect.cmdHangUp.SetFocus
            frmDisConnect.Hide
            If TransferError Then Me.ctlComm.PortOpen = False
    End Select
End If
End Sub

Public Sub ctlListKolekcija_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim b As Integer, cTMP As MSComctlLib.ListItem
If Me.ctlListKolekcija.ListItems(1).Checked Then
    b = 1
Else
    b = 0
End If
For Each cTMP In Me.ctlListKolekcija.ListItems
    If (cTMP.Checked And b = 0) Or (Not cTMP.Checked And b = 1) Then
        b = 2
        Exit For
    End If
Next
Me.chkSelect.Value = b
Me.cmdUpload.Enabled = Me.cmdDownload.Enabled And Me.chkSelect.Value > 0
End Sub

Private Sub ctlListKolekcija_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.cmdDelModel.Enabled = Me.ModelSelected
Me.cmdEditModel.Enabled = Me.cmdDelModel.Enabled
End Sub

Private Sub PrikažiSortiment(sInput As String)
Dim sSort As String, i As Integer, s As String, sTMP As String, iPoljeCount As Integer, iZnakCount As Integer
sSort = sInput

Dim sBrojevi() As Integer, sZnakovi() As String, bZnakoviOK As Boolean, bBrojeviOK As Boolean
' ako dužina polja nije 2,5,8 itd. izlazi iz funkcije
If (Len(sSort) + 1) Mod 3 <> 0 Then GoTo greška
ReDim sBrojevi(((Len(sSort) + 1) / 3) - 1)
' popunjava brojeve
For i = 0 To UBound(sBrojevi)
    If Mid(sSort, 3 * i + 1, 2) <> CStr(Val(Mid(sSort, 3 * i + 1, 2))) Then
        GoTo greška
    Else
        sBrojevi(i) = Val(Mid(sSort, 3 * i + 1, 2))
    End If
Next
'popunjava znakove
If Len(sSort) > 4 Then
    ReDim sZnakovi(UBound(sBrojevi) - 1)
    For i = 0 To UBound(sZnakovi)
        sZnakovi(i) = Mid(sSort, (i + 1) * 3, 1)
        If sZnakovi(i) <> "-" And sZnakovi(i) <> "," Then GoTo greška
    Next
    
    If UBound(sZnakovi) = 0 Then
        bZnakoviOK = True
        GoTo n1
    End If
    bZnakoviOK = True
    For i = 0 To UBound(sZnakovi) - 1
        If sZnakovi(i) = "-" And sZnakovi(i + 1) = "-" Then GoTo greška
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
        If sBrojevi(i) >= sBrojevi(i + 1) Then GoTo greška
    Next
End If

If bBrojeviOK And bZnakoviOK Then
    Select Case UBound(sBrojevi)
        Case 0
            Me.lblBroj(0) = sBrojevi(0)
            Me.lblBroj(0).Visible = True
        Case Is > 0
            iPoljeCount = 0
            iZnakCount = 0
t1:
            If sZnakovi(iZnakCount) = "," Then
                Me.lblBroj(iPoljeCount) = sBrojevi(iZnakCount)
                Me.lblBroj(iPoljeCount).Visible = True
                iPoljeCount = iPoljeCount + 1
                If UBound(sZnakovi) > iZnakCount Then
                    iZnakCount = iZnakCount + 1
                    GoTo t1
                Else
                    Me.lblBroj(iPoljeCount) = sBrojevi(iZnakCount + 1)
                    Me.lblBroj(iPoljeCount).Visible = True
                End If
            Else
                For i = sBrojevi(iZnakCount) To sBrojevi(iZnakCount + 1)
                    Me.lblBroj(iPoljeCount) = i
                    Me.lblBroj(iPoljeCount).Visible = True
                    iPoljeCount = iPoljeCount + 1
                Next
                If UBound(sZnakovi) > (iZnakCount + 1) Then
                    iZnakCount = iZnakCount + 2
                    GoTo t1
                ElseIf UBound(sZnakovi) = (iZnakCount + 1) Then
                    Me.lblBroj(iPoljeCount) = sBrojevi(iZnakCount + 2)
                    Me.lblBroj(iPoljeCount).Visible = True
                End If
            End If
    End Select
Else
    GoTo greška
End If

For i = 0 To 39
    If Me.lblBroj(i Mod 10).Visible And Me.lblBoja(Int(i / 10)).Visible Then Me.txtKolicina(i).Visible = True
Next
Me.Label3.Visible = True
Exit Sub


greška:
MsgBox "NEISPRAVAN SORTIMENT!!!" & vbCrLf & _
        "U polje ""Dodatni komentar:"" upišite svoju narudžbu.", vbCritical + vbOKOnly, "Poruèivanje obuæe"

End Sub

Private Sub PrikažiBoje(sInput As String)
Dim sBoja(3) As String, sTMP As String, iCount As Integer, i As Integer
iCount = 0
sTMP = sInput

Opet:

If InStr(sTMP, ",") = 0 Then
    sBoja(iCount) = sTMP
Else
    sBoja(iCount) = Left(sTMP, InStr(sTMP, ",") - 1)
    If iCount = 3 Then GoTo over
    iCount = iCount + 1
    sTMP = Mid(sTMP, InStr(sTMP, ",") + 1)
    GoTo Opet
End If
    
over:
For i = 0 To iCount
    Me.lblBoja(i).Caption = sBoja(i) & ":"
    Me.lblBoja(i).Visible = True
Next

End Sub

Private Sub PrikažiTabelu()
Dim i As Integer
For i = 0 To 39
    If Me.lblBroj(i Mod 10).Visible And Me.lblBoja(Int(i / 10)).Visible Then Me.txtKolicina(i).Visible = True
Next
End Sub

Private Sub ctlListNarudzbe_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim sTMP As String, iCount As Integer, sNBS(39) As Integer, i As Integer

For i = 0 To 9
    Me.lblBroj(i).Visible = False
    Me.txtKolicina(4 * i).Visible = False
    Me.txtKolicina(4 * i) = ""
    Me.txtKolicina(4 * i + 1).Visible = False
    Me.txtKolicina(4 * i + 1) = ""
    Me.txtKolicina(4 * i + 2).Visible = False
    Me.txtKolicina(4 * i + 2) = ""
    Me.txtKolicina(4 * i + 3).Visible = False
    Me.txtKolicina(4 * i + 3) = ""
Next
Me.txtKomentar = ""
Me.txtKomentar.Visible = False

For i = 0 To 3
    Me.lblBoja(i).Visible = False
Next

Me.Label2.Visible = False
Me.Label3.Visible = False

If NarudzbaSelected Then
    rsNarudzbe.FindFirst "ID=" & Mid(Me.ctlListNarudzbe.SelectedItem.Key, 3)
    rsModeli.FindFirst "ID=" & rsNarudzbe("ModelID")
    PrikažiSortiment rsModeli("Sortiment")
    PrikažiBoje rsModeli("Boja")
    PrikažiTabelu
    
    sTMP = rsNarudzbe("NarudzbaBroj")
Opet:
    If InStr(sTMP, "|") = 0 Then
        sNBS(iCount) = Val(sTMP)
    Else
        sNBS(iCount) = Val(Left(sTMP, InStr(sTMP, "|") - 1))
        iCount = iCount + 1
        sTMP = Mid(sTMP, InStr(sTMP, "|") + 1)
        GoTo Opet
    End If
    
    For i = 0 To 39
        If sNBS(i) > 0 Then Me.txtKolicina(i) = sNBS(i)
    Next
           
    If Not IsNull(rsNarudzbe("NarudzbaComment")) Then
        Me.txtKomentar.Text = rsNarudzbe("NarudzbaComment")
        If Me.txtKomentar <> "" Then
            Me.Label2.Visible = True
            Me.txtKomentar.Visible = True
        End If
    End If

End If
Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim a As Long
If Me.cmdDownload.Enabled Then
    Cancel = True
    Exit Sub
End If
If rsTasks.RecordCount > 0 Then
    rsTasks.MoveFirst
    rsKlijenti.FindFirst "ID=" & rsTasks("KlijentID")
    frmBrisati.Display "Sa klijentom " & rsKlijenti("Klijent") & " nije obavljen sav prenos!" & _
            vbCrLf & "Želite li prekinuti sa radom?", Me
    If frmBrisati.Odgovor = False Then Cancel = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsModeli = Nothing
Set rsKlijenti = Nothing
Set rsKolekcija = Nothing
Set rsNarudzbe = Nothing
Set dbData = Nothing
End Sub


Public Sub CheckConnection()
Dim k As Date
Do While Me.ctlComm.CDHolding
    DoEvents
Loop
If QuitMode Then GoTo k1
If TransferError Then
    MsgBox "Došlo je do greške u komunikaciji!", 16, "Poruèivanje obuæe"
Else
    MsgBox "Došlo je do prekida veze!", 16, "Poruèivanje obuæe"
End If
k1:
If Me.ctlComm.PortOpen Then Me.ctlComm.PortOpen = False
With Me
    .cboKlijenti.Locked = False
    .imgOff.ZOrder
    .cmdConnect.Caption = "Poveži se"
    .cmdConnect.Enabled = True
    .cmdConnect.ToolTipText = "Pritisnite da biste se povezali sa klijentom"
    .cmdUpload.Enabled = False
    .cmdAddModel.Enabled = True
    .cmdDelModel.Enabled = .ModelSelected
    .cmdEditModel.Enabled = .cmdDelModel.Enabled
    .chkSelect.Enabled = True
    .ctlListKolekcija.Enabled = True
    .cmdDownload.Enabled = False
    .cmdDelNarudzba.Enabled = .ctlListNarudzbe.ListItems.Count > 0
    .cmdPrintNarudzbe.Enabled = rsNarudzbe.RecordCount > 0
    .Height = MainFormHeight
    .lblAkcija.Visible = False
    .ctlProgressBar.Visible = False
End With
Close #1
CurrentMode = 0
Buffer = ""
FileSize = 0
TrenutniFajl = 0
TransferError = False
QuitMode = False
Me.ctlComm.RThreshold = 1
Me.ctlComm.InputLen = 0
Me.ctlComm.InputMode = comInputModeText
Set cFajloviZaSlanje = Nothing
End Sub


Public Sub LocirajKlijenta()
rsKlijenti.MoveFirst
Do Until rsKlijenti.EOF
    If rsKlijenti("Klijent") = Me.cboKlijenti Then Exit Do
    rsKlijenti.MoveNext
Loop
End Sub


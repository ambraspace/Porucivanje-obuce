VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNarudzba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lModelID As Long
Private m_sNarudzbaBroj As String
Private m_sNarudzbaComment As String

Public Property Get ModelID() As Long
ModelID = m_lModelID
End Property
Public Property Let ModelID(lInput As Long)
m_lModelID = lInput
End Property

Public Property Get NarudzbaBroj() As String
NarudzbaBroj = m_sNarudzbaBroj
End Property
Public Property Let NarudzbaBroj(sInput As String)
m_sNarudzbaBroj = sInput
End Property

Public Property Get NarudzbaComment() As String
NarudzbaComment = m_sNarudzbaComment
End Property
Public Property Let NarudzbaComment(sInput As String)
m_sNarudzbaComment = sInput
End Property



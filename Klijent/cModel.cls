VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lModelID As Long
Private m_sModel As String
Private m_sSlika As String
Private m_iTip As Integer
Private m_sMatLica As String
Private m_sMatDjona As String
Private m_sBoja As String
Private m_sSortiment As String
Private m_sCijena As Single
Private m_dRok As Date

Public Property Get ModelID() As Long
ModelID = m_lModelID
End Property
Public Property Let ModelID(lInput As Long)
m_lModelID = lInput
End Property

Public Property Get Model() As String
Model = m_sModel
End Property
Public Property Let Model(sInput As String)
m_sModel = sInput
End Property

Public Property Get Slika() As String
Slika = m_sSlika
End Property
Public Property Let Slika(sInput As String)
m_sSlika = sInput
End Property

Public Property Get Tip() As Integer
Tip = m_iTip
End Property
Public Property Let Tip(iInput As Integer)
m_iTip = iInput
End Property

Public Property Get MatLica() As String
MatLica = m_sMatLica
End Property
Public Property Let MatLica(sInput As String)
m_sMatLica = sInput
End Property

Public Property Get MatDjona() As String
MatDjona = m_sMatDjona
End Property
Public Property Let MatDjona(sInput As String)
m_sMatDjona = sInput
End Property

Public Property Get Boja() As String
Boja = m_sBoja
End Property
Public Property Let Boja(sInput As String)
m_sBoja = sInput
End Property

Public Property Get Sortiment() As String
Sortiment = m_sSortiment
End Property
Public Property Let Sortiment(sInput As String)
m_sSortiment = sInput
End Property

Public Property Get Cijena() As Single
Cijena = m_sCijena
End Property
Public Property Let Cijena(sInput As Single)
m_sCijena = sInput
End Property

Public Property Get Rok() As Date
Rok = m_dRok
End Property
Public Property Let Rok(dInput As Date)
m_dRok = dInput
End Property

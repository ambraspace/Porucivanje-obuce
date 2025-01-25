VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileToSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sFileName As String
Private m_sFileSendString As String
Private m_lTaskID As Long

Public Property Get FileName() As String
FileName = m_sFileName
End Property
Public Property Let FileName(sInput As String)
m_sFileName = sInput
End Property

Public Property Get FileSendString() As String
FileSendString = m_sFileSendString
End Property
Public Property Let FileSendString(sInput As String)
m_sFileSendString = sInput
End Property

Public Property Get TaskID() As Long
TaskID = m_lTaskID
End Property
Public Property Let TaskID(lInput As Long)
m_lTaskID = lInput
End Property


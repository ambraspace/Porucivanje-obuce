Attribute VB_Name = "Dial"
Option Explicit

Private Const RAS95_MaxEntryName = 256
Private Const RAS_MaxPhoneNumber = 128
Private Const RAS_MaxCallbackNumber = RAS_MaxPhoneNumber
Private Const UNLEN = 256
Private Const PWLEN = 256
Private Const DNLEN = 12

Private Type RASDIALPARAMS
    dwSize As Long
    szEntryName(RAS95_MaxEntryName) As Byte
    szPhoneNumber(RAS_MaxPhoneNumber) As Byte
    szCallbackNumber(RAS_MaxCallbackNumber) As Byte
    szUserName(UNLEN) As Byte
    szPassword(PWLEN) As Byte
    szDomain(DNLEN) As Byte
End Type

Private Declare Function RasDial Lib "rasapi32.dll" Alias "RasDialA" (ByVal lpRasDialExtensions As Long, ByVal lpCstr As String, ByRef lpRadDialParamsa As RASDIALPARAMS, ByVal dword As Long, lpVoid As Any, ByRef lpHRasConn As Long) As Long


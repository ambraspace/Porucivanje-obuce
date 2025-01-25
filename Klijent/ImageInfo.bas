Attribute VB_Name = "ImageInfo"
'I have released this source code into the public domain.  You may use it
'with no strings attached.
'Just call GetImageSize with a string containing the filename, and
'it will return a user defined type 'ImageSize'  (see below)
'Return values of 0 indicate an error of some sort.  The error handling
'in this module is limited.  There is *NO* error handling on the test
'form.  This routine is limited to X or Y sizes of 32767 pixels, but that
'should not be a problem.

'Check back at http://www.qtm.net/~davidc
'I may add support for more file types.

'supported in this version:
'JPEG
'GIF
'PNG

'This routine does not require any royalty fees for Unisys as it
'does nothing with the compressed part of GIF files.  It simply reads
'4 bytes to determine image size.

Option Explicit
Public Type ImageSize
    Width As Long
    Height As Long
End Type

Public Function GetImageSize(sFileName As String) As ImageSize
    On Error Resume Next        'you'll want to change this
    Dim iFN As Integer
    Dim bTemp(3) As Byte
    Dim lFlen As Long
    Dim lPos As Long
    Dim bHmsb As Byte
    Dim bHlsb As Byte
    Dim bWmsb As Byte
    Dim bWlsb As Byte
    Dim bBuf(7) As Byte
    Dim bDone As Byte
    Dim iCount As Integer

    lFlen = FileLen(sFileName)
    iFN = FreeFile
    Open sFileName For Binary As iFN
    Get #iFN, 1, bTemp()
    'JPEG file
    If bTemp(0) = &HFF And bTemp(1) = &HD8 And bTemp(2) = &HFF Then
        lPos = 3
        Do
            Do
                Get #iFN, lPos, bBuf(1)
                Get #iFN, lPos + 1, bBuf(2)
                lPos = lPos + 1
            Loop Until (bBuf(1) = &HFF And bBuf(2) <> &HFF) Or lPos > lFlen
        
            For iCount = 0 To 7
                Get #iFN, lPos + iCount, bBuf(iCount)
            Next iCount
            If bBuf(0) >= &HC0 And bBuf(0) <= &HC3 Then
                bHmsb = bBuf(4)
                bHlsb = bBuf(5)
                bWmsb = bBuf(6)
                bWlsb = bBuf(7)
                bDone = 1
            Else
                lPos = lPos + (CombineBytes(bBuf(2), bBuf(1))) + 1
            End If
        Loop While lPos < lFlen And bDone = 0
        GetImageSize.Width = CombineBytes(bWlsb, bWmsb)
        GetImageSize.Height = CombineBytes(bHlsb, bHmsb)
        
    End If
    Close iFN
    
End Function
Private Function CombineBytes(lsb As Byte, msb As Byte) As Long
    CombineBytes = CLng(lsb + (msb * 256))
End Function



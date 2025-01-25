Attribute VB_Name = "mBase64"
Option Explicit

Private Const Base64String As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-/"

Private Const l_262144 As Long = 262144
Private Const l_65536 As Long = 65536
Private Const l_4096 As Long = 4096
Private Const l_256 As Long = 256
Private Const l_64 As Long = 64
Private Const l_16 As Long = 16
Private Const l_4 As Long = 4


Public Function Base64Encode(b() As Byte, bLen As Long) As String
Dim i As Long, t1 As Long, t2 As Long, t3 As Long, t4 As Long, Product As Long, bCount As Long
bCount = bLen
If (bLen Mod 3) <> 0 Then
    ReDim Preserve b(3 * Int(bLen / 3) + 2)
    bCount = 3 * Int(bLen / 3) + 3
End If
For i = 1 To bCount / 3
    Product = b(i * 3 - 3) * l_65536 + b(i * 3 - 2) * l_256 + b(i * 3 - 1)
    t1 = Int(Product / l_262144)
    Product = Product - (t1 * l_262144)
    t2 = Int(Product / l_4096)
    Product = Product - (t2 * l_4096)
    t3 = Int(Product / l_64)
    t4 = Product - (l_64 * t3)
    Base64Encode = Base64Encode & Mid(Base64String, t1 + 1, 1) & Mid(Base64String, t2 + 1, 1) & Mid(Base64String, t3 + 1, 1) & Mid(Base64String, t4 + 1, 1)
Next
Select Case bLen Mod 3
    Case 1
        Base64Encode = Left(Base64Encode, Len(Base64Encode) - 2) & "=="
    Case 2
        Base64Encode = Left(Base64Encode, Len(Base64Encode) - 1) & "="
End Select
End Function



Public Sub Base64Decode(s As String, b() As Byte)
Dim Product As Long, i As Integer
ReDim b(2)
For i = 1 To 4
    If Mid(s, i, 1) <> "=" Then
        Product = Product + ((InStr(Base64String, Mid(s, i, 1)) - 1) * (l_64 ^ (4 - i)))
    Else
        Select Case i
            Case 3
                ReDim b(0)
                b(0) = Product / l_16
                Exit Sub
            Case 4
                ReDim b(1)
                Product = Product / l_4
                b(0) = Int(Product / l_256)
                b(1) = Product - (b(0) * l_256)
                Exit Sub
        End Select
    End If
Next
b(0) = Int(Product / l_65536)
Product = Product - (b(0) * l_65536)
b(1) = Int(Product / l_256)
b(2) = Product - (b(1) * l_256)
End Sub



Public Function Base64Size(lFileSize As Long) As Long
Select Case lFileSize Mod 3
    Case 0
        Base64Size = Int(lFileSize / 3) * 4
    Case Else
        Base64Size = Int(lFileSize / 3) * 4 + 4
End Select
End Function


Attribute VB_Name = "MSingle"
Option Explicit
Private Type TLong
    value As Long
End Type
Private Type TSingle
    value As Single
End Type
'#define undenormalise(sample) if(((*(unsigned int*)&sample)&0x7f800000)==0) sample=0.0f

Function flush_to_zero(ByVal value As Single) As Single
    Dim ts As TSingle: ts.value = value
    Dim tl As TLong:    LSet tl = ts
    If ((tl.value And &H7F800000) = 0) Then flush_to_zero = 0 Else flush_to_zero = value
End Function

Sub undenormalise(ByRef outValue As Single)
    If outValue = 0 Then Exit Sub
    Dim ts As TSingle: ts.value = outValue
    Dim tl As TLong:    LSet tl = ts
    If ((tl.value And &H7F800000) = 0) Then outValue = 0
End Sub

Public Function MaxSng(ByVal V1 As Single, ByVal V2 As Single) As Single
    If V1 > V2 Then MaxSng = V1 Else MaxSng = V2
End Function

Public Function MaxSngArr(SngArr() As Single) As Single
    
    Dim i As Long
    For i = 0 To UBound(SngArr) - 1
        If SngArr(i) > MaxSngArr Then MaxSngArr = SngArr(i)
    Next

End Function

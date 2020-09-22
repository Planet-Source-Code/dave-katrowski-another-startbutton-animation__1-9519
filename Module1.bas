Attribute VB_Name = "Module1"
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020

Function FindChildByClass(Parent As Long, Class As String) As Long
Static Temp As String, Class_ As String, CHILD As Long, Child_ As Long

Temp = String(255, 0): CHILD = GetWindow(Parent, 5)

Do: DoEvents
Class_ = Left(Temp, GetClassName(CHILD, Temp, 255))
If LCase(Class) = LCase(Class_) Then FindChildByClass = CHILD: Exit Function
If GetWindow(CHILD, 5) <> 0 Then
Child_ = FindChildByClass(CHILD, Class)
If Child_ <> 0 Then FindChildByClass = Child_: Exit Function
End If
CHILD = GetNextWindow(CHILD, 2)
Loop Until CHILD = 0

FindChildByClass = 0
End Function

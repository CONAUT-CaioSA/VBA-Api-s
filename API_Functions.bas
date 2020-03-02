Attribute VB_Name = "API_Functions"
Private Declare PtrSafe Function FindWindowA Lib "user32.dll" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowLongA Lib "user32.dll" (ByVal hWnd As LongPtr, ByVal nINDEX As Integer) As LongPtr
Private Declare PtrSafe Function SetWindowLongA Lib "user32.dll" (ByVal hWnd As LongPtr, ByVal nINDEX As Integer, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As LongPtr, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As LongPtr

Private Const GWL_STYLE = -16
Private Const GWL_EXSTYLE = (-20)
Private Const WS_SIZEBOX = &H40000
Private Const WS_USER  As Long = &H4000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Public Sub IncreaseElements(Caption As String)
    
    Dim strCaption As String
    Dim strClass As String
    Dim LngMaximize As LongPtr
 
    'class name changed in Office 2000
    If Val(Application.Version) >= 9 Then
        strClass = "ThunderDFrame"
    Else
        strClass = "ThunderXFrame"
    End If
    
    
    mlnghWnd = FindWindowA(strClass, Caption)
    mIntStyle = GetWindowLongA(mlnghWnd, GWL_STYLE)
    mIntStyle = mIntStyle Or WS_SIZEBOX Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX
    LngMaximize = SetWindowLongA(mlnghWnd, GWL_STYLE, mIntStyle)
    
End Sub
Public Sub SetOpacity(Object As Msforms.Control, Opacity As Byte)
    
    Dim mlnghWnd As LongPtr
    Dim mIntStyle As LongPtr
    Dim mlngOpacitySet As LongPtr
    Dim mlngOpacity As LongPtr
    
    mlnghWnd = Object.[_GethWnd]
    mIntStyle = GetWindowLongA(mlnghWnd, GWL_EXSTYLE)
    mlngOpacitySet = SetWindowLongA(mlnghWnd, GWL_EXSTYLE, mIntStyle Or WS_EX_LAYERED)
    mlngOpacity = SetLayeredWindowAttributes(mlnghWnd, 0, Opacity, LWA_ALPHA)
End Sub


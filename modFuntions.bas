Attribute VB_Name = "Module1"
'***********************************************'
' The New iSoftware Company                     '
'   Control Copier and Control Resizer Module   '
'                                Version 1.0    '
'                                               '
' Based on "Creating Controls At Runtime"       '
'          by John Smiley                       '
'***********************************************'

'  Notes for Handles:
'    The order for the handles is:
'           0  1  2
'           7     3        ^
'           6  5  4      North


'API calls and constants for resizablity of handles:
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const MousePress = &HA1
Public Const SizeN = 12
Public Const SizeS = 15
Public Const SizeW = 10
Public Const SizeE = 11
Public Const SizeNW = 13
Public Const SizeSW = 16
Public Const SizeNE = 14
Public Const SizeSE = 17

'Function CopyControl, no API
Public Function CopyControl(Control As Variant, Visible As Boolean, Top As Integer, Left As Integer, Width As Integer, Height As Integer)
    Dim NewIndex As Integer
    NewIndex = Control.Count + 1
    Load Control(NewIndex)
    With Control(NewIndex)
        .Visible = Visible
        .Top = Top
        .Left = Left
        .Width = Width
        .Height = Height
    End With
End Function

'Function CopyControlWithResize, optional API
Public Function CopyControlWithResize(Control As Variant, Visible As Boolean, Resize As Boolean, Handles As Variant, Top As Integer, Left As Integer, Width As Integer, Height As Integer)
    Dim NewIndex As Integer
    NewIndex = Control.Count + 1
    Load Control(NewIndex)
    With Control(NewIndex)
        .Visible = Visible
        .Top = Top
        .Left = Left
        .Width = Width
        .Height = Height
    End With
    If Resize = True Then
    X = 0
    Do Until X = 8
    Handles(X).Visible = True
    X = X + 1
    Loop
    HandlesMove Control(NewIndex), Handles
    End If
End Function

'Function ControlResize, API
Public Function ControlResize(ControlWithAPIHandle As Control, Handles As Variant, Index As Variant)
    ReleaseCapture
    Select Case Index
        Case 0
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeNW, 0
        Case 1
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeN, 0
        Case 2
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeNE, 0
        Case 3
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeE, 0
        Case 4
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeSE, 0
        Case 5
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeS, 0
        Case 6
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeSW, 0
        Case 7
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeW, 0
    End Select
    HandlesMove ControlWithAPIHandle, Handles
End Function

'Function HandlesMove, API
Public Function HandlesMove(ByVal Control As Control, Handles As Variant)
    Handles(0).Left = Control.Left - Handles(0).Width
    Handles(0).Top = Control.Top - Handles(0).Height
    Handles(1).Left = (Control.Width - Handles(3).Width) / 2 + Control.Left
    Handles(1).Top = Control.Top - Handles(1).Height
    Handles(2).Left = Control.Left + Control.Width
    Handles(2).Top = Control.Top - Handles(0).Height
    Handles(3).Left = Control.Left + Control.Width
    Handles(3).Top = (Control.Height - Handles(3).Height) / 2 + Control.Top
    Handles(4).Left = Control.Left + Control.Width
    Handles(4).Top = Control.Top + Control.Height
    Handles(5).Left = (Control.Width - Handles(5).Width) / 2 + Control.Left
    Handles(5).Top = Control.Top + Control.Height
    Handles(6).Left = Control.Left - Handles(6).Width
    Handles(6).Top = Control.Top + Control.Height
    Handles(7).Left = Control.Left - Handles(7).Width
    Handles(7).Top = (Control.Height - Handles(7).Height) / 2 + Control.Top
End Function


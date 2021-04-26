Attribute VB_Name = "MouseTrackModule"

Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As tagTRACKMOUSEEVENT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Const TME_HOVER = &H1
Private Const TME_LEAVE = &H2
Private Const TME_CANCEL = &H80000000
Private Const HOVER_DEFAULT = &HFFFFFFFF
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_MOUSELEAVE = &H2A3
Private Const WM_MOUSEHOVER = &H2A1
Private Const WM_MOUSEMOVE = &H200
Private Const GWL_WNDPROC = (-4)

Private Type WNDHInfo
    hWnd As Long
    isHovered As Boolean
    hasLimitedScrolling As Boolean
    MHControl As MouseTrackControl
    MHhWnd As Long
End Type
Private Type tagTRACKMOUSEEVENT
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Dim WNDHArray() As WNDHInfo
Dim WNDHTotal As Integer

Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim WNDHEvent As Integer
Dim WNDHIndex As Integer
Dim WNDHResult As Boolean

Dim MouseKeys As Long
Dim MouseRotation As Long
Dim MousePosX As Long
Dim MousePosY As Long

Select Case Lmsg
    Case WM_MOUSEMOVE:  WNDHEvent = 1
    Case WM_MOUSEHOVER: WNDHEvent = 2
    Case WM_MOUSELEAVE: WNDHEvent = 3
    Case WM_MOUSEWHEEL: WNDHEvent = 4
    Case Else:          WNDHEvent = 0
End Select

If WNDHEvent > 0 Then
    For WNDHIndex = 1 To WNDHTotal
        If WNDHArray(WNDHIndex).hWnd = Lwnd Then

            If WNDHEvent = 4 Then
            
                MouseKeys = wParam And 65535
                MouseRotation = wParam / 65536
                MousePosX = lParam And 65535
                MousePosY = lParam / 65536

                WNDHResult = Not WNDHArray(WNDHIndex).hasLimitedScrolling

                If WNDHResult = False Then WNDHResult = ReturnIsMouseOver(Lwnd, MousePosX, MousePosY)

                If WNDHResult = True Then
                    WNDHArray(WNDHIndex).MHControl.RaiseWheelScrollEvent Lwnd, MouseKeys, MouseRotation, MousePosX, MousePosY
                End If
                
            Else
            
                If WNDHEvent = 3 Then
                    WNDHResult = False
                Else
                    WNDHResult = True
                End If
            
                If WNDHArray(WNDHIndex).isHovered <> WNDHResult Then
                    WNDHArray(WNDHIndex).isHovered = WNDHResult
                    WNDHArray(WNDHIndex).MHControl.RaiseHoverChangeEvent Lwnd, WNDHResult
                End If
            
                If WNDHEvent = 1 Then RequestTracking Lwnd
            
            End If
            
        End If
     Next
End If

WindowProc = CallWindowProc(GetProp(Lwnd, "PrevWndProc"), Lwnd, Lmsg, wParam, lParam)

End Function

Public Sub DoTrackHook(ByVal hWnd As Long, ByVal MHControl As MouseTrackControl, ByVal MHhWnd As Long)

If hWnd = 0 Then Exit Sub

Call DoTrackUnhook(hWnd, MHhWnd)

On Error Resume Next

SetProp hWnd, "PrevWndProc", SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)

RequestTracking hWnd

WNDHTotal = WNDHTotal + 1

ReDim Preserve WNDHArray(WNDHTotal)

WNDHArray(WNDHTotal).hWnd = hWnd
WNDHArray(WNDHTotal).isHovered = False
Set WNDHArray(WNDHTotal).MHControl = MHControl
WNDHArray(WNDHTotal).MHhWnd = MHhWnd

End Sub

Private Function RequestTracking(hWnd As Long)

Dim trk As tagTRACKMOUSEEVENT

trk.cbSize = 16
trk.dwFlags = TME_LEAVE Or TME_HOVER
trk.dwHoverTime = 400
trk.hwndTrack = hWnd

TrackMouseEvent trk

End Function

Public Sub DoTrackUnhook(ByVal hWnd As Long, ByVal MHhWnd As Long)

If hWnd = 0 Then Exit Sub

On Error Resume Next

Dim WNDHIndex As Integer
Dim WNDResize As Integer

For WNDHIndex = 1 To WNDHTotal
    If WNDHArray(WNDHIndex).hWnd = hWnd And WNDHArray(WNDHIndex).MHhWnd = MHhWnd Then

        Dim trk As tagTRACKMOUSEEVENT

        SetWindowLong hWnd, GWL_WNDPROC, GetProp(hWnd, "PrevWndProc")

        trk.cbSize = 16
        trk.dwFlags = TME_LEAVE Or TME_HOVER Or TME_CANCEL
        trk.hwndTrack = hWnd

        TrackMouseEvent trk

        RemoveProp hWnd, "PrevWndProc"

        WNDHTotal = WNDHTotal - 1
        For WNDResize = WNDHIndex To WNDHTotal
            WNDHArray(WNDResize) = WNDHArray(WNDResize + 1)
        Next

        ReDim Preserve WNDHArray(WNDHTotal)

        Exit For

    End If
Next

End Sub

Public Sub DoTrackAllUnhook(ByVal MHhWnd As Long)

Dim WNDHIndex As Integer
Dim hWnd As Long

On Error Resume Next

For WNDHIndex = WNDHTotal To 1 Step -1
    If WNDHArray(WNDHIndex).MHhWnd = MHhWnd Then

        hWnd = WNDHArray(WNDHIndex).hWnd

        Call DoTrackUnhook(hWnd, MHhWnd)

    End If
Next

End Sub

Public Function DoHookMeCount(ByVal MHhWnd As Long) As Long

Dim WNDHIndex As Integer
Dim WNDCount As Integer

On Error Resume Next

WNDCount = 0
For WNDHIndex = WNDHTotal To 1 Step -1
    If WNDHArray(WNDHIndex).MHhWnd = MHhWnd Then WNDCount = WNDCount + 1
Next

DoHookMeCount = WNDCount

End Function

Public Function DoHookTotal() As Long

DoHookTotal = WNDHTotal

End Function

Public Function ReturnIsMouseOver(ByVal hWnd As Long, ByVal MousePosX As Long, ByVal MousePosY As Long) As Boolean

Dim ControlRect As RECT

GetWindowRect hWnd, ControlRect

With ControlRect
    ReturnIsMouseOver = (MousePosX >= .Left And MousePosX <= .Right And MousePosY >= .Top And MousePosY <= .Bottom)
End With

End Function

Public Sub DoChangeWheelScrollingWhenHoveringOnly(ByVal MHhWnd As Long, ChangeFlag As Boolean)

Dim WNDHIndex As Integer

On Error Resume Next

For WNDHIndex = WNDHTotal To 1 Step -1
    If WNDHArray(WNDHIndex).MHhWnd = MHhWnd Then

        WNDHArray(WNDHIndex).hasLimitedScrolling = ChangeFlag

    End If
Next

End Sub

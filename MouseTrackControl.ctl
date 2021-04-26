VERSION 5.00
Begin VB.UserControl MouseTrackControl 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   210
   InvisibleAtRuntime=   -1  'True
   Picture         =   "MouseTrackControl.ctx":0000
   ScaleHeight     =   210
   ScaleWidth      =   210
   ToolboxBitmap   =   "MouseTrackControl.ctx":0312
End
Attribute VB_Name = "MouseTrackControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Event HoverChange(ByVal hWnd As Long, isHovered As Boolean)
Public Event WheelScroll(ByVal hWnd As Long, MouseKeys As Long, MouseRotation As Long, MousePosX As Long, MousePosY As Long)

Private Sub UserControl_Resize()

UserControl.Width = 250
UserControl.Height = 250

End Sub

Private Sub UserControl_Terminate()

Call DoTrackAllUnhook(UserControl.hWnd)

End Sub

Public Sub ChangeWheelScrollingWhenHoveringOnly(ChangeFlag As Boolean)

Call DoChangeWheelScrollingWhenHoveringOnly(UserControl.hWnd, ChangeFlag)

End Sub

Public Sub TrackHook(ByVal hWnd As Long)

Call DoTrackHook(hWnd, Me, UserControl.hWnd)

End Sub

Public Sub TrackUnhook(ByVal hWnd As Long)

Call DoTrackUnhook(hWnd, UserControl.hWnd)

End Sub

Public Sub TrackAllUnhook()

Call DoTrackAllUnhook(UserControl.hWnd)

End Sub

Public Function HookMeCount() As Integer

HookMeCount = DoHookMeCount(UserControl.hWnd)

End Function

Public Function HookTotal() As Integer

HookTotal = DoHookTotal()

End Function

Friend Sub RaiseHoverChangeEvent(ByVal hWnd As Long, isHovered As Boolean)

RaiseEvent HoverChange(hWnd, isHovered)

End Sub

Friend Sub RaiseWheelScrollEvent(ByVal hWnd As Long, MouseKeys As Long, MouseRotation As Long, MousePosX As Long, MousePosY As Long)

RaiseEvent WheelScroll(hWnd, MouseKeys, MouseRotation, MousePosX, MousePosY)

End Sub

VERSION 5.00
Object = "{3D7C563F-B005-4DE9-A960-2E648FC472D2}#7.0#0"; "sbmtc1.ocx"
Begin VB.Form TestFormB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TestFormB"
   ClientHeight    =   7500
   ClientLeft      =   10860
   ClientTop       =   2175
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   4215
   Begin VB.CheckBox TestScrolling 
      Caption         =   "Wheel Scrolling Only When Hovering"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   7200
      Width           =   3975
   End
   Begin VB.ListBox TestLogging 
      Height          =   2595
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   3975
   End
   Begin VB.TextBox TestText1 
      BackColor       =   &H00C0C0FF&
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Text            =   "TestText1"
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox TestText2 
      BackColor       =   &H00C0C0FF&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Text            =   "TestText2(0)"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox TestText2 
      BackColor       =   &H00C0C0FF&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Text            =   "TestText2(1)"
      Top             =   1665
      Width           =   1815
   End
   Begin VB.ListBox TestListbox 
      BackColor       =   &H00C0C0FF&
      Height          =   1815
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox TestCombo 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Text            =   "TestCombo"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton TestButton 
      BackColor       =   &H00C0C0FF&
      Caption         =   "TestButton"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.FileListBox TestFileListBox 
      BackColor       =   &H00C0C0FF&
      Height          =   870
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
   End
   Begin VB.DriveListBox TestDriveListBox 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CheckBox TestCheck 
      BackColor       =   &H00C0C0FF&
      Caption         =   "TestCheck"
      Height          =   255
      Left            =   2280
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   2
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton TestOption 
      BackColor       =   &H00C0C0FF&
      Caption         =   "TestOption(0)"
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Top             =   3720
      Width           =   1815
   End
   Begin VB.OptionButton TestOption 
      BackColor       =   &H00C0C0FF&
      Caption         =   "TestOption(0)"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   0
      Top             =   4080
      Width           =   1815
   End
   Begin sbmtc1.MouseTrackControl TestMouseTrack 
      Left            =   120
      Top             =   480
      _ExtentX        =   450
      _ExtentY        =   450
   End
End
Attribute VB_Name = "TestFormB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

TestListbox.Clear
TestListbox.AddItem "TestListbox"

TestMouseTrack.TrackHook TestText1.hWnd
TestMouseTrack.TrackHook TestText2(0).hWnd
TestMouseTrack.TrackHook TestText2(1).hWnd
TestMouseTrack.TrackHook TestListbox.hWnd
TestMouseTrack.TrackHook TestButton.hWnd
TestMouseTrack.TrackHook TestCheck.hWnd
TestMouseTrack.TrackHook TestOption(0).hWnd
TestMouseTrack.TrackHook TestOption(1).hWnd
TestMouseTrack.TrackHook TestCombo.hWnd
TestMouseTrack.TrackHook TestDriveListBox.hWnd
TestMouseTrack.TrackHook TestFileListBox.hWnd

End Sub

Private Sub Form_Unload(Cancel As Integer)

TestMouseTrack.TrackAllUnhook

End Sub

Private Sub AddLogging(LoggingText As String)

If TestLogging.ListCount > 12 Then TestLogging.RemoveItem 0

TestLogging.AddItem LoggingText

End Sub

Private Sub TestMouseTrack_HoverChange(ByVal hWnd As Long, isHovered As Boolean)

Dim BackgroundColor As OLE_COLOR

AddLogging "HoverChange " & hWnd & " " & isHovered

If isHovered = True Then
    BackgroundColor = &HC0FFC0
Else
    BackgroundColor = &HC0C0FF
End If

Select Case hWnd
    Case TestText1.hWnd:        TestText1.BackColor = BackgroundColor
    Case TestText2(0).hWnd:     TestText2(0).BackColor = BackgroundColor
    Case TestText2(1).hWnd:     TestText2(1).BackColor = BackgroundColor
    Case TestListbox.hWnd:      TestListbox.BackColor = BackgroundColor
    Case TestButton.hWnd:       TestButton.BackColor = BackgroundColor
    Case TestCheck.hWnd:        TestCheck.BackColor = BackgroundColor
    Case TestOption(0).hWnd:    TestOption(0).BackColor = BackgroundColor
    Case TestOption(1).hWnd:    TestOption(1).BackColor = BackgroundColor
    Case TestCombo.hWnd:        TestCombo.BackColor = BackgroundColor
    Case TestDriveListBox.hWnd: TestDriveListBox.BackColor = BackgroundColor
    Case TestFileListBox.hWnd:  TestFileListBox.BackColor = BackgroundColor
End Select

End Sub

Private Sub TestMouseTrack_WheelScroll(ByVal hWnd As Long, MouseKeys As Long, MouseRotation As Long, MousePosX As Long, MousePosY As Long)

AddLogging "WheelScroll " & hWnd & " " & MouseKeys & " " & MouseRotation & " " & MousePosX & " " & MousePosY

End Sub

Private Sub TestScrolling_Click()

If TestScrolling.Value = 0 Then
    TestMouseTrack.ChangeWheelScrollingWhenHoveringOnly False
Else
    TestMouseTrack.ChangeWheelScrollingWhenHoveringOnly True
End If

End Sub

VERSION 5.00
Begin VB.UserControl Balloon 
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   885
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Balloon.ctx":0000
   ScaleHeight     =   555
   ScaleWidth      =   885
   ToolboxBitmap   =   "Balloon.ctx":04AC
End
Attribute VB_Name = "Balloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long

Private Type POINT_API
    X As Long
    Y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Enum IconType
    vbExclamation = 48
    vbCritical = 16
    vbInformation = 64
    vbNone = 0
End Enum

Public Sub popUpBalloon(message As String, _
                        Optional Title As String, _
                        Optional Icon As IconType, _
                        Optional showClose As Boolean = True, _
                        Optional autoCloseTime As Integer = 0, _
                        Optional fontSize As Integer = 8, _
                        Optional fontFace As String = "MS Sans Serif", _
                        Optional PutAtCurrentMousePos As Boolean = True, _
                        Optional XinPixels As Integer, _
                        Optional YinPixels As Integer, _
                        Optional CTRL As Control _
                        )
                        

Dim winRect As RECT
Dim Dot As POINT_API
On Error GoTo Default
Unload frmBalloon
If Title = "" Then Title = App.Title
Title = Title & "     "
With frmBalloon
    .Timer1.Interval = autoCloseTime * 1000
    .imgX.Visible = showClose
    .lblMsg = message
    .lblMsg.Font = fontFace
    .lblMsg.fontSize = fontSize
    .lblTitle.Font = fontFace
    .lblTitle.fontSize = fontSize
    .lblTitle = Title
    .imgIcon = .imgIconXP(Icon \ 16)
    .imgIcon.Top = 360 + 60 * Abs(Icon <> 48)
    .imgIcon.Visible = (Icon <> 0)
    .lblTitle.Left = 360 * Abs(.imgIcon.Visible) + 120
    .Width = IIf(.lblMsg.Width > ((360 * Abs(Icon <> 0)) + .lblTitle.Width), .lblMsg.Width, (360 * Abs(Icon <> 0)) + .lblTitle.Width) + 240
    .Height = .lblMsg.Top + .lblMsg.Height + 180
    If PutAtCurrentMousePos Then
        Call GetCursorPos(Dot)
        GoTo SHOWFORM
    ElseIf (XinPixels <> 0) And (YinPixels <> 0) Then
        Dot.X = XinPixels
        Dot.Y = YinPixels
        GoTo SHOWFORM
    Else
        Call GetWindowRect(CTRL.hWnd, winRect)
        Dot.X = winRect.Left * Screen.TwipsPerPixelX + 240
        Dot.Y = winRect.Bottom * Screen.TwipsPerPixelY - (CTRL.Height \ 2)
        Dot.X = Dot.X \ 15
        Dot.Y = Dot.Y \ 15
        GoTo SHOWFORM
    End If
End With

Default:
    Call GetCursorPos(Dot)

SHOWFORM:
    frmBalloon.Left = (Dot.X * 15) - 240
    frmBalloon.Top = Dot.Y * 15
    frmBalloon.Show
End Sub

Private Sub UserControl_Initialize()
    SizeIt
End Sub

Private Sub UserControl_Resize()
    SizeIt
End Sub

Private Sub SizeIt()
    UserControl.Width = 885
    UserControl.Height = 555
End Sub

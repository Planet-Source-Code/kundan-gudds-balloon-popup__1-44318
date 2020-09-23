VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gudds Balloon popUp Demo"
   ClientHeight    =   5400
   ClientLeft      =   2130
   ClientTop       =   1950
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   3600
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3480
      Top             =   3000
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Show close button"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Width           =   2655
      Begin VB.OptionButton Option2 
         Caption         =   "Put with ""Click On Me"" Button"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   720
         Width           =   2055
         Begin VB.TextBox Text3 
            Height          =   285
            Index           =   0
            Left            =   360
            TabIndex        =   21
            Text            =   "34"
            Top             =   0
            Width           =   375
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   16
            Text            =   "34"
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Y"
            Height          =   255
            Index           =   4
            Left            =   840
            TabIndex        =   15
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "X"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   14
            Top             =   0
            Width           =   135
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Put at a custom location"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Put at current mouse position"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Value           =   -1  'True
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3375
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "2"
         Top             =   270
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Test"
         Height          =   375
         Left            =   2400
         TabIndex        =   18
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Text            =   "Arial"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   1320
         List            =   "frmMain.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmMain.frx":002F
         Left            =   1320
         List            =   "frmMain.frx":003F
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Auto close after       seconds"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Font Face"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Font Size"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Icon"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   495
      End
   End
   Begin Project1.Balloon Balloon1 
      Left            =   1200
      Top             =   5520
      _ExtentX        =   1561
      _ExtentY        =   979
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click on Me"
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Customize the bubble"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'======================================================================================
'                              GUDDS BALLOON POP-UP
'======================================================================================
'  How to use:
'   Call the function popUpBallon to show the Balloon popup
'   various parameters are:
'       Message     - Required - The message you want to show
'       Title       - Optional - The title of the balloon; Default is the Project Title
'       Icon        - Optional - The icon you want to show;
'                                Critical, Exclamation or Information; Default is no Icon
'       ShowClose   - Optional - Want to show the close button; Default is True
'       autoCloseTime Optional - Specify nonzero time in seconds if you want to close
'                                the popUp automatically
'       fontSize    - Optional - Specifies the font size; Default is 8
'       fontFace    - Optional - Specifies a custom font; Default is MS Sans Serif
'       PutAtCurrentMousePos   - Optional; Default true; will place the balloon at the
'                                current cursor position
'       XinPixel    - Optional - Specifies the X-coordinate in pixel where ballon is to
'                                be placed;
'                                works only if PutAtCurrentMousePos=False
'       YinPixel    - Optional - Specifies the Y-coordinate in pixel where ballon is to
'                                be placed;
'                                works only if PutAtCurrentMousePos=False
'       CTRL        - Optional - Name of the control (Having hwnd property otherwise current
'                                Cursor position will be used) where you want the balloon to
'                                be placed
'                                works only if PutAtCurrentMousePos=False
'======================================================================================
'                   Email   : imkundan@yahoo.com
'                   Website : http://www20.brinkster.com/GSoftWorks
'                             (Currently under construction :(( visit later
'======================================================================================

Option Explicit

Private Sub Check1_Click()
Me.Height = 1860 + 4000 * Check1.Value
End Sub

Private Sub Command1_Click()
Call Balloon1.popUpBalloon("Hi I am a simple Balloon" & vbCr & _
                            "You can change my font face and font size" & vbCr & _
                            "You can configure me to close automatically after sometime" & vbCr & vbCr & _
                            "You can place me at any place like" & vbCr & _
                            "  1. at current mouse cursor location (Default)" & vbCr & _
                            "  2. at any specified location" & vbCr & _
                            "  3. at any control" & vbCr & vbCr & "Click on Customize the bubble to explore my feathers..:)", _
            "Gudds Balloon Pop-up Example", vbExclamation, , 3, , , False, , , Command1)
End Sub

Private Sub Command2_Click()
Call Balloon1.popUpBalloon("See my Icon..." & vbCr & _
                            "I'll close automatically after " & Check2 * Val(Text2) & " seconds" & vbCr & _
                            "I'm" & IIf(Check3.Value, "", " not ") & " having the close button, see my top right corner.." & vbCr & _
                            "My Font size is " & Combo2.List(Combo2.ListIndex) & vbCr & _
                            "My Font face is " & Text1 & vbCr & _
                            "I'm located at " & IIf(Option2(0).Value, "Current Cursor Position", IIf(Option2(1).Value, "(" & Text3(0) & "," & Text3(1) & ")", """" & "Click On Me Button" & """")), _
                            "I'm a customized Balloon", _
                            getIcon, Check3.Value, _
                            Check2 * Val(Text2), _
                            Combo2.List(Combo2.ListIndex), _
                            Text1, _
                            Option2(0).Value, _
                            IIf(Option2(1).Value, 1, 0) * Val(Text3(0)), _
                            IIf(Option2(1).Value, 1, 0) * Val(Text3(1)), _
                            IIf(Option2(2).Value, Command1, Command2) _
                            )
End Sub

Private Sub Form_Load()
Me.Hide
    Check1_Click
    Combo1.ListIndex = 1
    Combo2.ListIndex = 1
    Command1_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Balloon1.popUpBalloon("Please vote if you like this control..." & vbCr & "Send any comment to imkundan@yahoo.com" & vbCr & vbCr & "Thank you for using Gudds Balloon Popup", "Gudds Balloon Popup", vbExclamation, True, 4)
End Sub

Private Function getIcon() As IconType
Select Case Combo1.ListIndex
    Case 0: getIcon = vbNone
    Case 1: getIcon = vbCritical
    Case 2: getIcon = vbExclamation
    Case 3: getIcon = vbInformation
End Select
End Function

Private Sub Timer1_Timer()
    Me.Show
    Timer1.Enabled = False
End Sub

VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMenu 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Menu"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1620
   FillColor       =   &H00C0FFFF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   1620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer FocusChecker 
      Interval        =   50
      Left            =   1320
      Top             =   840
   End
   Begin MSForms.Label lblExit 
      Height          =   255
      Left            =   140
      TabIndex        =   0
      Top             =   1200
      Width           =   800
      BackColor       =   -2147483624
      VariousPropertyBits=   19
      Size            =   "1411;450"
      MousePointer    =   99
      MouseIcon       =   "frmMenu.frx":0000
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblRefresh 
      Height          =   420
      Left            =   140
      TabIndex        =   1
      Top             =   910
      Width           =   1360
      BackColor       =   -2147483624
      VariousPropertyBits=   19
      Size            =   "2399;741"
      MousePointer    =   99
      MouseIcon       =   "frmMenu.frx":0162
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblCP 
      Height          =   420
      Left            =   140
      TabIndex        =   2
      Top             =   600
      Width           =   1360
      BackColor       =   -2147483624
      VariousPropertyBits=   19
      Size            =   "2399;741"
      MousePointer    =   99
      MouseIcon       =   "frmMenu.frx":02C4
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblActive 
      Height          =   420
      Left            =   140
      TabIndex        =   3
      Top             =   120
      Width           =   1360
      BackColor       =   -2147483624
      VariousPropertyBits=   19
      Size            =   "2399;741"
      MousePointer    =   99
      MouseIcon       =   "frmMenu.frx":0426
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   1600
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3240
   End
   Begin VB.Line Line3 
      X1              =   1600
      X2              =   1600
      Y1              =   0
      Y2              =   3240
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   1600
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000013&
      X1              =   120
      X2              =   1550
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Label lblHighlight 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   80
      TabIndex        =   4
      Top             =   960
      Width           =   1400
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function GetActiveWindow Lib "user32" () As Integer
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Const SPI_GETWORKAREA As Long = 48&
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
    (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
'/* Operating system version information
Private Type OSVersionInfo
    OSVSize       As Long
    dwVerMajor    As Long
    dwVerMinor    As Long
    dwBuildNumber As Long
    PlatformID    As Long
    szCSDVersion  As String * 128
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Dim MyWindow As Long
Dim lngItemINdex As Long

Private Sub Form_Activate()
    MyWindow = GetActiveWindow
End Sub

Private Sub Form_Load()
    Me.Hide
    
    'Load Language
    lblExit.Caption = LANG_STRING_19 'Thoat
    lblRefresh.Caption = LANG_STRING_23 'Refresh
    lblCP.Caption = LANG_STRING_22 'Bang dieu khien
    If blnEnable = False Then 'Tat
       lblActive.Caption = LANG_STRING_20 'Bat
       lblActive.ForeColor = &HC00000
    Else
       lblActive.Caption = LANG_STRING_21 'Tat
       lblActive.ForeColor = &HFF&
    End If
    
    'Adjust menu items width in form
    Dim lngMax1 As Long
    Dim lngMax2 As Long
    Dim lngMax As Long
    Dim lngWidth1 As Long: lngWidth1 = lblExit.Width
    Dim lngWidth2 As Long: lngWidth2 = lblRefresh.Width
    Dim lngWidth3 As Long: lngWidth3 = lblCP.Width
    Dim lngWidth4 As Long: lngWidth4 = lblActive.Width
    
    lblHighlight.Top = -600 'Hide Highlight
    If lngWidth1 >= lngWidth2 Then
        lngMax1 = lngWidth1
    Else
        lngMax1 = lngWidth2
    End If
    
    If lngWidth3 >= lngWidth4 Then
        lngMax2 = lngWidth3
    Else
        lngMax2 = lngWidth4
    End If
    
    If lngMax1 >= lngMax2 Then
        lngMax = lngMax1
    Else
        lngMax = lngMax2
    End If
    lblExit.Width = lngMax + 230: lblRefresh.Width = lngMax + 230: lblCP.Width = lngMax + 230: lblActive.Width = lngMax + 230
    Me.Width = lngMax + 240
    lblHighlight.Width = lngMax + 100 'Resize Highlight item
    Line5.X2 = lngMax + 50
    Line3.X1 = lngMax + 230: Line3.X2 = lngMax + 230
    Line1.X2 = lngMax + 240
    Line4.X2 = lngMax + 240
    
    'Reposition form to match mouse cursor
    Dim typPA As POINTAPI
    Dim typRect As RECT
    
    'Determine current window dimensions
    GetWindowRect Me.hwnd, typRect
    
    'Determine mouse cursor position
    GetCursorPos typPA
    
    Dim Rc         As RECT
    Dim scrnRight  As Long
    'Dim scrnBottom     As Long         '/* Height of the screen - taskbar (if it is on the bottom)
    Dim frmRight As Long
    
    Dim OSV        As OSVersionInfo
    
    '/* Get OS compatability flag
    OSV.OSVSize = Len(OSV)
    
    '/* Get Screen and TaskBar size
    Call SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
    
    '/* Screen Height - Taskbar Height (if is is located at the bottom of the screen)
    'scrnBottom = Rc.Bottom * Screen.TwipsPerPixelY
    
    '/* Is the taskbar is located on the right side of the screen? (scrnRight < Screen.width)
    scrnRight = (Rc.Right * Screen.TwipsPerPixelX)
    
    '/* Locate Form to bottom right and set default size
    Top = typPA.Y * Screen.TwipsPerPixelY - frmMenu.Height
    
    frmRight = typPA.X * Screen.TwipsPerPixelX + Width
    If frmRight >= scrnRight Then
        Left = scrnRight - Width - 100
    Else
        Left = typPA.X * Screen.TwipsPerPixelX
    End If
    
    'Always show menu ontop
    Me.Show
    SetOnTop Me.hwnd, True

End Sub

Private Sub Form_LostFocus()
    Unload Me 'Unload menu if it lost focus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHighlight.Top = -600 'disable highlight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetOnTop Me.hwnd, False 'Restore OnTop status
End Sub

Private Sub lblExit_Click()
   'Release all
   Me.Hide
   Terminate
End Sub


Private Sub lblRefresh_Click()
    'Reload data list
    Me.Hide
    m_frmSysTray.tmrShowTooltip.Enabled = False
    m_frmSysTray.LoadData
    m_frmSysTray.tmrShowTooltip.Enabled = True
    Unload Me
End Sub

Private Sub lblCP_Click()
    'Show frmEdit
    Me.Hide
    frmEdit.Show
    m_frmSysTray.tmrShowTooltip.Enabled = False
    frmEdit.ZOrder
    Unload Me
End Sub

Private Sub lblActive_Click()
   Me.Hide
   blnEnable = Not blnEnable
   Unload Me
End Sub

Private Sub lblCP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHighlight.Top = lblCP.Top - 80
End Sub

Private Sub lblActive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHighlight.Top = lblActive.Top - 80
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHighlight.Top = lblExit.Top - 80
End Sub

Private Sub lblRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHighlight.Top = lblRefresh.Top - 80
End Sub



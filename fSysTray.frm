VERSION 5.00
Begin VB.Form frmSysTray 
   BorderStyle     =   0  'None
   Caption         =   "Language Tooltip"
   ClientHeight    =   1920
   ClientLeft      =   5595
   ClientTop       =   3045
   ClientWidth     =   4680
   Icon            =   "fSysTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrShowTooltip 
      Interval        =   1000
      Left            =   960
      Top             =   720
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   1680
      Picture         =   "fSysTray.frx":058A
      Top             =   480
      Width           =   240
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
   
Private Declare Function Shell_NotifyIconA Lib "shell32.dll" _
   (ByVal dwMessage As Long, lpData As NOTIFYICONDATAA) As Long
   
Private Declare Function Shell_NotifyIconW Lib "shell32.dll" _
   (ByVal dwMessage As Long, lpData As NOTIFYICONDATAW) As Long

Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4

Private Const NOTIFYICON_VERSION = 3

Private Type NOTIFYICONDATAA
   cbSize As Long             ' 4
   hwnd As Long               ' 8
   uID As Long                ' 12
   uFlags As Long             ' 16
   uCallbackMessage As Long   ' 20
   hIcon As Long              ' 24
   szTip As String * 128      ' 152
   dwState As Long            ' 156
   dwStateMask As Long        ' 160
   szInfo As String * 256     ' 416
   uTimeOutOrVersion As Long  ' 420
   szInfoTitle As String * 64 ' 484
   dwInfoFlags As Long        ' 488
   guidItem As Long           ' 492
End Type

Private Type NOTIFYICONDATAW
   cbSize As Long             ' 4
   hwnd As Long               ' 8
   uID As Long                ' 12
   uFlags As Long             ' 16
   uCallbackMessage As Long   ' 20
   hIcon As Long              ' 24
   szTip(0 To 255) As Byte    ' 280
   dwState As Long            ' 284
   dwStateMask As Long        ' 288
   szInfo(0 To 511) As Byte   ' 800
   uTimeOutOrVersion As Long  ' 804
   szInfoTitle(0 To 127) As Byte ' 932
   dwInfoFlags As Long        ' 936
   guidItem As Long           ' 940
End Type


Private nfIconDataA As NOTIFYICONDATAA
Private nfIconDataW As NOTIFYICONDATAW

Private Const NOTIFYICONDATAA_V1_SIZE_A = 88
Private Const NOTIFYICONDATAA_V1_SIZE_U = 152
Private Const NOTIFYICONDATAA_V2_SIZE_A = 488
Private Const NOTIFYICONDATAA_V2_SIZE_U = 936

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Const WM_USER = &H400

Private Const NIN_SELECT = WM_USER
Private Const NINF_KEY = &H1
Private Const NIN_KEYSELECT = (NIN_SELECT Or NINF_KEY)
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

' Version detection:
Private Declare Function GetVersion Lib "Kernel32" () As Long
Public Event SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
Public Event SysTrayMouseUp(ByVal eButton As MouseButtonConstants)
Public Event SysTrayMouseMove()
Public Event SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
Public Event MenuClick(ByVal lIndex As Long, ByVal sKey As String)
Public Event BalloonShow()
Public Event BalloonHide()
Public Event BalloonTimeOut()
Public Event BalloonClicked()

Public Enum EBalloonIconTypes
   NIIF_NONE = 0
   NIIF_INFO = 1
   NIIF_WARNING = 2
   NIIF_ERROR = 3
   NIIF_NOSOUND = &H10
End Enum

Private m_iCurIndex As Long 'Counter variable for Timer
Private m_bUseUnicode As Boolean
Private m_bSupportsNewVersion As Boolean

Private arrVoc() As String 'Dump all data to memory :)

'Show balloon tip
Public Sub ShowBalloonTip( _
      ByVal SMessage As String, _
      Optional ByVal sTitle As String, _
      Optional ByVal eIcon As EBalloonIconTypes, _
      Optional ByVal lTimeOutMs _
   )
Dim lR As Long
   If (m_bSupportsNewVersion) Then
      If (m_bUseUnicode) Then
         stringToArray SMessage, nfIconDataW.szInfo, 512
         stringToArray sTitle, nfIconDataW.szInfoTitle, 128
         nfIconDataW.uTimeOutOrVersion = lTimeOutMs
         nfIconDataW.dwInfoFlags = eIcon
         nfIconDataW.uFlags = NIF_INFO
         lR = Shell_NotifyIconW(NIM_MODIFY, nfIconDataW)
      Else
         nfIconDataA.szInfo = SMessage
         nfIconDataA.szInfoTitle = sTitle
         nfIconDataA.uTimeOutOrVersion = lTimeOutMs
         nfIconDataA.dwInfoFlags = eIcon
         nfIconDataA.uFlags = NIF_INFO
         lR = Shell_NotifyIconA(NIM_MODIFY, nfIconDataA)
      End If
   Else
      ' can't do it, fail silently.
   End If
End Sub

'Get ToolTip content
Public Property Get ToolTip() As String
Dim sTip As String
Dim iPos As Long
    sTip = nfIconDataA.szTip
    iPos = InStr(sTip, Chr$(0))
    If (iPos <> 0) Then
        sTip = Left$(sTip, iPos - 1)
    End If
    ToolTip = sTip
End Property

'Set Tooltip content
Public Property Let ToolTip(ByVal sTip As String)
   If (m_bUseUnicode) Then
      stringToArray sTip, nfIconDataW.szTip, unicodeSize(IIf(m_bSupportsNewVersion, 128, 64))
      nfIconDataW.uFlags = NIF_TIP
      Shell_NotifyIconW NIM_MODIFY, nfIconDataW
   Else
      If (sTip & Chr$(0) <> nfIconDataA.szTip) Then
         nfIconDataA.szTip = sTip & Chr$(0)
         nfIconDataA.uFlags = NIF_TIP
         Shell_NotifyIconA NIM_MODIFY, nfIconDataA
      End If
   End If
End Property
'Get Icon handle
Public Property Get IconHandle() As Long
    IconHandle = nfIconDataA.hIcon
End Property

'Set Icon handle
Public Property Let IconHandle(ByVal hIcon As Long)
   If (m_bUseUnicode) Then
      If (hIcon <> nfIconDataW.hIcon) Then
         nfIconDataW.hIcon = hIcon
         nfIconDataW.uFlags = NIF_ICON
         Shell_NotifyIconW NIM_MODIFY, nfIconDataW
      End If
   Else
      If (hIcon <> nfIconDataA.hIcon) Then
         nfIconDataA.hIcon = hIcon
         nfIconDataA.uFlags = NIF_ICON
         Shell_NotifyIconA NIM_MODIFY, nfIconDataA
      End If
   End If
End Property

Private Sub Form_Load()
   ' Get version:
   Dim lMajor As Long
   Dim lMinor As Long
   Dim bIsNt As Long
   GetWindowsVersion lMajor, lMinor, , , bIsNt
   
   ' Remove EnableBalloonTips in Registry (avoid case of setting value as FALSEb)
   DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "EnableBalloonTips"
   
   If (bIsNt) Then
      m_bUseUnicode = True
      If (lMajor >= 5) Then
         ' 2000 or XP
         m_bSupportsNewVersion = True
      End If
   ElseIf (lMajor = 4) And (lMinor = 90) Then
      ' Windows ME
      m_bSupportsNewVersion = True
   End If
   
   
   'Add the icon to the system tray...
   Dim lR As Long
   
   If (m_bUseUnicode) Then
      With nfIconDataW
         .hwnd = Me.hwnd
         .uID = Me.Icon
         .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
         .uCallbackMessage = WM_MOUSEMOVE
         .hIcon = Me.Icon.Handle
         stringToArray App.FileDescription, .szTip, unicodeSize(IIf(m_bSupportsNewVersion, 128, 64))
         If (m_bSupportsNewVersion) Then
            .uTimeOutOrVersion = NOTIFYICON_VERSION
         End If
         .cbSize = nfStructureSize
      End With
      lR = Shell_NotifyIconW(NIM_ADD, nfIconDataW)
      If (m_bSupportsNewVersion) Then
         Shell_NotifyIconW NIM_SETVERSION, nfIconDataW
      End If
   Else
      With nfIconDataA
         .hwnd = Me.hwnd
         .uID = Me.Icon
         .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
         .uCallbackMessage = WM_MOUSEMOVE
         .hIcon = Me.Icon.Handle
         .szTip = App.FileDescription & Chr$(0)
         If (m_bSupportsNewVersion) Then
            .uTimeOutOrVersion = NOTIFYICON_VERSION
         End If
         .cbSize = nfStructureSize
      End With
      lR = Shell_NotifyIconA(NIM_ADD, nfIconDataA)
      If (m_bSupportsNewVersion) Then
         lR = Shell_NotifyIconA(NIM_SETVERSION, nfIconDataA)
      End If
   End If
   IconHandle = imgIcon.Picture.Handle
   ToolTip = LANG_STRING_00 'Phan mem hoc tu vung

   'Load database
   LoadData
End Sub

'Dump all data to memory
Public Sub LoadData()
    Dim fNum As Long, B() As Byte, fp
    fp = App.Path & "\data\data.txt"
    fNum = FreeFile()
    Open fp For Binary Access Read As #fNum
        ReDim B(LOF(fNum))
    Get #fNum, , B
    Close #fNum
    B = Trim$(B)
    Dim i As Integer, s As String
    arrVoc = Split(B, vbCrLf)
    m_iCurIndex = -1
End Sub

Private Sub stringToArray( _
      ByVal sString As String, _
      bArray() As Byte, _
      ByVal lMaxSize As Long _
   )
Dim B() As Byte
Dim i As Long
Dim j As Long
   If Len(sString) > 0 Then
      B = sString
      For i = LBound(B) To UBound(B)
         bArray(i) = B(i)
         If (i = (lMaxSize - 2)) Then
            Exit For
         End If
      Next i
      For j = i To lMaxSize - 1
         bArray(j) = 0
      Next j
   End If

End Sub

'Get size of content
Private Function unicodeSize(ByVal lSize As Long) As Long
   If (m_bUseUnicode) Then
      unicodeSize = lSize * 2
   Else
      unicodeSize = lSize
   End If
End Function

Private Property Get nfStructureSize() As Long
   If (m_bSupportsNewVersion) Then
      If (m_bUseUnicode) Then
         nfStructureSize = NOTIFYICONDATAA_V2_SIZE_U
      Else
         nfStructureSize = NOTIFYICONDATAA_V2_SIZE_A
      End If
   Else
      If (m_bUseUnicode) Then
         nfStructureSize = NOTIFYICONDATAA_V1_SIZE_U
      Else
         nfStructureSize = NOTIFYICONDATAA_V1_SIZE_A
      End If
   End If
End Property

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lX As Long
   ' VB manipulates the x value according to scale mode:
   ' we must remove this before we can interpret the
   ' message windows was trying to send to us:
   lX = ScaleX(X, Me.ScaleMode, vbPixels)
   Select Case lX
   Case WM_MOUSEMOVE
      RaiseEvent SysTrayMouseMove
   Case WM_LBUTTONUP
      RaiseEvent SysTrayMouseDown(vbLeftButton)
   Case WM_LBUTTONUP
      RaiseEvent SysTrayMouseUp(vbLeftButton)
   Case WM_LBUTTONDBLCLK
        frmEdit.Show 'Double click to show frmEdit
        tmrShowTooltip.Enabled = False
        frmEdit.ZOrder
   Case WM_RBUTTONDOWN
       'ShowMenu
       Load frmMenu
   Case WM_RBUTTONUP
      RaiseEvent SysTrayMouseUp(vbRightButton)
   Case WM_RBUTTONDBLCLK
      RaiseEvent SysTrayDoubleClick(vbRightButton)
   Case NIN_BALLOONSHOW
      RaiseEvent BalloonShow
   Case NIN_BALLOONHIDE
      RaiseEvent BalloonHide
   Case NIN_BALLOONTIMEOUT
      RaiseEvent BalloonTimeOut
   Case NIN_BALLOONUSERCLICK
      RaiseEvent BalloonClicked
   End Select

End Sub

Private Sub GetWindowsVersion( _
      Optional ByRef lMajor = 0, _
      Optional ByRef lMinor = 0, _
      Optional ByRef lRevision = 0, _
      Optional ByRef lBuildNumber = 0, _
      Optional ByRef bIsNt = False _
   )
Dim lR As Long
   lR = GetVersion()
   lBuildNumber = (lR And &H7F000000) \ &H1000000
   If (lR And &H80000000) Then lBuildNumber = lBuildNumber Or &H80
   lRevision = (lR And &HFF0000) \ &H10000
   lMinor = (lR And &HFF00&) \ &H100
   lMajor = (lR And &HFF)
   bIsNt = ((lR And &H80000000) = 0)
End Sub
Private Sub tmrShowTooltip_Timer()
On Error GoTo 1
If blnEnable = False Or UBound(arrVoc) = 0 Then Exit Sub
If lngTimeShow_Count <> lngTimeShow Then
   lngTimeShow_Count = lngTimeShow_Count + 1
Else
    
    Dim arrWord() As String 'Split text line to arrWord
    If bQueue = False Then
        Dim lngRandNum As Long
        lngRandNum = RandomNumber(0, UBound(arrVoc) - 1)
        arrWord = Split(arrVoc(lngRandNum), "##") 'The Delimiter = ##
    Else
        m_iCurIndex = m_iCurIndex + 1
        If m_iCurIndex = UBound(arrVoc) Then
            m_iCurIndex = 0
        End If
        arrWord = Split(arrVoc(m_iCurIndex), "##") 'The Delimiter = ##
    End If
    
    'Show balloontip
    ShowBalloonTip Trim$(arrWord(2)), Trim$(arrWord(1)), NIIF_INFO Or NIIF_NOSOUND, lngTimeDelay

    lngTimeShow_Count = 0
    
    Erase arrWord()
End If
Exit Sub
1:
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If (m_bUseUnicode) Then
      Shell_NotifyIconW NIM_DELETE, nfIconDataW
   Else
      Shell_NotifyIconA NIM_DELETE, nfIconDataA
   End If
   'Release memory
   Erase arrVoc()
End Sub


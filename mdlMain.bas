Attribute VB_Name = "mdlMain"
Public m_frmSysTray As frmSysTray
Public Declare Function getprivateprofilestring Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function writeprivateprofilestring Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public lngTimeDelay As Long
Public bQueue As Boolean
Public lngTimeShow As Long
Public lngTimeShow_Count As Long
Public lngTimeDelay_Count  As Long
Public blnEnable As Boolean
Public strLangFile As String

Public Const MB_ICONINFORMATION As Long = &H40&
Public Const MB_ICONEXCLAMATION = &H30&
Public Const MB_ICONQUESTION = &H20&
Public Const MB_OK = &H0&
Public Const MB_OKCANCEL = &H1&
Public Const MB_YESNO = &H4&
Public Const MB_TASKMODAL As Long = &H2000&
'/* Set window in the Z order
Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
     ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Const ERROR_SUCCESS = 0&
    Const REG_SZ = 1 ' Unicode nul terminated String
    Const REG_DWORD = 4 ' 32-bit number
Public Enum HKeyTypes
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum
Public Sub SaveString(hKey As HKeyTypes, strPath As String, strValue As String, strdata As String)
    'EXAMPLE:
    '
    'Call savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub


Public Function DeleteValue(ByVal hKey As HKeyTypes, ByVal strPath As String, ByVal strValue As String)
    'EXAMPLE:
    '
    'Call DeleteValue(HKEY_CURRENT_USER, "So
    '     ftware\VBW\Registry", "Dword")
    '
    Dim keyhand As Long
    Dim r As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)

End Function


Public Function DeleteKey(ByVal hKey As HKeyTypes, ByVal strPath As String)
    'EXAMPLE:
    '
    'Call DeleteKey(HKEY_CURRENT_USER, "Soft
    '     ware\VBW\Registry")
    '
    Dim keyhand As Long
    Dim r As Long
    r = RegDeleteKey(hKey, strPath)
End Function
Public Sub AddToRun(ProgramName As String, FileToRun As String)
    'Add a program to the 'Run at Startup' r
    '     egistry keys
    Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName, FileToRun)
End Sub


Public Sub RemoveFromRun(ProgramName As String)
    'Remove a program from the 'Run at Start
    '     up' registry keys
    Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName)
End Sub


Function ReadINI(Section As String, KeyName As String, FileName As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, getprivateprofilestring(Section, ByVal KeyName$, "", sRet, Len(sRet), FileName))
End Function

Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFilename) As Integer
    Dim r
    r = writeprivateprofilestring(sSection, sKeyName, sNewString, sFilename)
End Function
Public Function FileExists(FullFileName As String) As Boolean
    On Error GoTo MakeF
    Dim lngSize As Long
    lngSize = -1
    lngSize = FileLen(FullFileName)
    If lngSize >= 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
    Exit Function
MakeF:
    FileExists = False
End Function

Sub Main()
   If App.PrevInstance = True Then End
    'Load ini file into "Current" frame
    Dim iniFile As String, ItsThere As Boolean
    iniFile = App.Path & "\data\settings.ini"
    ItsThere = FileExists(iniFile)
    If ItsThere = False Then
        Open iniFile For Output As #1
        Print #1, "[EngTip]"
        Print #1, "AutoStart = 0"
        Print #1, "Time1 = 0"
        Print #1, "Time2 = 0"
        Print #1, "Style =True"
        Print #1, "Hidepix=False"
        Print #1, "Langfile=english"
        Close #1
    Else
        Select Case Val(ReadINI("EngTip", "Time1", iniFile))
        Case 0
            lngTimeShow = 10
        Case 1
            lngTimeShow = 30
        Case 2
            lngTimeShow = 60
        Case 3
            lngTimeShow = 180
        Case 4
            lngTimeShow = 360
        Case 5
            lngTimeShow = 720
        End Select
    
        Select Case Val(ReadINI("EngTip", "Time2", iniFile))
        Case 0
            lngTimeDelay = 5000
        Case 1
            lngTimeDelay = 10000
        Case 2
            lngTimeDelay = 20000
        Case 3
            lngTimeDelay = 30000
        Case 4
            lngTimeDelay = 3600000
        End Select
    End If
    If ReadINI("EngTip", "AutoStart", iniFile) = "1" Then
        Call AddToRun("EngTip", App.Path & "\data\" & App.EXEName & ".exe")
    End If
    
    If ReadINI("EngTip", "Style", iniFile) = "True" Then
        bQueue = True
    Else
        bQueue = False
    End If
    
    If ReadINI("EngTip", "Hidepix", iniFile) = "1" Then
        bHidePix = False
    Else
        bHidePix = True
    End If
    strLangFile = ReadINI("EngTip", "langFile", iniFile)
    'Load language
    ChangeLang
    'Show Tray Icon
    Set m_frmSysTray = New frmSysTray
    With m_frmSysTray
       'Load frmMain
        Load m_frmSysTray
    End With
    blnEnable = True
End Sub
Function RandomNumber(Lowerbound As Long, Upperbound As Long)
    Randomize
    RandomNumber = Int((Upperbound - Lowerbound) * Rnd + Lowerbound)
End Function

Public Sub SetOnTop(lHwnd As Long, Optional ByVal bSetOnTop As Boolean = True)
  '/* The SetWindowPos function changes the size, position, and Z order of a child,
  '/* pop-up, or top-level window. Child, pop-up, and top-level windows are ordered
  '/* according to their appearance on the screen. The topmost window receives the
  '/* highest rank and is the first window in the Z order.
  Const Flags As Long = &H273
  '/* SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE or SWP_NOOWNERZORDER Or SWP_FRAMECHANGED
    
    If bSetOnTop Then
        Call SetWindowPos(lHwnd, -1, 0, 0, 0, 0, Flags)
    Else
        Call SetWindowPos(lHwnd, -2, 0, 0, 0, 0, Flags)
    End If
    
End Sub
Sub Terminate()
'Clear all string
        LANG_STRING_00 = vbNullString
        LANG_STRING_01 = vbNullString
        LANG_STRING_02 = vbNullString
        LANG_STRING_03 = vbNullString
        LANG_STRING_04 = vbNullString
        LANG_STRING_05 = vbNullString
        LANG_STRING_06 = vbNullString
        LANG_STRING_07 = vbNullString
        LANG_STRING_08 = vbNullString
        LANG_STRING_09 = vbNullString
        LANG_STRING_10 = vbNullString
        LANG_STRING_11 = vbNullString
        LANG_STRING_12 = vbNullString
        LANG_STRING_13 = vbNullString
        LANG_STRING_14 = vbNullString
        LANG_STRING_15 = vbNullString
        LANG_STRING_16 = vbNullString
        LANG_STRING_17 = vbNullString
        LANG_STRING_18 = vbNullString
        LANG_STRING_19 = vbNullString
        LANG_STRING_20 = vbNullString
        LANG_STRING_21 = vbNullString
        LANG_STRING_22 = vbNullString
        LANG_STRING_23 = vbNullString
        LANG_STRING_24 = vbNullString
        LANG_STRING_25 = vbNullString
        LANG_STRING_26 = vbNullString
        LANG_STRING_27 = vbNullString
        LANG_STRING_28 = vbNullString
        LANG_STRING_29 = vbNullString
    sCodePage = 0
    cnvUni2 = vbNullString
    cnvUni = vbNullString
    'Release memory
    Unload frmEdit
    Set frmEdit = Nothing
    Unload m_frmSysTray
    Set m_frmSysTray = Nothing
    End
End Sub

VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmEdit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Setting"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5655
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkStartup 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   1920
      Width           =   255
   End
   Begin VB.OptionButton optQueue 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   200
   End
   Begin VB.OptionButton optRandom 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   360
      Picture         =   "frmEdit.frx":0000
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   5325
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2010 Ngoc Phu - Dieu Ai"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Image imgOpenfile 
      Height          =   1425
      Left            =   4440
      MouseIcon       =   "frmEdit.frx":E1FA
      MousePointer    =   99  'Custom
      Picture         =   "frmEdit.frx":E504
      Top             =   480
      Width           =   1140
   End
   Begin MSForms.ComboBox cboTime 
      Height          =   345
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   1800
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      DisplayStyle    =   7
      Size            =   "3175;609"
      ListRows        =   10
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblQueue 
      Height          =   225
      Left            =   2640
      TabIndex        =   12
      Top             =   120
      Width           =   1245
      BackColor       =   16777215
      VariousPropertyBits=   268435483
      Caption         =   "lblQueue"
      Size            =   "2196;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblRandom 
      Height          =   225
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   1410
      BackColor       =   16777215
      VariousPropertyBits=   268435483
      Caption         =   "lblRandom"
      Size            =   "2487;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblTime 
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   555
      Width           =   2100
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblTime"
      Size            =   "3704;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblStyleshow 
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2100
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblStyleshow"
      Size            =   "3704;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblAutostart 
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   2100
      BackColor       =   16777215
      VariousPropertyBits=   268435483
      Caption         =   "lblAutoStart"
      Size            =   "3704;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ComboBox cboTimeDelay 
      Height          =   345
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   1800
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      DisplayStyle    =   7
      Size            =   "3175;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblDelay 
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   1035
      Width           =   2100
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblDelay"
      Size            =   "3704;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ComboBox cboLang 
      Height          =   345
      Left            =   2400
      TabIndex        =   4
      Top             =   1440
      Width           =   1800
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      DisplayStyle    =   7
      Size            =   "3175;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblLanguage 
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   1500
      Width           =   2100
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblLanguage"
      Size            =   "3704;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
    "GetOpenFileNameA" (pOpenfilename As OpenFilename) As Long
'Dim MyWindow As Long
Private Type OpenFilename
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    iFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Enum OFNFlagsEnum
    OFN_ALLOWMULTISELECT = &H200
    OFN_CREATEPROMPT = &H2000
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_EXPLORER = &H80000
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_FILEMUSTEXIST = &H1000
    OFN_HIDEREADONLY = &H4
    OFN_LONGNAMES = &H200000
    OFN_NOCHANGEDIR = &H8
    OFN_NODEREFERENCELINKS = &H100000
    OFN_NOLONGNAMES = &H40000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NOVALIDATE = &H100
    OFN_OVERWRITEPROMPT = &H2
    OFN_PATHMUSTEXIST = &H800
    OFN_READONLY = &H1
    OFN_SHAREAWARE = &H4000
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHARENOWARN = 1
    OFN_SHAREWARN = 0
    OFN_SHOWHELP = &H10
End Enum

Public Sub LoadLangFilesToCombo()
Dim sFile As String
Dim lElement As Long
Dim sAns() As String
ReDim sAns(0) As String

sFile = Dir(App.Path & "\lang\*.txt", vbNormal + vbHidden + vbReadOnly + _
   vbSystem + vbArchive)
If sFile <> "" Then
sFile = Left$(sFile, Len(sFile) - 4)
cboLang.AddItem sFile
    Do
        sFile = Dir
        If sFile = "" Then Exit Do
        sFile = Left$(sFile, Len(sFile) - 4)
        cboLang.AddItem sFile
    Loop
End If

End Sub
Private Sub ChangInterfaceLang()
    'Load language string
    ChangeLang
    
    lblTime.Caption = LANG_STRING_12  'Khoang thoi gian hien thi giua 2 tu
    lblStyleshow.Caption = LANG_STRING_14  'Hien thi tu vung theo
    lblAutostart.Caption = LANG_STRING_11  'Cho phep khoi dong cung windows
    lblDelay.Caption = LANG_STRING_13 'Moi tu se hien thi trong khoang thoi gian
    lblQueue.Caption = LANG_STRING_15  'Tuan tu
    lblRandom.Caption = LANG_STRING_16  'Ngau nhien
    lblLanguage.Caption = LANG_STRING_18 'Ngon ngu hien thi
    cboTime.Clear
    cboTime.AddItem "10 " & LANG_STRING_24 'giay
    cboTime.AddItem "30 " & LANG_STRING_24 'giay
    cboTime.AddItem "01 " & LANG_STRING_25 'phut
    cboTime.AddItem "10 " & LANG_STRING_25 'phut
    cboTime.AddItem "30 " & LANG_STRING_25 'phut
    cboTime.AddItem "01 " & LANG_STRING_26 'gio
    cboTime.AddItem "02 " & LANG_STRING_26 'gio
    cboTimeDelay.Clear
    cboTimeDelay.AddItem "05 " & LANG_STRING_24 'giay
    cboTimeDelay.AddItem "10 " & LANG_STRING_24 'giay
    cboTimeDelay.AddItem "20 " & LANG_STRING_24 'giay
    cboTimeDelay.AddItem "30 " & LANG_STRING_24 'giay

    
End Sub
Private Sub Form_Load()

    'Change interface language
    ChangInterfaceLang

    'Load setting values
    Dim iniFile As String
    iniFile = App.Path & "\data\settings.ini"

    chkStartup.Value = Val(ReadINI("EngTip", "AutoStart", iniFile))
    cboTime.ListIndex = Val(ReadINI("EngTip", "Time1", iniFile))
    cboTimeDelay.ListIndex = Val(ReadINI("EngTip", "Time2", iniFile))
    If ReadINI("EngTip", "Style", iniFile) = "True" Then
        optQueue.Value = True
        optRandom.Value = False
    Else
        optQueue.Value = False
        optRandom.Value = True
    End If
    
    'Load list of languages
    LoadLangFilesToCombo
    'Select the current language
    Dim i As Long
    For i = 0 To cboLang.ListCount - 1
        If cboLang.List(i) = strLangFile Then
            cboLang.ListIndex = i
            Exit For
        End If
    Next

End Sub

'Save all setting values
Private Sub DoSave()
 '
    Dim iniFile As String
    If strLangFile <> cboLang.List(cboLang.ListIndex) Then
        strLangFile = cboLang.List(cboLang.ListIndex)
    End If
    
    iniFile = App.Path & "\data\settings.ini"
    WriteINI "EngTip", "AutoStart", chkStartup.Value, iniFile
    WriteINI "EngTip", "Time1", cboTime.ListIndex, iniFile
    WriteINI "EngTip", "Time2", cboTimeDelay.ListIndex, iniFile
    WriteINI "EngTip", "Style", optQueue.Value, iniFile
    WriteINI "EngTip", "Langfile", strLangFile, iniFile
    
    'Show after
    Select Case cboTime.ListIndex
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
    
    'Hide after
    Select Case cboTimeDelay.ListIndex
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
    Me.Hide
    'ClearList
    m_frmSysTray.tmrShowTooltip.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DoSave 'Save before terminated
End Sub

Private Sub imgOpenfile_Click()
    'Open data.txt by notepad app
    Shell "Notepad.exe """ & App.Path & "/data/data.txt""", vbNormalFocus
End Sub

Private Sub lblRandom_Click()
    optQueue.Value = False
    optRandom.Value = True
End Sub

Private Sub lblQueue_Click()
    optQueue.Value = True
    optRandom.Value = False
End Sub

Private Sub optQueue_Click()
    optRandom.Value = Not optQueue.Value
    If optQueue.Value = True Then
        bQueue = True
    End If
End Sub

Private Sub optRandom_Click()
    optQueue.Value = Not optRandom.Value
    If optRandom.Value = True Then
        bQueue = False
    End If
End Sub

Private Sub chkStartup_Click()
    'Write to registry
    If chkStartup.Value = 1 Then
       Call AddToRun("EngTip", App.Path & "\data\" & App.EXEName & ".exe")
    Else
       Call RemoveFromRun("EngTip")
    End If
End Sub


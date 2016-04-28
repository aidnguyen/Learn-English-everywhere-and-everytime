Attribute VB_Name = "mdlLangStr"
Option Explicit

Public LANG_STRING_00 As String
Public LANG_STRING_01 As String
Public LANG_STRING_02 As String
Public LANG_STRING_03 As String
Public LANG_STRING_04 As String
Public LANG_STRING_05 As String
Public LANG_STRING_06 As String
Public LANG_STRING_07 As String
Public LANG_STRING_08 As String
Public LANG_STRING_09 As String
Public LANG_STRING_10 As String
Public LANG_STRING_11 As String
Public LANG_STRING_12 As String
Public LANG_STRING_13 As String
Public LANG_STRING_14 As String
Public LANG_STRING_15 As String
Public LANG_STRING_16 As String
Public LANG_STRING_17 As String
Public LANG_STRING_18 As String
Public LANG_STRING_19 As String
Public LANG_STRING_20 As String
Public LANG_STRING_21 As String
Public LANG_STRING_22 As String
Public LANG_STRING_23 As String
Public LANG_STRING_24 As String
Public LANG_STRING_25 As String
Public LANG_STRING_26 As String
Public LANG_STRING_27 As String
Public LANG_STRING_28 As String
Public LANG_STRING_29 As String
Public LANG_STRING_30 As String
Public LANG_STRING_31 As String
Public LANG_STRING_32 As String

Public Sub ChangeLang()
On Error Resume Next
Dim strContent As String
Dim strFilePath As String
strFilePath = App.Path & "\lang\" & strLangFile & ".txt"
Dim intFile As Long
intFile = FreeFile
    Open strFilePath For Binary As intFile
    strContent = InputB(FileLen(strFilePath), intFile)
    Close intFile
    Dim strArr() As String
    strArr() = Split(strContent, Chr(13) & Chr(10))
    If UBound(strArr) = 32 And Err.Number = 0 Then
        LANG_STRING_00 = Split(strArr(0), "#")(1)
        LANG_STRING_01 = strArr(1)
        LANG_STRING_02 = strArr(2)
        LANG_STRING_03 = strArr(3)
        LANG_STRING_04 = strArr(4)
        LANG_STRING_05 = strArr(5)
        LANG_STRING_06 = strArr(6)
        LANG_STRING_07 = strArr(7)
        LANG_STRING_08 = strArr(8)
        LANG_STRING_09 = strArr(9)
        LANG_STRING_10 = strArr(10)
        LANG_STRING_11 = strArr(11)
        LANG_STRING_12 = strArr(12)
        LANG_STRING_13 = strArr(13)
        LANG_STRING_14 = strArr(14)
        LANG_STRING_15 = strArr(15)
        LANG_STRING_16 = strArr(16)
        LANG_STRING_17 = strArr(17)
        LANG_STRING_18 = strArr(18)
        LANG_STRING_19 = strArr(19)
        LANG_STRING_20 = strArr(20) '
        LANG_STRING_21 = strArr(21)
        LANG_STRING_22 = strArr(22)
        LANG_STRING_23 = strArr(23)
        LANG_STRING_24 = strArr(24)
        LANG_STRING_25 = strArr(25)
        LANG_STRING_26 = strArr(26)
        LANG_STRING_27 = strArr(27)
        LANG_STRING_28 = strArr(28)
        LANG_STRING_29 = strArr(29)
        LANG_STRING_30 = strArr(30)
        LANG_STRING_31 = strArr(31)
        LANG_STRING_32 = strArr(32)
    Else
        strLangFile = "english" 'set default
        LANG_STRING_00 = "Phan mem hoc tu vung"
        LANG_STRING_01 = "Danh sach tu vung"
        LANG_STRING_02 = "Tu vung"
        LANG_STRING_03 = "Nghia cua tu"
        LANG_STRING_04 = "Them moi"
        LANG_STRING_05 = "Cap nhat"
        LANG_STRING_06 = "Xoa tu"
        LANG_STRING_07 = "Hinh anh"
        LANG_STRING_08 = "Nap hinh"
        LANG_STRING_09 = "Xoa hinh"
        LANG_STRING_10 = "Tim kiem"
        LANG_STRING_11 = "Cho phep khoi dong cung Windows"
        LANG_STRING_12 = "Khoang thoi gian hien thi giua 2 tu"
        LANG_STRING_13 = "Moi tu se hien thi trong khoang thoi gian"
        LANG_STRING_14 = "Hien thi tu vung theo"
        LANG_STRING_15 = "tuan tu"
        LANG_STRING_16 = "ngau nghien"
        LANG_STRING_17 = "Khong hien thi hinh anh"
        LANG_STRING_18 = "Ngon ngu hien thi"
        LANG_STRING_19 = "Thoat"
        LANG_STRING_20 = "Bat tooltip"
        LANG_STRING_21 = "Tat tooltip"
        LANG_STRING_22 = "Bang dieu khien"
        LANG_STRING_23 = "Gioi thieu"
        LANG_STRING_24 = "Giay"
        LANG_STRING_25 = "Phut"
        LANG_STRING_26 = "Gio"
        LANG_STRING_27 = "Tu vung cu se bi ghi de neu ban cap nhat nhung sua doi nay."
        LANG_STRING_28 = "Tu vung nay se bi xoa khoi danh sach."
        LANG_STRING_29 = "Ban co dong y khong?"
        LANG_STRING_30 = "Ban khong the cap nhat tu vung da ton tai trong danh sach."
        LANG_STRING_31 = "Huy bo"
        LANG_STRING_32 = "Sua doi"
        
    End If
Erase strArr
End Sub

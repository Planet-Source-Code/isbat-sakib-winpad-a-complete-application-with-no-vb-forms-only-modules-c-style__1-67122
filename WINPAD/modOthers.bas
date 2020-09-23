Attribute VB_Name = "modOthers"
'******************************************************************************
'******This code is downloaded from Planet Source Code
'****** Author: Isbat Sakib
'****** Email: iamsakib@gmail.com
'******************************************************************************


Option Explicit

Private OpenedAFile     As Boolean      'variable to check if new file or not


Public Sub TimeDate(ByVal hEdit As Long)

    Dim SysTime         As SYSTEMTIME
    Dim TD1             As String
    Dim TD2             As String
    Dim EText           As String
    Dim leng            As Long
    
    Call GetLocalTime(SysTime)          'Getting current date and time in SysTime
    
    leng = GetTimeFormat(LOCALE_USER_DEFAULT, TIME_NOSECONDS, SysTime, _
                                                    vbNullString, TD1, 0&)      'Returns the size of the buffer to hold time.
                                                                                'This is done if cchTime is zero.
    TD1 = String$(leng, Chr$(0))        'Initialiazing buffer
    
    Call GetTimeFormat(LOCALE_USER_DEFAULT, TIME_NOSECONDS, SysTime, _
                                                    vbNullString, TD1, leng)    'Gets the time in TD1. Now size of buffer is in cchTime.
    
    TD1 = Left$(TD1, leng - 1) & " "    'Stripping out the null character at the end and inserting a space
    
    leng = GetDateFormat(LOCALE_USER_DEFAULT, DATE_SHORTDATE, SysTime, _
                                                    vbNullString, TD2, 0&)      'Returns the size of the buffer to hold date.
                                                                                'This is done if cchTime is zero.
    TD2 = String$(leng, Chr$(0))        'Initialiazing buffer
    
    Call GetDateFormat(LOCALE_USER_DEFAULT, DATE_SHORTDATE, SysTime, _
                                                    vbNullString, TD2, leng)    'Gets the date in TD2. Now size of buffer is in cchTime.
    
    EText = TD1 & TD2                   'This is the final string and is sent to the editbox
    
    Call SendMessageByString(hEdit, EM_REPLACESEL, True, EText)         'Here if wParam is True, then the replacement operation can be undone.
                                                                        'Otherwise if False, then cannot be undone. If no text is selected then
                                                                        'text is inserted at the current cursor location. The text is passed in lParam.
    
End Sub

Public Function GetFromRegistry(ByVal AppSection As String, ByVal Key As String, _
                                ByVal DataType As RegistryDataType, _
                                Optional ByVal Default As Variant = Empty) As Variant
    
    Dim StringValue     As String
    Dim LongValue       As Long
    Dim ret             As Long
    Dim KeyHandle       As Long
    Dim LenOfBuffer     As Long
    
    If IsEmpty(Default) Then
        
        If DataType = RegistryLong Then
            Default = 0
        Else
            Default = ""
        
        End If
    
    End If
    
    ret = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\" & AppSection, 0&, KEY_ALL_ACCESS, KeyHandle)
    

    If ret <> 0 Then
        GetFromRegistry = Default
        Exit Function
    End If
    
    If DataType = RegistryString Then
    
        ret = RegQueryValueExByString(KeyHandle, Key, 0&, DataType, vbNullString, LenOfBuffer)
        StringValue = String$(LenOfBuffer, Chr$(0))
        ret = RegQueryValueExByString(KeyHandle, Key, 0&, DataType, StringValue, LenOfBuffer)
        
        If ret <> 0 Then
            GetFromRegistry = Default
            Exit Function
        End If
        
        Call RegCloseKey(KeyHandle)
        
        StringValue = Left$(StringValue, LenOfBuffer - 1)
        GetFromRegistry = StringValue
        
    Else
    
        ret = RegQueryValueExByLong(KeyHandle, Key, 0&, DataType, LongValue, 4)
        
        If ret <> 0 Then
            GetFromRegistry = Default
        Else
            GetFromRegistry = LongValue
        End If
        
        Call RegCloseKey(KeyHandle)
        
    End If
    
End Function


Public Sub SaveInRegistry(ByVal AppSection As String, ByVal Key As String, _
                          ByVal DataType As RegistryDataType, _
                          ByVal Setting As Variant)
    
    Dim KeyHandle       As Long
    Dim ret             As Long
    Dim ret2            As Long
    Dim LongValue       As Long
    Dim StringValue     As String
    
    ret = RegCreateKeyEx(HKEY_LOCAL_MACHINE, "Software\" & AppSection, 0&, _
                            vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                                                            0&, KeyHandle, ret2)
    
    If ret = 0 Then     'means successful
        
        If DataType = RegistryString Then
            
            StringValue = Setting & Chr$(0)
            Call RegSetValueExByString(KeyHandle, Key, 0&, DataType, StringValue, _
                                                                Len(StringValue))
        
        Else
        
            LongValue = Setting
            Call RegSetValueExByLong(KeyHandle, Key, 0&, DataType, LongValue, 4)
        
        End If
        
        Call RegCloseKey(KeyHandle)
        
    End If
    
End Sub

Public Sub GetFontInfoFromRegistry(ByRef lf As LOGFONT)
    
    Dim FontName        As String
    Dim temp()          As Byte
    Dim BufferLen       As Long
    Dim i               As Long
    
    With lf
        .lfHeight = GetFromRegistry("WinPad", "lfHeight", RegistryLong, 12)
        .lfWidth = GetFromRegistry("WinPad", "lfWidth", RegistryLong, 0)
        .lfEscapement = GetFromRegistry("WinPad", "lfEscapement", RegistryLong, 0)
        .lfOrientation = GetFromRegistry("WinPad", "lfOrientation", RegistryLong, 0)
        .lfWeight = GetFromRegistry("WinPad", "lfWeight", RegistryLong, FW_NORMAL)
        .lfItalic = CByte(GetFromRegistry("WinPad", "lfItalic", RegistryLong, 0))
        .lfUnderline = CByte(GetFromRegistry("WinPad", "lfUnderline", RegistryLong, 0))
        .lfStrikeOut = CByte(GetFromRegistry("WinPad", "lfStrikeOut", RegistryLong, 0))
        .lfCharSet = CByte(GetFromRegistry("WinPad", "lfCharSet", RegistryLong, ANSI_CHARSET))
        .lfOutPrecision = CByte(GetFromRegistry("WinPad", "lfOutPrecision", RegistryLong, OUT_DEFAULT_PRECIS))
        .lfClipPrecision = CByte(GetFromRegistry("WinPad", "lfClipPrecision", RegistryLong, CLIP_DEFAULT_PRECIS))
        .lfQuality = CByte(GetFromRegistry("WinPad", "lfQuality", RegistryLong, DEFAULT_QUALITY))
        .lfPitchAndFamily = CByte(GetFromRegistry("WinPad", "lfPitchAndFamily", RegistryLong, 0))
        FontName = GetFromRegistry("WinPad", "lfFaceName", RegistryString, "FixedSys")
        temp = StrConv(FontName, vbFromUnicode)
        BufferLen = UBound(temp) + 1
    
        If BufferLen > LF_FACESIZE - 1 Then
            BufferLen = LF_FACESIZE - 1
        End If
    
        For i = 0 To BufferLen - 1
            .lfFaceName(i) = temp(i)
        Next i
        
        .lfFaceName(i) = 0
    End With
    
End Sub

Public Sub SetFontInfoIntoRegistry(ByRef lf As LOGFONT)

    Dim FontName        As String
    Dim i               As Long

    With lf
        SaveInRegistry "WinPad", "lfHeight", RegistryLong, .lfHeight
        SaveInRegistry "WinPad", "lfWidth", RegistryLong, .lfWidth
        SaveInRegistry "WinPad", "lfEscapement", RegistryLong, .lfEscapement
        SaveInRegistry "WinPad", "lfOrientation", RegistryLong, .lfOrientation
        SaveInRegistry "WinPad", "lfWeight", RegistryLong, .lfWeight
        SaveInRegistry "WinPad", "lfItalic", RegistryLong, .lfItalic
        SaveInRegistry "WinPad", "lfUnderline", RegistryLong, .lfUnderline
        SaveInRegistry "WinPad", "lfStrikeOut", RegistryLong, .lfStrikeOut
        SaveInRegistry "WinPad", "lfCharSet", RegistryLong, .lfCharSet
        SaveInRegistry "WinPad", "lfOutPrecision", RegistryLong, .lfOutPrecision
        SaveInRegistry "WinPad", "lfClipPrecision", RegistryLong, .lfClipPrecision
        SaveInRegistry "WinPad", "lfQuality", RegistryLong, .lfQuality
        SaveInRegistry "WinPad", "lfPitchAndFamily", RegistryLong, .lfPitchAndFamily
        
        FontName = StrConv(.lfFaceName, vbUnicode)
        
        SaveInRegistry "WinPad", "lfFaceName", RegistryString, FontName
    End With

End Sub


Public Function OpenFile(ByVal hwnd As Long, ByVal hEdit As Long, ByRef FileNameGiven As Boolean, Optional FileName As String) As Boolean

    Dim FileText            As String
    Dim hFile               As Long
    Dim FileLength          As Long
    Dim BytesRead           As Long
    Dim ret                 As Long
    Dim ApplicationTitle    As String
    
    If FileNameGiven And (FileName = vbNullString Or FileName = "") Then
        Exit Function
    End If
    
    Call SetForegroundWindow(gHwnd)
    
    If gEditChanged Then
        ret = MessageBox(hwnd, "The text in the " & gPathOfFile & " file has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", "Text changed", MB_YESNOCANCEL Or MB_ICONEXCLAMATION)
        If ret = IDYES Then
            If SaveFile(hwnd, hEdit) = False Then
                OpenFile = False
                Exit Function
            End If
        ElseIf ret = IDCANCEL Then
            OpenFile = False
            Exit Function
        End If
    End If
    
    If FileNameGiven = False Then
    
        FileName = ShowOpen(, "Text Files (*.txt)|*.txt|All Files (*.*)|*.*", , "*.txt", , _
                                FileMustExist Or HideReadOnly Or PathMustExist, hwnd)
    End If
    
    
    If FileName <> vbNullString Or FileName <> "" Then
        
        hFile = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ, 0&, OPEN_EXISTING, _
                                                            FILE_ATTRIBUTE_NORMAL, 0&)
        
        If hFile = INVALID_HANDLE_VALUE Then
            Call MessageBox(hwnd, "Error occured while opening the file.", "Error", MB_OK Or MB_ICONSTOP)
            OpenFile = False
            Exit Function
        End If
        
        FileLength = GetFileSize(hFile, 0&)
        
        If FileLength > 60000 Then
            Call MessageBox(hwnd, "This file is too large for WinPad to open.", "Error", MB_OK)
            Call CloseHandle(hFile)
            OpenFile = False
            Exit Function
        End If
            
        Call PostMessage(hEdit, WM_SETREDRAW, False, 0&)
        
        FileText = String$(FileLength, Chr$(0))
        
        ret = ReadFile(hFile, ByVal FileText, FileLength, BytesRead, 0&)
        If ret = 0 Then
            Call MessageBox(hwnd, "Error occured while opening the file.", "Error", MB_OK Or MB_ICONSTOP)
            Call CloseHandle(hFile)
            Call PostMessage(hEdit, WM_SETREDRAW, True, 0&)
            OpenFile = False
            Exit Function
        End If
        
        FileText = Replace(FileText, Chr$(0), Chr$(32))
        
        Call SendMessageByString(hEdit, WM_SETTEXT, 0&, vbNullString)
        Call SendMessageByString(hEdit, WM_SETTEXT, 0&, FileText)
        Call CloseHandle(hFile)
        Call PostMessage(hEdit, WM_SETREDRAW, True, 0&)

        OpenedAFile = True
        gPathOfFile = FileName
        ret = InStrRev(FileName, "\")
        gNameOfFile = Right$(FileName, Len(FileName) - ret)
        ApplicationTitle = gNameOfFile & " - WinPad"
        Call SendMessageByString(hwnd, WM_SETTEXT, 0&, ApplicationTitle)
        gEditChanged = False
        OpenFile = True
    End If
    
End Function

Public Function SaveFile(ByVal hwnd As Long, ByVal hEdit As Long) As Boolean
    
    Dim hFile               As Long
    Dim ret                 As Long
    Dim ApplicationTitle    As String
    
    If OpenedAFile Then
        hFile = CreateFile(gPathOfFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
        If hFile <> INVALID_HANDLE_VALUE Then
            
            CloseHandle (hFile)
            
            ret = GetFileAttributes(gPathOfFile)
            If (ret And FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY Then
                SaveFile = SaveFileAs(hwnd, hEdit)
                Exit Function
            End If
            
            hFile = CreateFile(gPathOfFile, GENERIC_WRITE, FILE_SHARE_WRITE, 0&, TRUNCATE_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
            
            If hFile = INVALID_HANDLE_VALUE Then
                Call MessageBox(hwnd, "Error occured while saving the file.", "Error", MB_OK Or MB_ICONSTOP)
                SaveFile = False
                Exit Function
            End If
            
            If WriteToFile(hwnd, hEdit, hFile) = 0 Then
                Call MessageBox(hwnd, "Error occured while saving the file.", "Error", MB_OK Or MB_ICONSTOP)
                Call CloseHandle(hFile)
                SaveFile = False
                Exit Function
            End If
            
            Call CloseHandle(hFile)
            gEditChanged = False
            SaveFile = True
            
        Else
            
            SaveFile = SaveFileAs(hwnd, hEdit)
            
        End If
            
    Else
        SaveFile = SaveFileAs(hwnd, hEdit)
    End If
        
End Function

Public Function SaveFileAs(ByVal hwnd As Long, ByVal hEdit As Long) As Boolean

    Dim FileName            As String
    Dim hFile               As Long
    Dim ret                 As Long
    Dim ApplicationTitle    As String
    
    FileName = ShowSave(, "Text Files (*.txt)|*.txt|All Files (*.*)|*.*", , gPathOfFile, , HideReadOnly Or OverWritePrompt Or NoReadOnlyReturn, hwnd)
    
    If FileName = vbNullString Then
        SaveFileAs = False
        Exit Function
    End If
    
    hFile = CreateFile(FileName, 0&, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If hFile <> INVALID_HANDLE_VALUE Then
        Call CloseHandle(hFile)
        hFile = CreateFile(FileName, GENERIC_WRITE, FILE_SHARE_WRITE, 0&, TRUNCATE_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    Else
        hFile = CreateFile(FileName, GENERIC_WRITE, 0&, 0&, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, 0&)
    End If
    
    If hFile = INVALID_HANDLE_VALUE Then
        Call MessageBox(hwnd, "Error occured while saving the file.", "Error", MB_OK Or MB_ICONSTOP)
        SaveFileAs = False
        Exit Function
    End If
    
    If WriteToFile(hwnd, hEdit, hFile) = 0 Then
        Call MessageBox(hwnd, "Error occured while saving the file.", "Error", MB_OK Or MB_ICONSTOP)
        Call CloseHandle(hFile)
        SaveFileAs = False
        Exit Function
    End If
    
    Call CloseHandle(hFile)
    
    gPathOfFile = FileName
    ret = InStrRev(FileName, "\")
    gNameOfFile = Right$(FileName, Len(FileName) - ret)
    ApplicationTitle = gNameOfFile & " - WinPad"
    
    Call SendMessageByString(hwnd, WM_SETTEXT, 0&, ApplicationTitle)
    gEditChanged = False
    SaveFileAs = True
    OpenedAFile = True
    
End Function

Private Function WriteToFile(ByVal hwnd As Long, ByVal hEdit As Long, ByVal hFile As Long) As Long

    Dim EditText            As String
    Dim EditTextLength      As Long
    Dim BytesWritten        As Long
    
    EditTextLength = SendMessageByNum(hEdit, WM_GETTEXTLENGTH, 0&, 0&)
    EditText = String$(EditTextLength + 1, Chr$(0))
    Call SendMessageByString(hEdit, WM_GETTEXT, EditTextLength + 1, EditText)
    
    WriteToFile = WriteFile(hFile, ByVal EditText, EditTextLength, BytesWritten, 0&)
    
End Function

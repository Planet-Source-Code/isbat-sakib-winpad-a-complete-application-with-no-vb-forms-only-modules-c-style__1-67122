Attribute VB_Name = "modMain"
'******************************************************************************
'******This code is downloaded from Planet Source Code
'****** Author: Isbat Sakib
'****** Email: iamsakib@gmail.com
'******************************************************************************


Option Explicit

Public FindInvoked              As Boolean

Private Function CreateWindows() As Boolean

    gHwnd = CreateWindowEx(WS_EX_ACCEPTFILES, gClassName, gAppName, WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 0&, 0&, GetModuleHandle(vbNullString), ByVal 0&)
    If (gHwnd <> 0) Then
        Call ShowWindow(gHwnd, SW_SHOWNORMAL)
        Call SetForegroundWindow(gHwnd)
        Call UpdateWindow(gHwnd)
    End If
    CreateWindows = (gHwnd <> 0)

End Function

Private Function GetAddress(ByRef lngAddr As Long) As Long

    GetAddress = lngAddr

End Function

Public Function HiWord(ByRef Parameter As Long) As Long

    HiWord = Parameter \ 65536

End Function

Public Function LoWord1(ByRef Parameter As Long) As Long

    LoWord1 = Parameter And &HFFFF&

End Function

Public Function LoWord(ByRef DbleWord As Long) As Long

    If DbleWord And &H8000& Then
        LoWord = &H8000 Or (DbleWord And &H7FFF&)
    Else
        LoWord = DbleWord And &HFFFF&
    End If
    
End Function

Private Function MakeLong(ByVal LoWord As Long, ByVal HiWord As Long) As Long

    HiWord = HiWord * (2 ^ 16)
    MakeLong = LoWord Or HiWord

End Function

Public Sub Main()

    Dim wMsg    As msg

    If RegisterWindowClass = False Then
        Exit Sub
    End If

    If CreateWindows Then
        Do While (GetMessage(wMsg, 0&, 0&, 0&) > 0)
            If TranslateAccelerator(gHwnd, gAccTable, wMsg) = 0 Then
                Call TranslateMessage(wMsg)
                Call DispatchMessage(wMsg)
            End If
        Loop
    End If

    Call UnregisterClass(gClassName, App.hInstance)

End Sub

Private Function RegisterWindowClass() As Boolean

    Dim wc      As WNDCLASS

    With wc
        .style = 0&
        .lpfnwndproc = GetAddress(AddressOf WndProc)
        .hInstance = App.hInstance
        .hIcon = LoadIconByNum(0&, IDI_APPLICATION)
        .hCursor = LoadCursorByNum(0&, IDC_ARROW)
        .hbrBackground = COLOR_WINDOW
        .lpszClassName = gClassName
        .lpszMenuName = vbNullString
        .cbClsextra = 0&
        .cbWndExtra2 = 0&
    End With
    RegisterWindowClass = RegisterClass(wc) <> 0

End Function

Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim ret                     As Long
    Dim hDrop                   As Long
    Dim rc                      As RECT
    Dim pt                      As POINTAPI
    Dim FileName                As String
    
    If uMsg = msgFind Then
        
        ret = FindTextInEdit(lParam, gEditHwnd)
            
        If ret <> 0 Then
            Call MessageBox(hwnd, "Error occured.", "Error", MB_ICONSTOP)
        End If
        
        WndProc = 0
        Exit Function
        
    End If
    
    Select Case uMsg
    
        Case WM_CREATE
            
            CreateMenus hwnd
            CreateKeyboardShortcuts
            gBkBrush = CreateSolidBrush(16777215)
            gEditHwnd = CreateEditBox(hwnd)
            
            msgFind = RegisterWindowMessage(FINDMSGSTRING)
            
            Call SendMessageByString(hwnd, WM_SETTEXT, 0&, "Untitled - WinPad")
            gPathOfFile = "Untitled"
                          
            
        Case WM_CLOSE
            
            If gEditChanged Then
                ret = MessageBox(hwnd, "The text in the " & gPathOfFile & " file has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", "Text changed", MB_YESNOCANCEL Or MB_ICONEXCLAMATION)
                If ret = IDYES Then
                    If SaveFile(hwnd, gEditHwnd) = False Then
                        WndProc = 0
                        Exit Function
                    End If
                ElseIf ret = IDCANCEL Then
                    WndProc = 0
                    Exit Function
                End If
            End If
            
     
            Call DeleteObject(ghFont)
            Call DeleteObject(gBkBrush)
            Call DestroyAcceleratorTable(gAccTable)
            
            If ghDevMode Then
                Call GlobalFree(ghDevMode)
            End If
            
            If ghDevNames Then
                Call GlobalFree(ghDevNames)
            End If
            
            Call DestroyWindow(hwnd)
            
        Case WM_DESTROY
            Call PostQuitMessage(0&)
    
        Case WM_SETFOCUS
            Call SetFocus(gEditHwnd)
    
        Case WM_SIZE
            Call MoveWindow(gEditHwnd, 0, 0, LoWord(lParam), HiWord(lParam), True)

        Case WM_INITMENUPOPUP
            CheckMenuStates
            
        Case WM_CTLCOLOREDIT
            Call SetTextColor(wParam, gTextColor)
            Call SetBkMode(wParam, OPAQUE)
            Call SetBkColor(wParam, 16777215)
            WndProc = gBkBrush
            Exit Function
            
        Case WM_DROPFILES
            FileName = String$(260, Chr$(0))
            Call DragQueryFile(wParam, 0&, FileName, 261)
            Call DragFinish(wParam)
            
            ret = InStr(FileName, Chr$(0))
            If ret > 0 Then
                FileName = IIf(ret = 1, vbNullString, Left$(FileName, ret - 1))
            End If
            OpenFile hwnd, gEditHwnd, True, FileName
            
        Case WM_COMMAND
            Select Case (LoWord(wParam))
                
                Case ID_FILE_NEW
                    If gEditChanged Then
                        ret = MessageBox(hwnd, "The text in the " & gPathOfFile & " file has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", "Text changed", MB_YESNOCANCEL Or MB_ICONEXCLAMATION)
                        If ret = IDYES Then
                            If SaveFile(hwnd, gEditHwnd) = False Then
                                WndProc = 0
                                Exit Function
                            End If
                        ElseIf ret = IDCANCEL Then
                            WndProc = 0
                            Exit Function
                        End If
                    End If
                    
                    Call SendMessageByString(gEditHwnd, WM_SETTEXT, 0&, vbNullString)
                    Call SendMessageByString(hwnd, WM_SETTEXT, 0&, "Untitled - WinPad")
                    gEditChanged = False
                    gPathOfFile = "Untitled"
                    FindInvoked = False
                    
                Case ID_FILE_OPEN
                    OpenFile hwnd, gEditHwnd, False
                    
                Case ID_FILE_SAVE
                    Call SaveFile(hwnd, gEditHwnd)
                    
                Case ID_FILE_SAVEAS
                    Call SaveFileAs(hwnd, gEditHwnd)
                    
                Case ID_FILE_PAGESETUP
                    'I dropped the development here.
                    'Maybe you could continue it?
                    
                Case ID_FILE_PRINT
                    'I dropped the development here.
                    'Maybe you could continue it?
                    
                Case ID_FILE_EXIT
                    Call PostMessage(hwnd, WM_CLOSE, 0&, 0&)
                    
                Case ID_EDIT_UNDO
                    Call SendMessageByNum(gEditHwnd, WM_UNDO, 0&, 0&)
                    
                Case ID_EDIT_CUT
                    Call SendMessageByNum(gEditHwnd, WM_CUT, 0&, 0&)
                    
                Case ID_EDIT_COPY
                    Call SendMessageByNum(gEditHwnd, WM_COPY, 0&, 0&)
                    
                Case ID_EDIT_PASTE
                    Call SendMessageByNum(gEditHwnd, WM_PASTE, 0&, 0&)
                    
                Case ID_EDIT_DELETE
                    Call SendMessageByNum(gEditHwnd, WM_CLEAR, 0&, 0&)
                
                Case ID_EDIT_SELECTALL
                    Call SendMessageByNum(gEditHwnd, EM_SETSEL, 0&, -1&)
                    Call SendMessageByNum(gEditHwnd, EM_SCROLLCARET, 0&, 0&)
                    
                Case ID_EDIT_TIMEDATE
                    TimeDate gEditHwnd
                    
                Case ID_EDIT_WORDWRAP
                    ret = GetMenuState(GetMenu(hwnd), ID_EDIT_WORDWRAP, MF_BYCOMMAND)
                    If Not (ret And MF_CHECKED) = MF_CHECKED Then
                        Call CheckMenuItem(GetMenu(hwnd), ID_EDIT_WORDWRAP, MF_BYCOMMAND Or MF_CHECKED)
                        SaveInRegistry "WinPad", "WordWrap", RegistryLong, 1
                        WordWrapSub hwnd, gEditHwnd, True
                    Else
                        Call CheckMenuItem(GetMenu(hwnd), ID_EDIT_WORDWRAP, MF_BYCOMMAND Or MF_UNCHECKED)
                        SaveInRegistry "WinPad", "WordWrap", RegistryLong, 0
                        WordWrapSub hwnd, gEditHwnd, False
                    End If
                    
                Case ID_EDIT_SETFONT
                    ChooseFontForEditBox hwnd
                    
                Case ID_SEARCH_FIND
                    FindInvoked = True
                    SearchForText
                
                Case ID_SEARCH_FINDNEXT
                    If FindInvoked = False Then
                        SearchForText
                        WndProc = 0
                        FindInvoked = True
                        Exit Function
                    End If
                    
                    ret = FindNext(gEditHwnd)
                    
                    If ret <> 0 Then
                        Call MessageBox(hwnd, "Error occured.", "Error", MB_ICONSTOP)
                    End If
            
                    WndProc = 0
                    Exit Function
                    
                Case ID_HELP_ABOUT
                    Call ShellAbout(hwnd, "WinPad", "WinPad is created by Isbat Sakib." & vbCrLf & "Email: sakib039@hotmail.com", 0&)
                    
                Case IDC_MAIN_EDIT
                    Select Case HiWord(wParam)
                        Case EN_MAXTEXT
                            Call MessageBox(hwnd, "Maximum limit of text has been reached. Cannot insert any more text.", "Maximum limit reached", MB_ICONSTOP Or MB_OK)
                        
                        Case EN_ERRSPACE
                            Call MessageBox(hwnd, "Not enough memory available to complete this operation. Quit one or more applications to increase available memory, and then try again.", "Shortage of Memory", MB_ICONSTOP Or MB_OK)
                        
                        Case EN_CHANGE
                            gEditChanged = True
                            
                    End Select
            
            End Select
            
            
        Case Else
            WndProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
            Exit Function
    
    End Select
        
    WndProc = 0

End Function

Private Sub WordWrapSub(ByVal hwnd As Long, ByRef hEdit As Long, ByRef WordWrap As Boolean)

    Dim OldLength       As Long
    Dim NewEditHwnd     As Long
    Dim hfont           As Long
    Dim TextSrc         As String
    Dim rc              As RECT
    Dim ret             As Long
    
    Call GetClientRect(hwnd, rc)
    
    If WordWrap Then
        NewEditHwnd = CreateWindowEx(WS_EX_CLIENTEDGE, "EDIT", vbNullString, WS_CHILD Or WS_VSCROLL Or ES_MULTILINE Or ES_LEFT Or ES_AUTOVSCROLL Or ES_NOHIDESEL Or ES_WANTRETURN, rc.Top, rc.Left, rc.Right, rc.Bottom, hwnd, IDC_MAIN_EDIT, GetWindowLong(hwnd, GWL_HINSTANCE), ByVal 0&)
    Else
        NewEditHwnd = CreateWindowEx(WS_EX_CLIENTEDGE, "EDIT", vbNullString, WS_CHILD Or WS_VSCROLL Or WS_HSCROLL Or ES_MULTILINE Or ES_LEFT Or ES_AUTOVSCROLL Or ES_AUTOHSCROLL Or ES_NOHIDESEL Or ES_WANTRETURN, rc.Top, rc.Left, rc.Right, rc.Bottom, hwnd, IDC_MAIN_EDIT, GetWindowLong(hwnd, GWL_HINSTANCE), ByVal 0&)
    End If
    
    If NewEditHwnd = 0 Then
        Call MessageBox(hwnd, "Error in WordWrap process", "Error", MB_OK Or MB_ICONHAND)
        Exit Sub
    End If
    
    Call SendMessageByNum(NewEditHwnd, EM_LIMITTEXT, 60000, 0&)
    
    Call DeleteObject(ghFont)
    hfont = CreateFontForEditBox(hwnd)
    Call SendMessageByNum(NewEditHwnd, WM_SETFONT, hfont, 0&)
    
        
    OldLength = SendMessageByNum(hEdit, WM_GETTEXTLENGTH, 0&, 0&)
    TextSrc = String$(OldLength + 1, Chr$(0))
    Call SendMessageByString(hEdit, WM_GETTEXT, OldLength + 1, TextSrc)
    Call DestroyWindow(hEdit)
    Call ShowWindow(NewEditHwnd, SW_SHOWNORMAL)
    Call SendMessageByString(NewEditHwnd, WM_SETTEXT, 0&, TextSrc)
    
    Call SetFocus(NewEditHwnd)
    
    hEdit = NewEditHwnd
    
End Sub

Private Sub CreateMenus(ByRef hwnd As Long)
    
    Dim hMenu       As Long
    Dim hSubmenu    As Long
    
    hMenu = CreateMenu()
    ghMenu = hMenu
    
    hSubmenu = CreatePopupMenu()
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_FILE_NEW, "&New" & vbTab & "Ctrl+N")
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_FILE_OPEN, "&Open..." & vbTab & "Ctrl+O")
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_FILE_SAVE, "&Save" & vbTab & "Ctrl+S")
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_FILE_SAVEAS, "Save &As...")
    Call AppendMenuByString(hSubmenu, MF_SEPARATOR, 5000, "-")
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_FILE_PAGESETUP, "Page Se&tup...")
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_FILE_PRINT, "&Print" & vbTab & "Ctrl+P")
    Call AppendMenuByString(hSubmenu, MF_SEPARATOR, 5001, "-")
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_FILE_EXIT, "E&xit" & vbTab & "Ctrl+Q")
    Call AppendMenuByString(hMenu, MF_STRING Or MF_POPUP, hSubmenu, "&File")
            
    hSubmenu = CreatePopupMenu()
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_EDIT_UNDO, "&Undo" & vbTab & "Ctrl+Z")
    Call AppendMenuByString(hSubmenu, MF_SEPARATOR, 5002, "-")
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_EDIT_CUT, "Cu&t" & vbTab & "Ctrl+X")
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_EDIT_COPY, "&Copy" & vbTab & "Ctrl+C")
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_EDIT_PASTE, "&Paste" & vbTab & "Ctrl+V")
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_EDIT_DELETE, "De&lete" & vbTab & "Del")
    Call AppendMenuByString(hSubmenu, MF_SEPARATOR, 5003, "-")
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_EDIT_SELECTALL, "Select &All" & vbTab & "Ctrl+A")
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_EDIT_TIMEDATE, "Time/&Date" & vbTab & "F5")
    Call AppendMenuByString(hSubmenu, MF_SEPARATOR, 5004, "-")
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_EDIT_WORDWRAP, "&Word Wrap")
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_EDIT_SETFONT, "Set &Font...")
    Call AppendMenuByString(hMenu, MF_STRING Or MF_POPUP, hSubmenu, "&Edit")
            
    hSubmenu = CreatePopupMenu()
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_SEARCH_FIND, "&Find..." & vbTab & "F2")
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_SEARCH_FINDNEXT, "Find &Next" & vbTab & "F3")
    Call AppendMenuByString(hMenu, MF_STRING Or MF_POPUP, hSubmenu, "&Search")
            
    hSubmenu = CreatePopupMenu()
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_HELP_HELPTOPICS, "&Help Topics")
    Call AppendMenuByString(hSubmenu, MF_SEPARATOR, 5005, "-")
    Call AppendMenuByString(hSubmenu, MF_STRING, ID_HELP_ABOUT, "&About Winpad")
    Call AppendMenuByString(hMenu, MF_STRING Or MF_POPUP, hSubmenu, "&Help")
            
    Call SetMenu(hwnd, hMenu)
    
End Sub

Private Sub CreateKeyboardShortcuts()

    Dim AcTables(13)            As ACCEL
    Dim i                       As Long
    
    For i = 0 To 9
        AcTables(i).fVirt = FCONTROL Or FNOINVERT Or FVIRTKEY
    Next i
    
    For i = 10 To 13
        AcTables(i).fVirt = FNOINVERT Or FVIRTKEY
    Next i
    
    AcTables(0).Key = VkKeyScan(AscB("n"))  'passing ASCII value of 'n' to VkKeyScan to get the virtual key-code
    AcTables(0).cmd = ID_FILE_NEW
    
    AcTables(1).Key = VkKeyScan(AscB("o"))  'passing ASCII value of 'o' to VkKeyScan to get the virtual key-code
    AcTables(1).cmd = ID_FILE_OPEN
    
    AcTables(2).Key = VkKeyScan(AscB("s"))  'passing ASCII value of 's' to VkKeyScan to get the virtual key-code
    AcTables(2).cmd = ID_FILE_SAVE
    
    AcTables(3).Key = VkKeyScan(AscB("p"))  'passing ASCII value of 'p' to VkKeyScan to get the virtual key-code
    AcTables(3).cmd = ID_FILE_PRINT
    
    AcTables(4).Key = VkKeyScan(AscB("q"))  'passing ASCII value of 'q' to VkKeyScan to get the virtual key-code
    AcTables(4).cmd = ID_FILE_EXIT
    
    AcTables(5).Key = VkKeyScan(AscB("z"))  'passing ASCII value of 'z' to VkKeyScan to get the virtual key-code
    AcTables(5).cmd = ID_EDIT_UNDO
    
    AcTables(6).Key = VkKeyScan(AscB("x"))  'passing ASCII value of 'x' to VkKeyScan to get the virtual key-code
    AcTables(6).cmd = ID_EDIT_CUT
    
    AcTables(7).Key = VkKeyScan(AscB("c"))  'passing ASCII value of 'c' to VkKeyScan to get the virtual key-code
    AcTables(7).cmd = ID_EDIT_COPY
    
    AcTables(8).Key = VkKeyScan(AscB("v"))  'passing ASCII value of 'v' to VkKeyScan to get the virtual key-code
    AcTables(8).cmd = ID_EDIT_PASTE
    
    AcTables(9).Key = VkKeyScan(AscB("a"))  'passing ASCII value of 'a' to VkKeyScan to get the virtual key-code
    AcTables(9).cmd = ID_EDIT_SELECTALL
    
    AcTables(10).Key = VK_F5
    AcTables(10).cmd = ID_EDIT_TIMEDATE
    
    AcTables(11).Key = VK_F2
    AcTables(11).cmd = ID_SEARCH_FIND
    
    AcTables(12).Key = VK_F3
    AcTables(12).cmd = ID_SEARCH_FINDNEXT
    
    AcTables(13).Key = VK_F1
    AcTables(13).cmd = ID_HELP_HELPTOPICS
       
    gAccTable = CreateAcceleratorTable(AcTables(0), 14)

End Sub

Private Function CreateEditBox(ByRef hwnd As Long) As Long

    Dim hfont       As Long
    Dim hEdit       As Long
    
    If GetFromRegistry("WinPad", "WordWrap", RegistryLong, 0) = 1 Then
        Call CheckMenuItem(GetMenu(hwnd), ID_EDIT_WORDWRAP, MF_BYCOMMAND Or MF_CHECKED)
        hEdit = CreateWindowEx(WS_EX_CLIENTEDGE, "EDIT", vbNullString, WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or ES_MULTILINE Or ES_AUTOVSCROLL Or ES_NOHIDESEL Or ES_WANTRETURN, 0, 0, 490, 455, hwnd, IDC_MAIN_EDIT, GetWindowLong(hwnd, GWL_HINSTANCE), ByVal 0&)
    Else
        hEdit = CreateWindowEx(WS_EX_CLIENTEDGE, "EDIT", vbNullString, WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_HSCROLL Or ES_MULTILINE Or ES_AUTOVSCROLL Or ES_AUTOHSCROLL Or ES_NOHIDESEL Or ES_WANTRETURN, 0, 0, 490, 455, hwnd, IDC_MAIN_EDIT, GetWindowLong(hwnd, GWL_HINSTANCE), ByVal 0&)
    End If
    
    CreateEditBox = hEdit
    
    If hEdit = 0 Then
        Call MessageBox(hwnd, "Could not create edit box.", "Error", MB_OK Or MB_ICONHAND)
        Exit Function
    End If
    
    Call SendMessageByNum(hEdit, EM_LIMITTEXT, 60000, 0&)
    
    hfont = CreateFontForEditBox(hwnd)
    Call SendMessageByNum(hEdit, WM_SETFONT, hfont, 0&)
    gTextColor = GetFromRegistry("WinPad", "TextColor", RegistryLong, 0)
    
    Call SetFocus(hEdit)
    
End Function

Private Function CreateFontForEditBox(ByRef hwnd As Long) As Long
    
    Dim lf              As LOGFONT
    Dim ret             As Long
    
    GetFontInfoFromRegistry lf

    ret = CreateFontIndirect(lf)
    
    If ret = 0 Then
        Call MessageBox(hwnd, "Cannot create font for edit box.", "Error", MB_OK Or MB_ICONHAND)
        ret = GetStockObject(SYSTEM_FONT)
    End If
    
    ghFont = ret
    CreateFontForEditBox = ret
   
End Function

Private Sub ChooseFontForEditBox(ByRef hwnd As Long)

    Dim lf              As LOGFONT
    Dim ret             As Long
    Dim rc              As RECT
    
    GetFontInfoFromRegistry lf
    ret = ShowFont(, ForceFontExist Or ShowScreenFontsOnly Or UseEffects Or UseLogFontStructure, VarPtr(lf), gTextColor, hwnd)
    
    If ret = 0 Then
        Exit Sub
    Else
        ret = CreateFontIndirect(lf)
        If ret = 0 Then
            Call MessageBox(hwnd, "Cannot create font for edit box.", "Error", MB_OK Or MB_ICONHAND)
            Exit Sub
        End If
        
        Call DeleteObject(ghFont)
        ghFont = ret
        Call SendMessageByNum(gEditHwnd, WM_SETFONT, ghFont, 0&)
        
        Call RedrawWindow(gEditHwnd, 0&, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ERASENOW)
        SetFontInfoIntoRegistry lf
        SaveInRegistry "WinPad", "TextColor", RegistryLong, gTextColor
        
    End If
    
End Sub

Private Sub CheckMenuStates()

    Dim ret As Long

    If SendMessageByNum(gEditHwnd, EM_CANUNDO, 0&, 0&) <> 0 Then
        Call EnableMenuItem(ghMenu, ID_EDIT_UNDO, MF_BYCOMMAND Or MF_ENABLED)
    Else
        Call EnableMenuItem(ghMenu, ID_EDIT_UNDO, MF_BYCOMMAND Or MF_GRAYED)
    End If
            
    If IsClipboardFormatAvailable(CF_TEXT) <> 0 Then
        Call EnableMenuItem(ghMenu, ID_EDIT_PASTE, MF_BYCOMMAND Or MF_ENABLED)
    Else
        Call EnableMenuItem(ghMenu, ID_EDIT_PASTE, MF_BYCOMMAND Or MF_GRAYED)
    End If
            
    ret = SendMessageByNum(gEditHwnd, EM_GETSEL, 0&, 0&)
            
    If LoWord(ret) <> HiWord(ret) Then
        Call EnableMenuItem(ghMenu, ID_EDIT_CUT, MF_BYCOMMAND Or MF_ENABLED)
        Call EnableMenuItem(ghMenu, ID_EDIT_COPY, MF_BYCOMMAND Or MF_ENABLED)
        Call EnableMenuItem(ghMenu, ID_EDIT_DELETE, MF_BYCOMMAND Or MF_ENABLED)
    Else
        Call EnableMenuItem(ghMenu, ID_EDIT_CUT, MF_BYCOMMAND Or MF_GRAYED)
        Call EnableMenuItem(ghMenu, ID_EDIT_COPY, MF_BYCOMMAND Or MF_GRAYED)
        Call EnableMenuItem(ghMenu, ID_EDIT_DELETE, MF_BYCOMMAND Or MF_GRAYED)
    End If
            
    ret = SendMessageByNum(gEditHwnd, WM_GETTEXTLENGTH, 0&, 0&)
            
    If ret <> 0 Then
        Call EnableMenuItem(ghMenu, ID_EDIT_SELECTALL, MF_BYCOMMAND Or MF_ENABLED)
    Else
        Call EnableMenuItem(ghMenu, ID_EDIT_SELECTALL, MF_BYCOMMAND Or MF_GRAYED)
    End If
    
End Sub

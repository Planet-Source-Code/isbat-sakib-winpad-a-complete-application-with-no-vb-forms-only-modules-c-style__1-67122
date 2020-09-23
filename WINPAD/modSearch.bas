Attribute VB_Name = "modSearch"
'******************************************************************************
'******This code is downloaded from Planet Source Code
'****** Author: Isbat Sakib
'****** Email: iamsakib@gmail.com
'******************************************************************************


Option Explicit

'Used internally for remembering search information.
Private Type FindInfo
    SearchDown                      As Boolean
    MatchCase                       As Boolean
End Type

'API Type used in FindText function.
Private Type FINDREPLACE
    lStructSize                     As Long
    hwndOwner                       As Long
    hInstance                       As Long
    flags                           As Long
    lpstrFindWhat                   As Long
    lpstrReplaceWith                As Long
    wFindWhatLen                    As Integer
    wReplaceWithLen                 As Integer
    lCustData                       As Long
    lpfnHook                        As Long
    lpTemplateName                  As String
End Type


Public msgFind                      As Long             'Very important. This is the message registered in
                                                        'WndProc function to receive messages from Find dialog.
                                                        
Public Const FINDMSGSTRING          As String = "commdlg_FindReplace"       'This is the parameter passed in ResiterWindowMessage function
                                                                            'in WndProc to register this message and get it into msgFind.

'These are API constants used for different search options.
Private Const FR_HIDEWHOLEWORD      As Long = &H10000
Private Const FR_DIALOGTERM         As Long = &H40
Private Const FR_FINDNEXT           As Long = &H8
Private Const FR_MATCHCASE          As Long = &H4
Private Const FR_DOWN               As Long = &H1


'Just indicates the maximum length of string that is to be searched.
Private Const MaxSearchLen          As Long = 50


'API functions used only in this module.
Private Declare Function FindText Lib "comdlg32.dll" Alias "FindTextA" (pFindreplace As FINDREPLACE) As Long
Private Declare Function IsBadStringPtrByNum Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long
Private Declare Function CopyPointer2String Lib "kernel32" Alias "lstrcpyA" (ByVal NewString As String, ByVal OldString As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessageNoByVal Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Long, lParam As Long) As Long


'Important variables.
Private FndInfo                     As FindInfo
Private Fr                          As FINDREPLACE
Private SearchText                  As String
Private SearchBuffer()              As Byte

'This variable is used to indicate is FindNext button is clicked once or not.
Private FindNextClickedOnce         As Boolean

'The global handle of the Find Dialog.
Public gFindDlgHandle               As Long


'This function is called to create the Find dialog window.
Public Sub SearchForText()

    If gFindDlgHandle > 0 Then              'If the Find window already exists
        Call SetFocus(gFindDlgHandle)       'then just set the focus to it and
        Exit Sub                            'exit the sub as nothing more is required.
    End If

    If msgFind = 0 Then                     'If for some reason the message registering
                                            'failed, display a message box and exit sub.
                                            
        Call MessageBox(gHwnd, "Registering message failed. Unable to proceed.", "Register message error", MB_ICONSTOP Or MB_OK)
        Exit Sub
    End If
    
    InitializeDialog            'If everything ok, create the Find dialog.
    
End Sub

Private Sub InitializeDialog()

    Dim wMsg                       As msg          'Msg type used for the Find dialog.
    
    If FindNextClickedOnce = False Then             'If the FindNext button in the Find dialog
                                                    'window is not even clicked once, then
                                                    'initialize the search string and then the search
                                                    'buffer.
        SearchText = String$(MaxSearchLen, Chr$(0))
        SearchBuffer = StrConv(SearchText, vbFromUnicode)
    End If
    
    With Fr                             'Filling up the FINDREPLACE structure.
        .lStructSize = Len(Fr)
        .hwndOwner = gHwnd
        .hInstance = 0&
        .flags = FR_HIDEWHOLEWORD Or FR_DOWN        'Here 'Match whole word only' option
                                                    'is made unavailable.
        .lpstrFindWhat = VarPtr(SearchBuffer(0))
        .lpstrReplaceWith = 0&
        .wFindWhatLen = MaxSearchLen
        .wReplaceWithLen = 0
        .lCustData = 0
        .lpfnHook = 0&
        .lpTemplateName = 0&
    End With
   
    gFindDlgHandle = FindText(Fr)                   'This shows the Find dialog.
    
    Do While (GetMessage(wMsg, 0&, 0&, 0&) > 0) And (gFindDlgHandle > 0)      'This is the message loop for Find dialog window.
        If IsDialogMessage(gFindDlgHandle, wMsg) = 0 And TranslateAccelerator(gHwnd, gAccTable, wMsg) = 0 Then
            Call TranslateMessage(wMsg)
            Call DispatchMessage(wMsg)
        End If
    Loop
    
    If wMsg.Message = WM_QUIT Then
        Call PostQuitMessage(0&)
    End If

End Sub


'This function is called after getting msgFind message in WndProc function.
'The address of the FINDREPLACE structure is passed in lParam.

Public Function FindTextInEdit(ByRef lParam As Long, ByVal hEdit As Long) As Long

    Dim Fr1             As FINDREPLACE
    
    Call CopyMemory(Fr1, ByVal lParam, Len(Fr1))            'The FINDREPLACE structure is fetched
                                                            'from its address in lParam.
    
    If (Fr1.flags And FR_DIALOGTERM) = FR_DIALOGTERM Then   'If Cancel button is clicked,
        gFindDlgHandle = 0                                  'set the window handle to zero.
        FindTextInEdit = 0
        Exit Function                                       'Also exit this function.
    End If
            
    With Fr1
        
        If .wFindWhatLen > 0 Then                   'If there is something to search....
                
            If .lpstrFindWhat = 0 Then              'If length of search string is zero,
                SearchText = vbNullString           'then there is nothing to search.
            
            ElseIf Not IsBadStringPtrByNum(.lpstrFindWhat, .wFindWhatLen) Then  'If length is not zero and if the pointer is ok,
                SearchText = String$(MaxSearchLen, Chr$(0))                     'then initialize SearchText variable.
                        
                Call CopyPointer2String(SearchText, .lpstrFindWhat)             'Get the text to search in SearchText.
                SearchText = Left$(SearchText, InStr(SearchText, Chr$(0)) - 1)  'Strip out the null character at the end.
                    
            Else
                SearchText = vbNullString               'Or else there is surely nothing to search.
                
            End If
        End If
                  
        If (.flags And FR_DOWN) = FR_DOWN Then  'If Down option button is selected....
            FndInfo.SearchDown = True
        Else                                    'Else if Up option button is selected..
            FndInfo.SearchDown = False
        End If
        
        If (.flags And FR_MATCHCASE) = FR_MATCHCASE Then    'If Match Case is selected.....
            FndInfo.MatchCase = True
        Else                                                'Else if not selected....
            FndInfo.MatchCase = False
        End If
                  
        If (.flags And FR_FINDNEXT) = FR_FINDNEXT Then      'If FindNext button is clicked,
            FindNextClickedOnce = True                      'then set this variable to true
            FindTextInEdit = FindNext(hEdit)                'and go to the main search function.
        End If
    
    End With
    
End Function


'This function performs the main search operation. It is called directly when
'FindNext menu is clicked and from the FindTextInEdit function.

Public Function FindNext(ByVal hEdit As Long) As Long
    
    Dim ret         As Long
    Dim ret2        As Long
    Dim CaretPos    As Long         'This indicated the position of the caret.
    Dim CurLine     As Long
    Dim temp        As Long
    Dim TextSrc     As String
    Dim i           As Long
    
    
    'The following operation is done to see if Find dialog was previously created
    'by Find menu. If not, then if FindNext button was never clicked, create the
    'dialog window, else just exit the function.
    
    If StrComp(SearchText, String$(MaxSearchLen, Chr$(0)), vbBinaryCompare) = 0 Then
        If FindNextClickedOnce = False Then
            SearchForText
        End If
        FindNext = 0
        Exit Function
    End If
      
      
    ret = SendMessageByNum(hEdit, WM_GETTEXTLENGTH, 0&, 0&)     'Gets the length of the text without null-character.
    TextSrc = String$(ret + 1, Chr$(0))                         'Initialize the string with the null-character.
    Call SendMessageByString(hEdit, WM_GETTEXT, ret + 1, TextSrc)   'Gets the text in Edit box. Length of string plus the
                                                                    'null char is provided in wParam and the string in lParam.
    
    Call SendMessageNoByVal(hEdit, EM_GETSEL, ret, ret2)         'Gets the starting cursor position of selection in ret and
                                                                'cursor position after the ending of selection in ret2.
    
    With FndInfo
    
        If .SearchDown Then             'If search is downwards, then position of
            CaretPos = ret2             'caret should be the position just after the
                                        'end of current selection,
        Else                            'else if search is upwards, then position of
            CaretPos = ret - 1          'caret should be the position of the character
        End If                          'just before starting position of current selection.
        
        
        
        'If starting of selection was the first character and search was upwards,
        'then CaretPos will be -1. In that case, find the parent window and display
        'the 'not found' message.
        
        If CaretPos = -1 Then
            Call MessageBox(IIf(gFindDlgHandle, gFindDlgHandle, GetParent(hEdit)), "Not found '" & SearchText & "'", "Not found", MB_ICONINFORMATION Or MB_OK)
            FindNext = 0
            Exit Function
        End If
        
        
        If .SearchDown Then                 'If search is downwards.....
            If .MatchCase Then
                temp = InStr(CaretPos + 1, TextSrc, SearchText, vbBinaryCompare)
                                            'If Match Case is selected, then search with
                                            'vbBinaryCompare option. InStr is 1 based, but
                                            'CaretPos is 0 based. So adding 1 to CaretPos.
            Else
                temp = InStr(CaretPos + 1, TextSrc, SearchText, vbTextCompare)
                                            'If Match Case is not selected, then search with
                                            'vbTextCompare option.
            End If
        Else
            If .MatchCase Then
                temp = InStrRev(TextSrc, SearchText, CaretPos + 1, vbBinaryCompare)
            Else
                temp = InStrRev(TextSrc, SearchText, CaretPos + 1, vbTextCompare)
            End If
        End If
            
        If temp = 0 Then
            Call MessageBox(IIf(gFindDlgHandle, gFindDlgHandle, GetParent(hEdit)), "Not found '" & SearchText & "'", "Not found", MB_ICONINFORMATION Or MB_OK)
        Else
            Call SendMessageByNum(hEdit, EM_SETSEL, temp - 1, temp - 1 + Len(SearchText))
            Call SendMessageByNum(hEdit, EM_SCROLLCARET, 0&, 0&)
        End If
        
        FindNext = 0
        
    End With

End Function

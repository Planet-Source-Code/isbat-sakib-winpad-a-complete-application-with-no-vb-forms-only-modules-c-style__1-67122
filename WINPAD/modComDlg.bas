Attribute VB_Name = "modComdlg"
'******************************************************************************
'******This code is downloaded from Planet Source Code
'****** Author: Isbat Sakib
'****** Email: iamsakib@gmail.com
'******************************************************************************

'This module (modComdlg) has been developed with the help of many such common-dialog
'control related submissions from Planet Source Code.



Option Explicit

Private Type OPENFILENAME
    lStructSize                     As Long
    hwndOwner                       As Long
    hInstance                       As Long
    lpstrFilter                     As String
    lpstrCustomFilter               As String
    nMaxCustFilter                  As Long
    nFilterIndex                    As Long
    lpstrFile                       As String
    nMaxFile                        As Long
    lpstrFileTitle                  As String
    nMaxFileTitle                   As Long
    lpstrInitialDir                 As String
    lpstrTitle                      As String
    flags                           As Long
    nFileOffset                     As Integer
    nFileExtension                  As Integer
    lpstrDefExt                     As String
    lCustData                       As Long
    lpfnHook                        As Long
    lpTemplateName                  As String
End Type

Private Type ChooseFont
    lStructSize                     As Long
    hwndOwner                       As Long
    hdc                             As Long
    lpLogFont                       As Long
    iPointSize                      As Long
    flags                           As Long
    rgbColors                       As Long
    lCustData                       As Long
    lpfnHook                        As Long
    lpTemplateName                  As String
    hInstance                       As Long
    lpszStyle                       As String
    nFontType                       As Integer
    MISSING_ALIGNMENT               As Integer
    nSizeMin                        As Long
    nSizeMax                        As Long
End Type

Public Type PrintDlg
    lStructSize                     As Long
    hwndOwner                       As Long
    hDevMode                        As Long
    hDevNames                       As Long
    hdc                             As Long
    flags                           As Long
    nFromPage                       As Integer
    nToPage                         As Integer
    nMinPage                        As Integer
    nMaxPage                        As Integer
    nCopies                         As Integer
    hInstance                       As Long
    lCustData                       As Long
    lpfnPrintHook                   As Long
    lpfnSetupHook                   As Long
    lpPrintTemplateName             As String
    lpSetupTemplateName             As String
    hPrintTemplate                  As Long
    hSetupTemplate                  As Long
End Type

Public Type PageSetupDlg
    lStructSize                     As Long
    hwndOwner                       As Long
    hDevMode                        As Long
    hDevNames                       As Long
    flags                           As Long
    ptPaperSize                     As POINTAPI
    rtMinMargin                     As RECT
    rtMargin                        As RECT
    hInstance                       As Long
    lCustData                       As Long
    lpfnPageSetupHook               As Long
    lpfnPagePaintHook               As Long
    lpPageSetupTemplateName         As String
    hPageSetupTemplate              As Long
End Type


Private Const MAX_PATH              As Long = 2048
Private Const MAX_FILE              As Long = 2048

Private Const OFN_FILEMUSTEXIST     As Long = &H1000
Private Const OFN_HIDEREADONLY      As Long = &H4
Private Const OFN_NOCHANGEDIR       As Long = &H8
Private Const OFN_NOREADONLYRETURN  As Long = &H8000&
Private Const OFN_OVERWRITEPROMPT   As Long = &H2
Private Const OFN_PATHMUSTEXIST     As Long = &H800
Private Const CF_EFFECTS            As Long = &H100&
Private Const CF_FORCEFONTEXIST     As Long = &H10000
Private Const CF_INITTOLOGFONTSTRUCT As Long = &H40&
Private Const CF_SCREENFONTS        As Long = &H1

Private Const PD_ALLPAGES           As Long = &H0
Private Const PD_NOPAGENUMS         As Long = &H8
Private Const PD_NOSELECTION        As Long = &H4
Private Const PD_HIDEPRINTTOFILE    As Long = &H100000
Private Const PD_RETURNDC           As Long = &H100
Private Const PD_RETURNDEFAULT      As Long = &H400
Private Const PD_USEDEVMODECOPIESANDCOLLATE As Long = &H40000

Private Const PSD_DISABLEORIENTATION    As Long = &H100
Private Const PSD_INTHOUSANDTHSOFINCHES As Long = &H4
Private Const PSD_MARGINS           As Long = &H2
Private Const PSD_RETURNDEFAULT     As Long = &H400



Public Enum EnumOpenSaveFlags
    FileMustExist = OFN_FILEMUSTEXIST
    HideReadOnly = OFN_HIDEREADONLY
    NoChangeDir = OFN_NOCHANGEDIR
    NoReadOnlyReturn = OFN_NOREADONLYRETURN
    OverWritePrompt = OFN_OVERWRITEPROMPT
    PathMustExist = OFN_PATHMUSTEXIST
End Enum

Public Enum EnumFontFlags
    ForceFontExist = CF_FORCEFONTEXIST
    ShowScreenFontsOnly = CF_SCREENFONTS
    UseEffects = CF_EFFECTS
    UseLogFontStructure = CF_INITTOLOGFONTSTRUCT
End Enum

Public Enum EnumPrintFlags
    HideSelection = PD_NOSELECTION
    HidePrintToFile = PD_HIDEPRINTTOFILE
    UseDevModeCopiesAndCollate = PD_USEDEVMODECOPIESANDCOLLATE
    ReturnDC = PD_RETURNDC
    ReturnDefault = PD_RETURNDEFAULT
End Enum

Public ghDevMode        As Long
Public ghDevNames       As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Public Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PrintDlg) As Long
Public Declare Function PageSetupDlg Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PageSetupDlg) As Long


Public Function ShowOpen(Optional InitDir As String, _
                        Optional Filter As String = "All (*.*)| *.*", _
                        Optional DefaultExt As String = "txt", _
                        Optional FileName As String, _
                        Optional DialogTitle As String = "Open", _
                        Optional ShowOpenFlags As EnumOpenSaveFlags = 0, _
                        Optional Owner As Long = -1) As String
    
    Dim Ofn                         As OPENFILENAME
    Dim TempString                  As String
    Dim ch                          As String
    Dim i                           As Long
    Dim ret                         As Long
    
    With Ofn
    
        .lStructSize = Len(Ofn)
        
        If Owner <> -1 Then
            .hwndOwner = Owner
        End If
        
        
        For i = 1 To Len(Filter)
            ch = Mid$(Filter, i, 1)
            If ch = "|" Then
                 TempString = TempString & Chr$(0)
            Else
                 TempString = TempString & ch
            End If
        Next i
    
        
        TempString = TempString & Chr$(0) & Chr$(0)
        .lpstrFilter = TempString
        .nFilterIndex = 1&
        
        
        TempString = FileName & String$(MAX_PATH - Len(FileName), Chr$(0))
        .lpstrFile = TempString
        .nMaxFile = MAX_PATH
        
        
        .lpstrFileTitle = vbNullString
        .nMaxFileTitle = 0&
        
        .lpstrInitialDir = InitDir
        .lpstrTitle = DialogTitle
        .flags = ShowOpenFlags
        .lpstrDefExt = DefaultExt
        
        ret = GetOpenFileName(Ofn)
        
        If ret <> 0 Then
            
            ret = InStr(.lpstrFile, Chr$(0) & Chr$(0))
             
            If ret > 1 Then
                ShowOpen = Left$(.lpstrFile, ret - 1)
            End If
            
            ret = InStr(ShowOpen, Chr$(0))
            If ret > 0 Then
                ShowOpen = IIf(ret = 1, vbNullString, Left$(ShowOpen, ret - 1))
            End If
             
        Else
            ShowOpen = vbNullString
            If CommDlgExtendedError() Then
                Call MessageBox(Owner, "Error occured while opening the file.", "Error", MB_OK Or MB_ICONSTOP)
            End If
        End If
    End With
    
End Function

Public Function ShowSave(Optional InitDir As String, _
                        Optional Filter As String = "All (*.*)| *.*", _
                        Optional DefaultExt As String = "txt", _
                        Optional FileName As String, _
                        Optional DialogTitle As String = "Save As", _
                        Optional ShowSaveFlags As EnumOpenSaveFlags = 0, _
                        Optional Owner As Long = -1) As String

    Dim Ofn                         As OPENFILENAME
    Dim TempString                  As String
    Dim ch                          As String
    Dim i                           As Long
    Dim ret                         As Long
    
    With Ofn
    
        .lStructSize = Len(Ofn)
        
        If Owner <> -1 Then
            .hwndOwner = Owner
        End If
        
        
        For i = 1 To Len(Filter)
            ch = Mid$(Filter, i, 1)
            If ch = "|" Then
                 TempString = TempString & Chr$(0)
            Else
                 TempString = TempString & ch
            End If
        Next i
        
        
        TempString = TempString & Chr$(0) & Chr$(0)
        .lpstrFilter = TempString
        .nFilterIndex = 1&
        
        
        TempString = FileName & String$(MAX_PATH - Len(FileName), Chr$(0))
        .lpstrFile = TempString
        .nMaxFile = MAX_PATH
        
        .lpstrFileTitle = vbNullString
        .nMaxFileTitle = 0&
        
        .lpstrInitialDir = InitDir
        .lpstrTitle = DialogTitle
        .flags = ShowSaveFlags
        .lpstrDefExt = DefaultExt
        
        ret = GetSaveFileName(Ofn)
        
        If ret <> 0 Then
            
            ret = InStr(.lpstrFile, Chr$(0) & Chr$(0))
             
            If ret > 1 Then
                ShowSave = Left$(.lpstrFile, ret - 1)
            End If
            
            ret = InStr(ShowSave, Chr$(0))
            If ret > 0 Then
                ShowSave = IIf(ret = 1, vbNullString, Left$(ShowSave, ret - 1))
            End If
             
        Else
            ShowSave = vbNullString
            If CommDlgExtendedError() Then
                Call MessageBox(Owner, "Error occured while saving the file.", "Error", MB_OK Or MB_ICONSTOP)
            End If
        End If
        
    End With

End Function

Public Function ShowFont(Optional DialogTitle As String = "Font", _
                    Optional ShowFontFlags As EnumFontFlags = 0, _
                    Optional AddressOfLOGFONTStruct As Long = 0, _
                    Optional RGBColor As Long = 0, _
                    Optional Owner As Long = -1) As Long


    Dim Cf                          As ChooseFont
    Dim ret                         As Long
    
    Const PointsPerTwip             As Long = 1440 / 72
    
    With Cf
        .lStructSize = Len(Cf)
        .hwndOwner = Owner
        .hdc = 0&
        .lpLogFont = AddressOfLOGFONTStruct
        .iPointSize = 0&
        .flags = ShowFontFlags
        .rgbColors = RGBColor
        .lCustData = 0&
        .lpfnHook = 0&
        .lpTemplateName = vbNullString
        .hInstance = 0&
        .lpszStyle = vbNullString
        .nFontType = 0&
        .MISSING_ALIGNMENT = 0&
        .nSizeMin = 0&
        .nSizeMax = 0&
    End With
    
    ret = ChooseFont(Cf)
    
    If ret <> 0 Then
        RGBColor = Cf.rgbColors
    Else
        If CommDlgExtendedError() Then
            Call MessageBox(Owner, "Error occured while font processing.", "Error", MB_OK Or MB_ICONSTOP)
        End If
    End If
    
    ShowFont = ret
    
End Function

Public Function ShowPageSetup(Owner As Long) As Long

    Dim Psd         As PageSetupDlg
    Dim ret         As Long
    
    With Psd
        .lStructSize = Len(Psd)
        .hwndOwner = gHwnd
        .hDevMode = ghDevMode
        .hDevNames = ghDevNames
        .flags = PSD_MARGINS Or PSD_INTHOUSANDTHSOFINCHES
        .rtMargin.Top = IIf(gMargins.Top = 0, 1000, gMargins.Top)
        .rtMargin.Bottom = IIf(gMargins.Bottom = 0, 1000, gMargins.Bottom)
        .rtMargin.Left = IIf(gMargins.Left = 0, 750, gMargins.Left)
        .rtMargin.Right = IIf(gMargins.Right = 0, 750, gMargins.Right)
        .hInstance = App.hInstance
        .lCustData = 0&
        .lpfnPageSetupHook = 0&
        .lpfnPagePaintHook = 0&
        .lpPageSetupTemplateName = vbNullString
        .hPageSetupTemplate = 0&
    End With
    
    ShowPageSetup = PageSetupDlg(Psd)
    
    If ShowPageSetup = 0 Then
        If CommDlgExtendedError() Then
            Call MessageBox(Owner, "Error occured while setting up the page.", "Error", MB_OK Or MB_ICONSTOP)
        End If
        Exit Function
    End If
    
    ghDevMode = Psd.hDevMode
    ghDevNames = Psd.hDevNames
    gMargins = Psd.rtMargin
    
End Function

Public Function ShowPrint(Optional Owner As Long = -1, Optional PrintFlags As EnumPrintFlags = 0, _
                          Optional PrinterDC As Long = 0, Optional DMode As Long = 0, _
                          Optional DNames As Long = 0) As Long

    Dim Pd          As PrintDlg
    Dim ret         As Long
    
    With Pd
        .lStructSize = Len(Pd)
        .hwndOwner = Owner
        .hDevMode = ghDevMode
        .hDevNames = ghDevNames
        .hdc = 0&
        .flags = PrintFlags
        .nFromPage = 1
        .nToPage = 1
        .nMinPage = 1
        .nMaxPage = 1
        .nCopies = 1
        .hInstance = App.hInstance
        .lCustData = 0&
        .lpfnPrintHook = 0&
        .lpfnSetupHook = 0&
        .lpPrintTemplateName = vbNullString
        .lpSetupTemplateName = vbNullString
        .hPrintTemplate = 0&
        .hSetupTemplate = 0&
    End With
    
    ShowPrint = PrintDlg(Pd)
    
    If ShowPrint = 0 Then
        If CommDlgExtendedError() Then
            Call MessageBox(Owner, "Error occured in the printing mechanism.", "Error", MB_OK Or MB_ICONSTOP)
        End If
        Exit Function
    End If
    
    DMode = Pd.hDevMode
    DNames = Pd.hDevNames
    PrinterDC = Pd.hdc
    
End Function

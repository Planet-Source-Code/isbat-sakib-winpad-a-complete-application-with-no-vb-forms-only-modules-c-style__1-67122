Attribute VB_Name = "modPrint"
'******************************************************************************
'******This code is downloaded from Planet Source Code
'****** Author: Isbat Sakib
'****** Email: iamsakib@gmail.com
'******************************************************************************

'This is the part where I decided to discontinue with its development. Maybe you could
'continue it?


Option Explicit

Public Type DEVMODE
    dmDeviceName                    As String * 32
    dmSpecVersion                   As Integer
    dmDriverVersion                 As Integer
    dmSize                          As Integer
    dmDriverExtra                   As Integer
    dmFields                        As Long
    dmOrientation                   As Integer
    dmPaperSize                     As Integer
    dmPaperLength                   As Integer
    dmPaperWidth                    As Integer
    dmScale                         As Integer
    dmCopies                        As Integer
    dmDefaultSource                 As Integer
    dmPrintQuality                  As Integer
    dmColor                         As Integer
    dmDuplex                        As Integer
    dmYResolution                   As Integer
    dmTTOption                      As Integer
    dmCollate                       As Integer
    dmFormName                      As String * 32
    dmLogPixels                     As Integer
    dmBitsPerPel                    As Long
    dmPelsWidth                     As Long
    dmPelsHeight                    As Long
    dmDisplayFlags                  As Long
    dmDisplayFrequency              As Long
    dmICMMethod                     As Long
    dmICMIntent                     As Long
    dmMediaType                     As Long
    dmDitherType                    As Long
    dmReserved1                     As Long
    dmReserved2                     As Long
    
    'If you use Win2000 or WinXP then you can uncomment these, though it will
    'also work without them anyway.
    'dmPanningWidth                 As Long
    'dmPanningHeight                As Long

End Type

Public Type DEVNAMES
    wDriverOffset                   As Integer
    wDeviceOffset                   As Integer
    wOutputOffset                   As Integer
    wDefault                        As Integer
End Type

Private Type DOCINFO
    cbSize                          As Long
    lpszDocName                     As String
    lpszOutput                      As String
    lpszDatatype                    As String
    fwType                          As Long
End Type

Private Type SIZE
    cx                              As Long
    cy                              As Long
End Type

Private Const VERTRES = 10
Private Const HORZRES = 8
Private Const HORZSIZE = 4
Private Const VERTSIZE = 6

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Private Const PHYSICALHEIGHT = 111
Private Const PHYSICALWIDTH = 110
Private Const PHYSICALOFFSETX = 112
Private Const PHYSICALOFFSETY = 113

Public gMargins         As RECT

Private Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As Long, lpdi As DOCINFO) As Long
Private Declare Function StartPage Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPage Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function EndDoc Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const MM_HIENGLISH = 5
Private Const MM_TWIPS As Long = 6
Private Const MM_LOENGLISH As Long = 4
Private Const MM_TEXT As Long = 1
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetCharWidth Lib "gdi32.dll" Alias "GetCharWidthA" (ByVal hdc As Long, ByVal wFirstChar As Long, ByVal wLastChar As Long, ByRef lpBuffer As Long) As Long
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Public Const DT_CALCRECT = &H400
Public Const DT_LEFT = &H0
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_TOP = &H0
Public Const DT_WORDBREAK = &H10
Public Const DT_EDITCONTROL As Long = &H2000
Private Const DT_EXPANDTABS As Long = &H40

Public Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Public Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Public Const TA_LEFT = 0
Public Const TA_TOP = 0
Public Const TA_NOUPDATECP = 0
Public Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type

Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Public Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long



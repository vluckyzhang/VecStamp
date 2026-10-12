Attribute VB_Name = "VecStamp"
' =====================================================================
'  VecStamp - export selected PowerPoint shapes as vector graphics
'  Chinese name: Shi Yin ("vector stamp"). Project page and docs: README.md
'
'  Copyright (c) 2026 vluckyzhang <vluckyzhang@gmail.com>
'  Released under the MIT License. See the LICENSE file for details.
'
'  Notes for developers
'  - This is the readable UTF-8 source. build/build_bas.py turns it into
'    dist/VecStamp.bas: every Chinese string literal becomes U("hex") so the
'    module imports correctly under any Windows locale, and the generated
'    block at the end is filled with the VSDX template parts from src/vsdx
'    and the XPS -> SVG converter from src/helpers/xps2svg.ps1.
'  - All PowerPoint objects are late bound (As Object) and enum values are
'    numeric constants, so the module only depends on the VBA library and
'    compiles the same way on Windows and macOS.
'  - Ribbon callbacks are prefixed VS_ (see src/customUI14.xml).
'  - macOS: dialogs, moving files out of the PowerPoint sandbox, gzip, zip, the
'    colour picker, the progress dialog and the clipboard are done by the
'    AppleScript helper VecStamp.scpt that
'    the macOS installer puts in ~/Library/Application Scripts/com.microsoft.Powerpoint/.
' =====================================================================
Option Explicit

#If Mac Then
' (no Windows API declarations on macOS)
#Else
#If VBA7 Then
Private Type OPENFILENAMEW
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As LongPtr
    lpstrCustomFilter As LongPtr
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As LongPtr
    nMaxFile As Long
    lpstrFileTitle As LongPtr
    nMaxFileTitle As Long
    lpstrInitialDir As LongPtr
    lpstrTitle As LongPtr
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As LongPtr
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As LongPtr
    pvReserved As LongPtr
    dwReserved As Long
    FlagsEx As Long
End Type
Private Type CHOOSECOLORW
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    rgbResult As Long
    lpCustColors As LongPtr
    Flags As Long
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As LongPtr
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type INITCOMMONCONTROLSEX_T
    dwSize As Long
    dwICC As Long
End Type
Private Declare PtrSafe Function GetSaveFileNameW Lib "comdlg32.dll" (ByRef ofn As OPENFILENAMEW) As Long
Private Declare PtrSafe Function ChooseColorW Lib "comdlg32.dll" (ByRef cc As CHOOSECOLORW) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, _
    ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, _
    ByVal lpParam As LongPtr) As LongPtr
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SetWindowTextW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Function SendMessageW Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function CreateFontW Lib "gdi32" (ByVal nHeight As Long, ByVal nWidth As Long, _
    ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, _
    ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, _
    ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function InitCommonControlsEx Lib "comctl32" (ByRef icc As INITCOMMONCONTROLSEX_T) As Long
Private Type STARTUPINFOW
    cb As Long
    lpReserved As LongPtr
    lpDesktop As LongPtr
    lpTitle As LongPtr
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As LongPtr
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type
Private Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type
Private Declare PtrSafe Function CreateProcessW Lib "kernel32" (ByVal lpApplicationName As LongPtr, ByVal lpCommandLine As LongPtr, _
    ByVal lpProcessAttributes As LongPtr, ByVal lpThreadAttributes As LongPtr, ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, ByVal lpEnvironment As LongPtr, ByVal lpCurrentDirectory As LongPtr, _
    ByRef lpStartupInfo As STARTUPINFOW, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByRef lpExitCode As Long) As Long
Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
' colour panel (a native window driven by a message loop in ColorPanel)
Private Type MSG_T
    hwnd As LongPtr
    message As Long
    wParam As LongPtr
    lParam As LongPtr
    msgTime As Long
    ptX As Long
    ptY As Long
End Type
Private Declare PtrSafe Function PeekMessageW Lib "user32" (ByRef lpMsg As MSG_T, ByVal hwnd As LongPtr, _
    ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare PtrSafe Function TranslateMessage Lib "user32" (ByRef lpMsg As MSG_T) As Long
Private Declare PtrSafe Function DispatchMessageW Lib "user32" (ByRef lpMsg As MSG_T) As LongPtr
Private Declare PtrSafe Function IsDialogMessageW Lib "user32" (ByVal hDlg As LongPtr, ByRef lpMsg As MSG_T) As Long
Private Declare PtrSafe Function EnableWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal bEnable As Long) As Long
Private Declare PtrSafe Function SetFocusW Lib "user32" Alias "SetFocus" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function GetWindowTextW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpString As LongPtr, _
    ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function LoadImageW Lib "user32" (ByVal hInst As LongPtr, ByVal lpszName As LongPtr, _
    ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As LongPtr
Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDoubleClickTime Lib "user32" () As Long
Private Declare PtrSafe Function PostQuitMessage Lib "user32" (ByVal nExitCode As Long) As Long
Private Declare PtrSafe Sub SleepX Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function AdjustWindowRectEx Lib "user32" (ByRef lpRect As RECT, ByVal dwStyle As Long, _
    ByVal bMenu As Long, ByVal dwExStyle As Long) As Long
Private gPgWnd As LongPtr, gPgMsg As LongPtr, gPgBar As LongPtr, gPgSub As LongPtr, gPgFont As LongPtr, gPgFontB As LongPtr
Private gPgPct As LongPtr
#Else
Private Type OPENFILENAMEW
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As Long
    lpstrCustomFilter As Long
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As Long
    nMaxFile As Long
    lpstrFileTitle As Long
    nMaxFileTitle As Long
    lpstrInitialDir As Long
    lpstrTitle As Long
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
    pvReserved As Long
    dwReserved As Long
    FlagsEx As Long
End Type
Private Type CHOOSECOLORW
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type INITCOMMONCONTROLSEX_T
    dwSize As Long
    dwICC As Long
End Type
Private Declare Function GetSaveFileNameW Lib "comdlg32.dll" (ByRef ofn As OPENFILENAMEW) As Long
Private Declare Function ChooseColorW Lib "comdlg32.dll" (ByRef cc As CHOOSECOLORW) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, _
    ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, _
    ByVal lpParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFontW Lib "gdi32" (ByVal nHeight As Long, ByVal nWidth As Long, _
    ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, _
    ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, _
    ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32" (ByRef icc As INITCOMMONCONTROLSEX_T) As Long
Private Type STARTUPINFOW
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Private Declare Function CreateProcessW Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, _
    ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
    ByRef lpStartupInfo As STARTUPINFOW, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Type MSG_T
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    msgTime As Long
    ptX As Long
    ptY As Long
End Type
Private Declare Function PeekMessageW Lib "user32" (ByRef lpMsg As MSG_T, ByVal hwnd As Long, _
    ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (ByRef lpMsg As MSG_T) As Long
Private Declare Function DispatchMessageW Lib "user32" (ByRef lpMsg As MSG_T) As Long
Private Declare Function IsDialogMessageW Lib "user32" (ByVal hDlg As Long, ByRef lpMsg As MSG_T) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal bEnable As Long) As Long
Private Declare Function SetFocusW Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long, _
    ByVal nMaxCount As Long) As Long
Private Declare Function LoadImageW Lib "user32" (ByVal hInst As Long, ByVal lpszName As Long, _
    ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDoubleClickTime Lib "user32" () As Long
Private Declare Function PostQuitMessage Lib "user32" (ByVal nExitCode As Long) As Long
Private Declare Sub SleepX Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private Declare Function AdjustWindowRectEx Lib "user32" (ByRef lpRect As RECT, ByVal dwStyle As Long, _
    ByVal bMenu As Long, ByVal dwExStyle As Long) As Long
Private gPgWnd As Long, gPgMsg As Long, gPgBar As Long, gPgSub As Long, gPgFont As Long, gPgFontB As Long
Private gPgPct As Long
#End If
#End If

' ---- product information
Private Const VS_NAME As String = "VecStamp"
Private Const VS_VERSION As String = "1.2.2"
Private Const VS_AUTHOR As String = "vluckyzhang"
Private Const VS_EMAIL As String = "vluckyzhang@gmail.com"
Private Const VS_YEAR As String = "2026"
Private Const VS_REPO As String = "https://github.com/vluckyzhang/VecStamp"
Private Const REG_ROOT As String = "HKCU\Software\VecStamp\"
Private Const MAC_SCRIPT As String = "VecStamp.scpt"
Private Const MAC_SEP As String = "|*|"

' ---- PowerPoint / Office enum values (numeric: late binding, no type-library dependency)
Private Const PP_FMT_PNG As Long = 2
Private Const PP_FMT_WMF As Long = 4
Private Const PP_FMT_EMF As Long = 5
Private Const PP_FMT_SVG As Long = 6
Private Const PP_RELATIVE_TO_SLIDE As Long = 1
Private Const PP_SEL_SLIDES As Long = 1
Private Const PP_SEL_SHAPES As Long = 2
Private Const PP_SEL_TEXT As Long = 3
Private Const PP_PASTE_EMF As Long = 2
Private Const PP_LAYOUT_BLANK As Long = 12
Private Const PP_SAVE_AS_PDF As Long = 32
Private Const MSO_TRUE As Long = -1
Private Const MSO_FALSE As Long = 0
Private Const MSO_RECTANGLE As Long = 1
Private Const MSO_AUTOSHAPE As Long = 1
Private Const MSO_CALLOUT As Long = 2
Private Const MSO_FREEFORM As Long = 5
Private Const MSO_GROUP As Long = 6
Private Const MSO_LINE As Long = 9
Private Const MSO_TEXTEFFECT As Long = 15
Private Const MSO_TEXTBOX As Long = 17
Private Const KAPPA As Double = 0.552284749830794
Private Const PI2 As Double = 6.28318530717959
Private Const PDEG As Double = 2.90888208665722E-07     ' radians per DrawingML angle unit (1/60000 degree)
Private Const MSO_SEND_TO_BACK As Long = 1
Private Const MSO_PLACEHOLDER As Long = 14
Private Const MSO_FOLDER_PICKER As Long = 4

' ---- our format index (= order in the ribbon drop-down and the save dialog)
Private Const FMT_EMF As Long = 0
Private Const FMT_EMZ As Long = 1
Private Const FMT_SVG As Long = 2
Private Const FMT_WMF As Long = 3
Private Const FMT_PNG As Long = 4
Private Const FMT_PDF As Long = 5
Private Const FMT_VSDX As Long = 6
Private Const FMT_LAST As Long = 6

' ---- backgrounds (= order in the ribbon drop-down)
Private Const BG_NONE As Long = 0
Private Const BG_WHITE As Long = 1
Private Const BG_BEIGE As Long = 2
Private Const BG_GRID As Long = 3
Private Const BG_CUSTOM As Long = 4
Private Const DEF_CUSTOM_COLOR As String = "255,248,235"

Private Const DPI_MAX As Long = 3000
Private Const SVG_TIMEOUT As Long = 180          ' seconds for the built-in XPS -> SVG conversion
Private Const PNG_WARN_PIXELS As Double = 120000000#

#If Mac Then
#Else
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_NOCHANGEDIR As Long = &H8
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_EXPLORER As Long = &H80000
Private Const CC_RGBINIT As Long = &H1
Private Const CC_FULLOPEN As Long = &H2
Private Const CC_ANYCOLOR As Long = &H100
Private Const WS_POPUP As Long = &H80000000
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const SS_NOPREFIX As Long = &H80
Private Const WM_SETFONT As Long = &H30
Private Const PBM_SETPOS As Long = &H402
Private Const PBM_SETRANGE32 As Long = &H406
Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNOACTIVATE As Long = 4
Private Const LOGPIXELSY As Long = 90
Private Const VK_ESCAPE As Long = &H1B
Private Const ICC_PROGRESS_CLASS As Long = &H20
Private Const SS_RIGHT As Long = &H2
Private Const CREATE_NO_WINDOW As Long = &H8000000
Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const WAIT_TIMEOUT As Long = &H102
Private Const PM_REMOVE As Long = 1
Private Const WM_QUIT As Long = &H12
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const VK_RETURN As Long = &HD
Private Const VK_SPACE As Long = &H20
Private Const WS_TABSTOP As Long = &H10000
Private Const WS_BORDER As Long = &H800000
Private Const WS_EX_CONTROLPARENT As Long = &H10000
Private Const SS_NOTIFY As Long = &H100
Private Const SS_BITMAP As Long = &HE
Private Const STM_SETIMAGE As Long = &H172
Private Const IMAGE_BITMAP As Long = 0
Private Const LR_LOADFROMFILE As Long = &H10
Private Const BS_DEFPUSHBUTTON As Long = &H1
Private Const ES_AUTOHSCROLL As Long = &H80
Private Const EM_LIMITTEXT As Long = &HC5
Private Const EM_SETSEL As Long = &HB1
Private Const WS_EX_CLIENTEDGE As Long = &H200
Private Const SW_SHOW As Long = 5
Private Const COLOR_BTNFACE As Long = 15
#End If

' ---- results of RunHidden besides the program's own exit code
Private Const RUN_FAILED As Long = -1
Private Const RUN_CANCELLED As Long = -2
Private Const RUN_TIMEOUT As Long = -3
Private Const ERR_CANCEL As Long = vbObjectError + 599

Private gRibbon As Object
Private gLang As String             ' interface language (zh-CN, zh-TW, en, ja)
Private gStr As Collection          ' key -> text in gLang
Private gSvgFallback As Long        ' per run: 0 = not asked yet, 1 = save EMF instead, 2 = skip
Private gSvgNative As Long          ' per run: 0 = untested, 1 = PowerPoint writes valid SVG, 2 = it does not
Private gSvgXpsOff As Boolean       ' per run: the XPS -> SVG conversion timed out, use the other methods
Private gBigPngOk As Boolean        ' per run: user accepted a very large PNG
Private gCustColors(0 To 15) As Long
Private gBusy As Boolean            ' an export is running (ribbon clicks during DoEvents are ignored)
' progress window: items (shapes, files or slides) and the position inside the current item
Private gPgActive As Boolean, gPgItems As Long, gPgItem As Long, gPgFrac As Double, gPgTitle As String
Private gPgLabel As String, gPgDetail As String, gPgCancel As Boolean, gPgFile As String
#If Mac Then
Private gPgMacPid As String, gPgMacPct As Long, gPgMacTime As Double, gPgMacText As String, gPgMacCheck As Double
#End If
' native VSDX progress
Private gVxTotal As Long, gVxDone As Long
' native VSDX writer
Private gVxSb() As String, gVxSbN As Long, gVxId As Long, gVxMedia As Collection
Private gVxScratch As Object, gVxPres As Object, gVxNative As Long, gVxPictures As Long, gVxCopy As Long
Private gGeo As String, gGeoSec As String, gGeoIx As Long, gGeoRow As Long
Private gGeoW As Double, gGeoH As Double, gGeoFlipH As Boolean, gGeoFlipV As Boolean
Private gGeoX0 As Double, gGeoY0 As Double, gGeoLX As Double, gGeoLY As Double
Private gGeoNoFill As Boolean, gGeoNoLine As Boolean, gGeoOpenOnly As Boolean, gGeoTextBox As Boolean
Private gGeoCU As Double, gGeoCV As Double, gGeoSU As Double, gGeoSV As Double, gGeoClosed As Boolean
' preset geometry engine: guide values, adjust values (DrawingML units), shape size
Private gPv() As Double, gPa() As Double, gPw As Double, gPh As Double


' =====================================================================
'  Public macros (can be bound to the Quick Access Toolbar or run by name)
' =====================================================================
Public Sub VS_ExportDefault()
    DoExport GetIdx("Format", FMT_LAST)
End Sub

Public Sub VS_ExportEMF()
    DoExport FMT_EMF
End Sub

Public Sub VS_ExportEMZ()
    DoExport FMT_EMZ
End Sub

Public Sub VS_ExportSVG()
    DoExport FMT_SVG
End Sub

Public Sub VS_ExportWMF()
    DoExport FMT_WMF
End Sub

Public Sub VS_ExportPNG()
    DoExport FMT_PNG
End Sub

Public Sub VS_ExportPDF()
    DoExport FMT_PDF
End Sub

Public Sub VS_ExportVSDX()
    DoExport FMT_VSDX
End Sub

Public Sub VS_ExportSlides()
    DoBatch
End Sub

Public Sub VS_PictureCopy()
    DoPictureCopy
End Sub

Public Sub VS_PickColor()
    DoPickColor
End Sub

Public Sub VS_MoreColors()
    DoMoreColors
End Sub

Public Sub VS_Eyedropper()
    DoEyedropper
End Sub

Public Sub VS_Feedback()
    DoFeedback
End Sub


' =====================================================================
'  Ribbon callbacks
'  Every text comes from the string tables ("rb.<control id>" = label, ".tip" = screentip,
'  ".sup" = supertip, ".desc" = description), so the whole tab switches language at once when
'  the ribbon is invalidated (VS_OnCheck on a language entry).
' =====================================================================
Public Sub VS_OnLoad(ribbon As Object)
    Set gRibbon = ribbon
End Sub

Public Sub VS_OnExport(control As Object)
    Select Case control.Id
        Case "vsBtnEMF", "vsMnuEMF": DoExport FMT_EMF
        Case "vsBtnEMZ", "vsMnuEMZ": DoExport FMT_EMZ
        Case "vsBtnSVG", "vsMnuSVG": DoExport FMT_SVG
        Case "vsBtnWMF", "vsMnuWMF": DoExport FMT_WMF
        Case "vsBtnPNG", "vsMnuPNG": DoExport FMT_PNG
        Case "vsBtnPDF", "vsMnuPDF": DoExport FMT_PDF
        Case "vsBtnVSDX", "vsMnuVSDX": DoExport FMT_VSDX
        Case Else: DoExport GetIdx("Format", FMT_LAST)
    End Select
End Sub

Public Sub VS_OnBatch(control As Object)
    DoBatch
End Sub

Public Sub VS_OnPicture(control As Object)
    DoPictureCopy
End Sub

Public Sub VS_OnPickColor(control As Object)
    DoPickColor
End Sub

Public Sub VS_OnMoreColors(control As Object)
    DoMoreColors
End Sub

Public Sub VS_OnEyedropper(control As Object)
    DoEyedropper
End Sub

' ---- texts
Public Sub VS_GetLabel(control As Object, ByRef label)
    Dim id As String
    id = control.Id
    Select Case id
        Case "vsBtnExport"
            label = Tx("rb.vsBtnExport", UCase$(FmtExt(GetIdx("Format", FMT_LAST))))
        Case "vsMnuBg"
            label = Tx("rb.vsMnuBg", BgShortName(BgTypeIdx()))
        Case "vsGalColor", "vsBgCustomMac"
            label = Tx("rb." & id)
            If BgTypeIdx() = BG_CUSTOM Then label = label & "  " & ChrW$(&H2713)
        Case "vsMnuLang"
            label = LangName(CurLang())
        Case "vsMnuEMF", "vsMnuEMZ", "vsMnuSVG", "vsMnuWMF", "vsMnuPDF", "vsMnuVSDX", "vsMnuPNG"
            label = Tx("fmt." & LCase$(Mid$(id, 6)))
        Case Else
            label = Tx("rb." & id)
    End Select
End Sub

Public Sub VS_GetScreentip(control As Object, ByRef tip)
    Dim id As String
    id = control.Id
    If Left$(id, 6) = "vsBgSw" Then id = "vsMnuBg"
    If TxHas("rb." & id & ".tip") Then tip = Tx("rb." & id & ".tip") Else VS_GetLabel control, tip
End Sub

Public Sub VS_GetSupertip(control As Object, ByRef tip)
    Dim id As String
    id = control.Id
    If Left$(id, 6) = "vsBgSw" Then id = "vsMnuBg"
    Select Case id
        Case "vsBtnFolder": tip = Tx("rb.vsBtnFolder.sup", FolderSettingText())
        Case "vsBtnOpenFolder": tip = Tx("rb.vsBtnOpenFolder.sup", FileManagerName())
        Case "vsGalColor", "vsEbRgb", "vsBgCustomMac": tip = Tx("rb.vsGalColor.sup", ColorSettingText())
        Case "vsMnuBg": tip = Tx("rb.vsMnuBg.sup", BgName(BgTypeIdx()))
        Case "vsMnuLang": tip = Tx("rb.vsMnuLang.sup", LangName(CurLang()))
        Case Else
            If TxHas("rb." & id & ".sup") Then tip = Tx("rb." & id & ".sup") Else tip = SettingsSummary()
    End Select
End Sub

Public Sub VS_GetDescription(control As Object, ByRef desc)
    If TxHas("rb." & control.Id & ".desc") Then desc = Tx("rb." & control.Id & ".desc") Else desc = ""
End Sub

Public Sub VS_GetTitle(control As Object, ByRef title)
    title = Tx("rb." & control.Id)
End Sub

' ---- background: preview square (Windows: drawn; macOS: one of the static icons) and menu
Public Sub VS_GetVisible(control As Object, ByRef visible)
    Dim mac As Boolean, bg As Long
#If Mac Then
    mac = True
#End If
    bg = BgTypeIdx()
    Select Case control.Id
        Case "vsBgSwatch", "vsGalColor": visible = Not mac
        Case "vsBgCustomMac": visible = mac
        Case "vsBgSwNone": visible = mac And bg = BG_NONE
        Case "vsBgSwWhite": visible = mac And bg = BG_WHITE
        Case "vsBgSwBeige": visible = mac And bg = BG_BEIGE
        Case "vsBgSwGrid": visible = mac And bg = BG_GRID
        Case "vsBgSwCustom": visible = mac And bg = BG_CUSTOM
        Case Else: visible = True
    End Select
End Sub

Public Sub VS_GetImage(control As Object, ByRef image)
#If Mac Then
#Else
    Select Case control.Id
        Case "vsBgSwatch": Set image = BgPicture(BgTypeIdx(), 16)
        Case "vsGalColor": Set image = BgPicture(BG_CUSTOM, 32)
    End Select
#End If
End Sub

' ---- background colour gallery (theme / standard / recent colours), Windows
Public Sub VS_GalColorCount(control As Object, ByRef count)
    Dim rc As Variant
    rc = RecentColors()
    count = 70 + UBound(rc) + 1
End Sub

Public Sub VS_GalColorImage(control As Object, index As Integer, ByRef image)
#If Mac Then
#Else
    Set image = SwatchPicture(GalColor(index), 16)
#End If
End Sub

Public Sub VS_GalColorTip(control As Object, index As Integer, ByRef tip)
    Dim label As String
    GalColor index, label
    tip = label
End Sub

Public Sub VS_GalColorPick(control As Object, id As String, index As Integer)
    SetBgColor GalColor(index)
End Sub

' ---- drop-downs with translated items
Public Sub VS_GetItemCount(control As Object, ByRef count)
    Select Case control.Id
        Case "vsDdMode", "vsDdNaming": count = 2
        Case "vsDdDpi": count = 4
        Case Else: count = 0
    End Select
End Sub

Public Sub VS_GetItemLabel(control As Object, index As Integer, ByRef label)
    Select Case control.Id
        Case "vsDdMode": label = Tx(IIf(index = 0, "mode.merge", "mode.split"))
        Case "vsDdNaming": label = Tx(IIf(index = 0, "naming.ask", "naming.auto"))
        Case "vsDdDpi"
            If index < 3 Then label = CStr(Array(150, 300, 600)(index)) & " DPI" Else label = Tx("dpi.custom")
        Case Else: label = ""
    End Select
End Sub

Public Sub VS_GetItemID(control As Object, index As Integer, ByRef id)
    id = control.Id & "_" & index
End Sub

' ---- donations, project page
Public Sub VS_OnDonate(control As Object, id As String, index As Integer)
    If index = 0 Then DoDonateThanks Tx("don.alipay") Else DoDonateThanks Tx("don.wechat")
End Sub

Public Sub VS_OnDonateInfo(control As Object)
    DoAbout
End Sub

Public Sub VS_OnHome(control As Object)
    If Not OpenUrlX(VS_REPO) Then MsgBox Tx("msg.home", VS_REPO), vbInformation, AppTitle()
End Sub

Public Sub VS_OnOpenFolder(control As Object)
    Dim f As String
    f = GetSet("LastFolder", "")
    If Len(f) = 0 Or Not FolderExistsX(f) Then f = GetSet("Folder", "")
    If Len(f) = 0 Or Not FolderExistsX(f) Then
        MsgBox Tx("msg.noFolderYet"), vbInformation, AppTitle()
        Exit Sub
    End If
    OpenPathX f
End Sub

Public Sub VS_GetIndex(control As Object, ByRef index)
    Select Case control.Id
        Case "vsDdFormat": index = GetIdx("Format", FMT_LAST)
        Case "vsDdMode": index = GetIdx("Mode", 1)
        Case "vsDdNaming": index = GetIdx("Naming", 1)
        Case "vsDdDpi": index = GetIdx("Dpi", 3, 1)
        Case Else: index = 0
    End Select
End Sub

Public Sub VS_OnDropDown(control As Object, id As String, index As Integer)
    Select Case control.Id
        Case "vsDdFormat": PutSet "Format", CStr(index)
        Case "vsDdMode": PutSet "Mode", CStr(index)
        Case "vsDdNaming": PutSet "Naming", CStr(index)
        Case "vsDdDpi": PutSet "Dpi", CStr(index)
    End Select
    RefreshRibbon
End Sub

' ---- check boxes and toggle buttons (background entries, languages)
Public Sub VS_GetPressed(control As Object, ByRef returnedVal)
    Select Case control.Id
        Case "vsChkOpen": returnedVal = GetFlag("OpenFolder", "0")
        Case "vsChkVsdx": returnedVal = GetFlag("VsdxNative", "1")
        Case "vsBgNone": returnedVal = (BgTypeIdx() = BG_NONE)
        Case "vsBgWhite": returnedVal = (BgTypeIdx() = BG_WHITE)
        Case "vsBgBeige": returnedVal = (BgTypeIdx() = BG_BEIGE)
        Case "vsBgGrid": returnedVal = (BgTypeIdx() = BG_GRID)
        Case "vsBgCustomMac": returnedVal = (BgTypeIdx() = BG_CUSTOM)
        Case "vsLangZhCN": returnedVal = (CurLang() = "zh-CN")
        Case "vsLangZhTW": returnedVal = (CurLang() = "zh-TW")
        Case "vsLangEn": returnedVal = (CurLang() = "en")
        Case "vsLangJa": returnedVal = (CurLang() = "ja")
        Case Else: returnedVal = False
    End Select
End Sub

Public Sub VS_OnCheck(control As Object, pressed As Boolean)
    Dim v As String
    If pressed Then v = "1" Else v = "0"
    Select Case control.Id
        Case "vsChkOpen": PutSet "OpenFolder", v
        Case "vsChkVsdx": PutSet "VsdxNative", v
        Case "vsBgNone": PutSet "BgType", CStr(BG_NONE)
        Case "vsBgWhite": PutSet "BgType", CStr(BG_WHITE)
        Case "vsBgBeige": PutSet "BgType", CStr(BG_BEIGE)
        Case "vsBgGrid": PutSet "BgType", CStr(BG_GRID)
        Case "vsBgCustomMac": DoMoreColors
        Case "vsLangZhCN": SetLang "zh-CN"
        Case "vsLangZhTW": SetLang "zh-TW"
        Case "vsLangEn": SetLang "en"
        Case "vsLangJa": SetLang "ja"
    End Select
    RefreshRibbon
End Sub

Public Sub VS_GetEnabled(control As Object, ByRef enabled)
    Select Case control.Id
        Case "vsEbDpi": enabled = (GetIdx("Dpi", 3, 1) = 3)
        Case "vsEbGrid": enabled = (BgTypeIdx() = BG_GRID)
        Case "vsEbAlpha": enabled = (BgTypeIdx() <> BG_NONE)
        Case Else: enabled = True
    End Select
End Sub

Public Sub VS_GetText(control As Object, ByRef strText)
    Select Case control.Id
        Case "vsEbMargin": strText = NumText(GetMargin())
        Case "vsEbDpi": strText = CStr(CustomDpi())
        Case "vsEbRgb": strText = ColorSettingText()
        Case "vsEbAlpha": strText = CStr(BgAlpha())
        Case "vsEbGrid": strText = NumText(GridStep())
        Case Else: strText = ""
    End Select
End Sub

Public Sub VS_OnChange(control As Object, strText As String)
    Dim v As Double, r As Long, g As Long, b As Long
    Select Case control.Id
        Case "vsEbMargin"
            If ReadNumber(strText, 0, 500, v) Then PutSet "Margin", NumText(v) _
                Else MsgBox Tx("msg.badMargin"), vbExclamation, AppTitle()
        Case "vsEbDpi"
            If ReadNumber(strText, 1, DPI_MAX, v) Then
                PutSet "DpiCustom", CStr(CLng(v))
            Else
                MsgBox Tx("msg.badDpi", DPI_MAX), vbExclamation, AppTitle()
            End If
        Case "vsEbAlpha"
            If ReadNumber(strText, 0, 100, v) Then PutSet "BgAlpha", CStr(CLng(v)) _
                Else MsgBox Tx("msg.badAlpha"), vbExclamation, AppTitle()
        Case "vsEbGrid"
            If ReadNumber(strText, 2, 200, v) Then PutSet "GridStep", NumText(v) _
                Else MsgBox Tx("msg.badGrid"), vbExclamation, AppTitle()
        Case "vsEbRgb"
            If ParseColor(strText, r, g, b) Then
                SetBgColor RGB(r, g, b)
            Else
                MsgBox Tx("col.bad"), vbExclamation, AppTitle()
            End If
    End Select
    RefreshRibbon
End Sub

Public Sub VS_OnFolder(control As Object)
    Dim cur As String, f As String
    cur = GetSet("Folder", "")
    If Len(cur) = 0 Or Not FolderExistsX(cur) Then cur = DesktopPathX()
    f = PickFolderX(cur, Tx("set.pickDefaultFolder"))
    If Len(f) > 0 Then
        PutSet "Folder", f
        RefreshRibbon
    End If
End Sub

Public Sub VS_OnReset(control As Object)
    If MsgBox(Tx("msg.askReset"), vbQuestion + vbYesNo, AppTitle()) <> vbYes Then Exit Sub
    ResetSettings
    RefreshRibbon
End Sub

Public Sub VS_OnSettings(control As Object)
    MsgBox SettingsSummary(), vbInformation, AppTitle() & " - " & Tx("set.current")
End Sub

Public Sub VS_OnAbout(control As Object)
    DoAbout
End Sub

Public Sub VS_OnFeedback(control As Object)
    DoFeedback
End Sub

Public Sub VS_OnHelp(control As Object)
    MsgBox Tx("help.ppt", DPI_MAX), vbInformation, AppTitle()
End Sub


' =====================================================================
'  Main flows
' =====================================================================
Private Sub DoExport(ByVal fmt As Long)
    Dim win As Object, sr As Object, pres As Object
    Dim files As New Collection
    Dim path As String, folder As String, baseName As String, p As String
    Dim i As Long, autoSave As Boolean, separate As Boolean

    If gBusy Then Exit Sub
    On Error GoTo EH
    If Not GetTarget(win, sr, True) Then Exit Sub
    Set pres = win.Presentation
    ResetRunState
    autoSave = (GetIdx("Naming", 1) = 1)
    separate = (GetIdx("Mode", 1) = 1) And (sr.Count > 1)
    baseName = BaseNameFor(pres, sr)

    If separate Then
        If autoSave Then
            folder = DefaultFolder(pres)
        Else
            folder = PickFolderX(InitialFolder(pres), Tx("dlg.pickFolderEach"))
            If Len(folder) = 0 Then Exit Sub
        End If
        gBusy = True
        ProgressBegin Tx("prog.exportN", sr.Count, UCase$(FmtExt(fmt))), sr.Count
        For i = 1 To sr.Count
            ProgressItem i, Tx("prog.item", i, sr.Count, ShortName(sr.Item(i).Name))
            CheckCancel
            path = UniquePath(folder, baseName & "_" & SafeName(sr.Item(i).Name), FmtExt(fmt))
            p = ExportItem(pres, sr.Item(i), path, fmt)
            If Len(p) > 0 Then files.Add p
        Next i
    Else
        If autoSave Then
            path = UniquePath(DefaultFolder(pres), baseName & "_" & Format$(Now, "yyyymmdd_hhnnss"), FmtExt(fmt))
        Else
            path = AskSavePath(InitialFolder(pres), baseName, fmt)
            If Len(path) = 0 Then Exit Sub
        End If
        gBusy = True
        ProgressBegin Tx("prog.export1", UCase$(FmtExt(fmt))), 1
        ProgressItem 1, ""
        p = ExportItem(pres, sr, path, fmt)
        If Len(p) > 0 Then files.Add p
    End If

    ProgressEnd
    gBusy = False
    Report files, autoSave Or separate
    Exit Sub
EH:
    ProgressEnd
    gBusy = False
    If Err.Number = ERR_CANCEL Then
        ReportCancelled files
    Else
        MsgBox Tx("msg.exportFailed", Err.Description), vbExclamation, AppTitle()
    End If
End Sub

' One file per slide: all slides, or only the slides selected in the thumbnail pane.
Private Sub DoBatch()
    Dim win As Object, pres As Object, sld As Object, rng As Object
    Dim slideList As New Collection, files As New Collection
    Dim fmt As Long, i As Long, ans As VbMsgBoxResult
    Dim folder As String, baseName As String, path As String, p As String, skipped As Long

    On Error Resume Next
    Set win = Application.ActiveWindow
    On Error GoTo EH
    If win Is Nothing Then
        MsgBox Tx("msg.openPres"), vbInformation, AppTitle()
        Exit Sub
    End If
    Set pres = win.Presentation
    fmt = GetIdx("Format", FMT_LAST)

    If win.Selection.Type = PP_SEL_SLIDES Then
        If win.Selection.SlideRange.Count > 1 Then
            ans = MsgBox(Tx("msg.askWhich", win.Selection.SlideRange.Count, pres.Slides.Count), _
                         vbQuestion + vbYesNoCancel, AppTitle())
            If ans = vbCancel Then Exit Sub
            If ans = vbYes Then
                For i = 1 To win.Selection.SlideRange.Count
                    slideList.Add win.Selection.SlideRange.Item(i)
                Next i
            End If
        End If
    End If
    If slideList.Count = 0 Then
        For i = 1 To pres.Slides.Count
            slideList.Add pres.Slides.Item(i)
        Next i
    End If
    If slideList.Count = 0 Then
        MsgBox Tx("msg.noSlides"), vbInformation, AppTitle()
        Exit Sub
    End If

    If GetIdx("Naming", 1) = 1 Then
        folder = DefaultFolder(pres)
    Else
        folder = PickFolderX(InitialFolder(pres), Tx("dlg.pickFolderBatch", UCase$(FmtExt(fmt))))
        If Len(folder) = 0 Then Exit Sub
    End If

    ResetRunState
    baseName = SafeName(BaseNameOf(pres.Name))
    If gBusy Then Exit Sub
    gBusy = True
    ProgressBegin Tx("prog.batch", slideList.Count, UCase$(FmtExt(fmt))), slideList.Count
    For i = 1 To slideList.Count
        Set sld = slideList(i)
        ProgressItem i, Tx("prog.page", i, slideList.Count, sld.SlideIndex)
        CheckCancel
        Set rng = ExportableRange(sld)
        If rng Is Nothing Then
            skipped = skipped + 1
        Else
            path = UniquePath(folder, baseName & "_" & Tx("file.page", sld.SlideIndex), FmtExt(fmt))
            p = ExportItem(pres, rng, path, fmt)
            If Len(p) > 0 Then files.Add p
        End If
    Next i
    ProgressEnd
    gBusy = False

    If files.Count = 0 Then
        MsgBox Tx("msg.noShapesSlides"), vbInformation, AppTitle()
        Exit Sub
    End If
    If skipped > 0 Then
        MsgBox Tx("msg.skipped", skipped), vbInformation, AppTitle()
    End If
    Report files, True
    Exit Sub
EH:
    ProgressEnd
    gBusy = False
    If Err.Number = ERR_CANCEL Then
        ReportCancelled files
    Else
        MsgBox Tx("msg.batchFailed", Err.Description), vbExclamation, AppTitle()
    End If
End Sub

' Pastes an EMF picture copy of the selection next to it (the original stays untouched).
Private Sub DoPictureCopy()
    Dim win As Object, sr As Object, sld As Object, pic As Object
    On Error GoTo EH
    If Not GetTarget(win, sr, False) Then Exit Sub
    Set sld = sr.Parent
    Set pic = PasteFromClipboard(sld, sr, True).Item(1)
    pic.Name = "VecStamp EMF " & Format$(Now, "hhnnss")
    pic.Select
    Exit Sub
EH:
    MsgBox Tx("msg.pictureFailed", Err.Description), vbExclamation, AppTitle()
End Sub

' The system colour dialog for the custom background (Windows: the common colour dialog,
' macOS: the macOS colour panel). Public macro VS_PickColor.
Private Sub DoPickColor()
    Dim c As Long
    On Error GoTo EH
    c = BgColorValue(BG_CUSTOM)
#If Mac Then
    If MacColorPicker(c) Then SetBgColor c
#Else
    If SysColorDialog(c, GetActiveWindow()) Then SetBgColor c
#End If
    Exit Sub
EH:
    MsgBox Tx("msg.colorFailed", Err.Description), vbExclamation, AppTitle()
End Sub

#If Mac Then
' macOS colour panel (the same panel PowerPoint's "More Colors" opens on the Mac). Without the
' helper script VecStamp's own colour list is used instead.
Private Function MacColorPicker(ByRef c As Long) As Boolean
    Dim res As String, r As Long, g As Long, b As Long
    On Error GoTo Fallback
    res = MacScript2("chooseColor", (c And &HFF&) & "," & ((c \ &H100&) And &HFF&) & "," & ((c \ &H10000) And &HFF&))
    If res = "error" Then GoTo Fallback
    If Len(res) = 0 Then Exit Function                  ' cancelled
    If ParseColor(res, r, g, b) Then
        c = RGB(r, g, b)
        MacColorPicker = True
    End If
    Exit Function
Fallback:
    MacColorPicker = ColorListPanel(c)
End Function
#Else
Private Function SysColorDialog(ByRef c As Long, ByVal owner As LongPtr) As Boolean
    Dim cc As CHOOSECOLORW
    cc.lStructSize = LenB(cc)
    cc.hwndOwner = owner
    cc.rgbResult = c
    cc.lpCustColors = VarPtr(gCustColors(0))
    cc.Flags = CC_RGBINIT Or CC_FULLOPEN Or CC_ANYCOLOR
    If ChooseColorW(cc) = 0 Then Exit Function
    c = cc.rgbResult And &HFFFFFF
    SysColorDialog = True
End Function

' VecStamp's own colour panel, used when PowerPoint's "More Colors" dialog cannot be opened:
' theme colours with tints and shades, standard and recent colours, a colour value field
' (R,G,B or #RRGGBB) with a preview, the system colour dialog, OK and Cancel. It is a native
' window driven by the message loop below - clicks are read from the messages, so no window
' procedure is needed. Returns 1 (c = chosen colour), 0 (cancelled) or -1 (no window).
Private Function ColorPanel(ByRef c As Long) As Long
    Dim owner As LongPtr, hwnd As LongPtr, hdc As LongPtr, font As LongPtr, fontB As LongPtr
    Dim hTheme As LongPtr, hStd As LongPtr, hRecent As LongPtr, hEdit As LongPtr, hPrev As LongPtr, hInfo As LongPtr
    Dim hOk As LongPtr, hCancel As LongPtr, hSys As LongPtr, clicked As LongPtr, f As LongPtr
    Dim bmp(0 To 2) As LongPtr, i As Long
    Dim m As MSG_T, r As RECT, k As Double, dpi As Long, cs As Long, gp As Long, rowGap As Long, lm As Long
    Dim x As Long, y As Long, w As Long, h As Long, cw As Long, ch As Long, px As Long, py As Long, pw As Long, ph As Long
    Dim idx As Long, act As Long, cur As Long, txt As String, lastTxt As String, handled As Boolean
    Dim rc As Variant, nRecent As Long, lastClick As Double, lastIdx As Long, lbl As String
    Dim rr As Long, gg As Long, bb As Long, cls As String, cap As String, face As String, cols() As Long

    On Error GoTo Failed
    ColorPanel = -1
    owner = GetActiveWindow()
    hdc = GetDC(0)
    dpi = GetDeviceCaps(hdc, LOGPIXELSY)
    ReleaseDC 0, hdc
    If dpi < 72 Then dpi = 96
    k = dpi / 96#
    cs = CLng(18 * k): gp = CLng(4 * k): rowGap = CLng(5 * k): lm = CLng(16 * k)
    rc = RecentColors()
    nRecent = UBound(rc) + 1
    cur = c
    lastIdx = -1

    ' layout (client coordinates)
    cw = 2 * lm + 10 * cs + 9 * gp
    If cw < CLng(320 * k) Then cw = CLng(320 * k)
    pw = CLng(64 * k): ph = CLng(24 * k)
    y = CLng(12 * k)
    ch = y + CLng(20 * k) + 6 * cs + rowGap + CLng(10 * k) + CLng(20 * k) + cs + CLng(10 * k)
    If nRecent > 0 Then ch = ch + CLng(20 * k) + cs + CLng(10 * k)
    ch = ch + CLng(32 * k) + CLng(42 * k) + CLng(42 * k)

    r.Left = 0: r.Top = 0: r.Right = cw: r.Bottom = ch
    AdjustWindowRectEx r, WS_POPUP Or WS_CAPTION, 0, WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT
    w = r.Right - r.Left
    h = r.Bottom - r.Top
    If owner <> 0 And GetWindowRect(owner, r) <> 0 Then
        x = r.Left + (r.Right - r.Left - w) \ 2
        y = r.Top + (r.Bottom - r.Top - h) \ 3
    Else
        x = (GetSystemMetrics(0) - w) \ 2
        y = (GetSystemMetrics(1) - h) \ 3
    End If
    cls = "#32770"
    cap = AppTitle() & " - " & Tx("col.title")
    hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT, StrPtr(cls), StrPtr(cap), _
                           WS_POPUP Or WS_CAPTION, x, y, w, h, owner, 0, 0, 0)
    If hwnd = 0 Then Exit Function
    face = UiFontName()
    font = CreateFontW(-CLng(9 * dpi / 72), 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 5, 0, StrPtr(face))
    fontB = CreateFontW(-CLng(9 * dpi / 72), 0, 0, 0, 700, 0, 0, 0, 1, 0, 0, 5, 0, StrPtr(face))

    ' theme colours (row 0) with their tints and shades (rows 1-5)
    y = CLng(12 * k)
    PanelChild hwnd, "Static", Tx("col.theme"), SS_NOPREFIX, lm, y, cw - 2 * lm, CLng(18 * k), fontB
    y = y + CLng(20 * k)
    ReDim cols(0 To 59)
    For i = 0 To 59: cols(i) = GalColor(i): Next i
    hTheme = PanelChild(hwnd, "Static", "", SS_BITMAP Or SS_NOTIFY, lm, y, 0, 0, 0)
    bmp(0) = PanelBitmap(hTheme, "theme", 10, 6, cols, cs, cs, gp, rowGap)
    y = y + 6 * cs + rowGap + CLng(10 * k)
    ' standard colours
    PanelChild hwnd, "Static", Tx("col.standard"), SS_NOPREFIX, lm, y, cw - 2 * lm, CLng(18 * k), fontB
    y = y + CLng(20 * k)
    ReDim cols(0 To 9)
    For i = 0 To 9: cols(i) = GalColor(60 + i): Next i
    hStd = PanelChild(hwnd, "Static", "", SS_BITMAP Or SS_NOTIFY, lm, y, 0, 0, 0)
    bmp(1) = PanelBitmap(hStd, "std", 10, 1, cols, cs, cs, gp, 0)
    y = y + cs + CLng(10 * k)
    ' recent colours
    If nRecent > 0 Then
        PanelChild hwnd, "Static", Tx("col.recent"), SS_NOPREFIX, lm, y, cw - 2 * lm, CLng(18 * k), fontB
        y = y + CLng(20 * k)
        ReDim cols(0 To nRecent - 1)
        For i = 0 To nRecent - 1: cols(i) = CLng(rc(i)): Next i
        hRecent = PanelChild(hwnd, "Static", "", SS_BITMAP Or SS_NOTIFY, lm, y, 0, 0, 0)
        bmp(2) = PanelBitmap(hRecent, "recent", nRecent, 1, cols, cs, cs, gp, 0)
        y = y + cs + CLng(10 * k)
    End If
    ' colour value, preview and the name of the colour
    PanelChild hwnd, "Static", Tx("col.value"), SS_NOPREFIX, lm, y + CLng(4 * k), CLng(80 * k), CLng(18 * k), font
    hEdit = PanelChild(hwnd, "Edit", HexRGB(cur), WS_TABSTOP Or ES_AUTOHSCROLL, lm + CLng(84 * k), y, _
                       CLng(110 * k), ph, font, WS_EX_CLIENTEDGE)
    SendMessageW hEdit, EM_LIMITTEXT, 24, 0
    hPrev = PanelChild(hwnd, "Static", "", SS_BITMAP, lm + CLng(202 * k), y, 0, 0, 0)
    PanelPreview hPrev, cur, pw, ph
    y = y + CLng(32 * k)
    hInfo = PanelChild(hwnd, "Static", Tx("col.hint"), SS_NOPREFIX, lm, y, cw - 2 * lm, CLng(36 * k), font)
    y = y + CLng(42 * k)
    hSys = PanelChild(hwnd, "Button", Tx("col.system"), WS_TABSTOP, lm, y, CLng(128 * k), CLng(28 * k), font)
    hOk = PanelChild(hwnd, "Button", Tx("btn.ok"), WS_TABSTOP Or BS_DEFPUSHBUTTON, cw - lm - CLng(176 * k), y, _
                     CLng(84 * k), CLng(28 * k), font)
    hCancel = PanelChild(hwnd, "Button", Tx("btn.cancel"), WS_TABSTOP, cw - lm - CLng(84 * k), y, _
                         CLng(84 * k), CLng(28 * k), font)

    ' modal loop
    If owner <> 0 Then EnableWindow owner, 0
    ShowWindow hwnd, SW_SHOW
    UpdateWindow hwnd
    SetFocusW hEdit
    SendMessageW hEdit, EM_SETSEL, 0, -1
    lastTxt = HexRGB(cur)
    ColorPanel = 0
    Do
        If IsWindow(hwnd) = 0 Then Exit Do
        If PeekMessageW(m, 0, 0, 0, PM_REMOVE) <> 0 Then
            handled = False
            Select Case m.message
                Case WM_QUIT
                    PostQuitMessage CLng(m.wParam)
                    act = 2
                    handled = True
                Case WM_LBUTTONDOWN, WM_LBUTTONDBLCLK
                    px = CLng(m.lParam And &HFFFF&)
                    py = CLng((m.lParam \ &H10000) And &HFFFF&)
                    idx = -1
                    If m.hwnd = hTheme And hTheme <> 0 Then idx = PanelHit(px, py, 10, 6, cs, gp, rowGap, 0)
                    If m.hwnd = hStd And hStd <> 0 Then idx = PanelHit(px, py, 10, 1, cs, gp, 0, 60)
                    If m.hwnd = hRecent And hRecent <> 0 Then idx = PanelHit(px, py, nRecent, 1, cs, gp, 0, 70)
                    If idx >= 0 Then
                        cur = GalColor(idx, lbl)
                        lastTxt = HexRGB(cur)
                        SetWindowTextW hEdit, StrPtr(lastTxt)
                        PanelPreview hPrev, cur, pw, ph
                        SetWindowTextW hInfo, StrPtr(lbl)
                        If m.message = WM_LBUTTONDBLCLK Or (idx = lastIdx And Abs(Timer - lastClick) < GetDoubleClickTime() / 1000#) Then act = 1
                        lastIdx = idx
                        lastClick = Timer
                        handled = True
                    End If
                Case WM_LBUTTONUP
                    If m.hwnd = hOk Or m.hwnd = hCancel Or m.hwnd = hSys Then
                        GetClientRect m.hwnd, r
                        px = CLng(m.lParam And &HFFFF&)
                        py = CLng((m.lParam \ &H10000) And &HFFFF&)
                        If px < r.Right And py < r.Bottom Then clicked = m.hwnd
                    End If
                Case WM_KEYDOWN
                    Select Case CLng(m.wParam)
                        Case VK_ESCAPE
                            act = 2
                            handled = True
                        Case VK_RETURN
                            f = GetFocus()
                            If f = hCancel Then
                                act = 2
                            ElseIf f = hSys Then
                                clicked = hSys
                            Else
                                act = 1
                            End If
                            handled = True
                    End Select
                Case WM_KEYUP
                    If CLng(m.wParam) = VK_SPACE Then
                        f = GetFocus()
                        If f = hOk Or f = hCancel Or f = hSys Then clicked = f
                    End If
            End Select
            If Not handled Then
                If IsDialogMessageW(hwnd, m) = 0 Then
                    TranslateMessage m
                    DispatchMessageW m
                End If
            End If
            If clicked <> 0 Then
                If clicked = hOk Then act = 1
                If clicked = hCancel Then act = 2
                If clicked = hSys Then
                    If SysColorDialog(cur, hwnd) Then
                        lastTxt = HexRGB(cur)
                        SetWindowTextW hEdit, StrPtr(lastTxt)
                        PanelPreview hPrev, cur, pw, ph
                        lbl = Tx("col.withValue", Tx("col.systemPicked"), ColorText(cur))
                        SetWindowTextW hInfo, StrPtr(lbl)
                    End If
                End If
                clicked = 0
            End If
        Else
            SleepX 10
        End If
        txt = WindowText(hEdit)
        If txt <> lastTxt Then                      ' typed colour value
            lastTxt = txt
            If ParseColor(txt, rr, gg, bb) Then
                cur = RGB(rr, gg, bb)
                PanelPreview hPrev, cur, pw, ph
                SetWindowTextW hInfo, StrPtr(ColorText(cur))
            End If
        End If
        If act = 1 Then
            If ParseColor(WindowText(hEdit), rr, gg, bb) Then
                c = RGB(rr, gg, bb)
                ColorPanel = 1
                Exit Do
            End If
            lbl = Tx("col.bad")
            SetWindowTextW hInfo, StrPtr(lbl)
            act = 0
        ElseIf act = 2 Then
            Exit Do
        End If
    Loop

Cleanup:
    On Error Resume Next
    If owner <> 0 Then EnableWindow owner, 1
    If hPrev <> 0 Then DeleteObject SendMessageW(hPrev, STM_SETIMAGE, IMAGE_BITMAP, 0)
    If hwnd <> 0 Then DestroyWindow hwnd
    For i = 0 To 2
        If bmp(i) <> 0 Then DeleteObject bmp(i)
    Next i
    If font <> 0 Then DeleteObject font
    If fontB <> 0 Then DeleteObject fontB
    Exit Function
Failed:
    If hwnd = 0 Then ColorPanel = -1 Else ColorPanel = 0
    Resume Cleanup
End Function

Private Function PanelChild(ByVal parent As LongPtr, ByVal cls As String, ByVal text As String, ByVal style As Long, _
                            ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal font As LongPtr, _
                            Optional ByVal exStyle As Long = 0) As LongPtr
    Dim hw As LongPtr
    hw = CreateWindowExW(exStyle, StrPtr(cls), StrPtr(text), WS_CHILD Or WS_VISIBLE Or style, x, y, w, h, parent, 0, 0, 0)
    If hw <> 0 And font <> 0 Then SendMessageW hw, WM_SETFONT, font, 1
    PanelChild = hw
End Function

Private Function WindowText(ByVal hw As LongPtr) As String
    Dim buf As String, n As Long
    buf = String$(128, vbNullChar)
    n = GetWindowTextW(hw, StrPtr(buf), 128)
    If n > 0 Then WindowText = Left$(buf, n)
End Function

' Index of the colour cell at (x, y) of a palette bitmap (first = index of its first cell), or -1.
Private Function PanelHit(ByVal x As Long, ByVal y As Long, ByVal cols As Long, ByVal rows As Long, ByVal cs As Long, _
                          ByVal gp As Long, ByVal rowGap As Long, ByVal first As Long) As Long
    Dim col As Long, row As Long
    PanelHit = -1
    If x Mod (cs + gp) >= cs Then Exit Function
    col = x \ (cs + gp)
    If col >= cols Then Exit Function
    If y < cs Then
        row = 0
    ElseIf rows > 1 And y >= cs + rowGap Then
        row = 1 + (y - cs - rowGap) \ cs
    Else
        Exit Function
    End If
    If row >= rows Then Exit Function
    PanelHit = first + row * cols + col
End Function

' Shows a palette (cols x rows colour cells) in a bitmap static control; returns the bitmap.
Private Function PanelBitmap(ByVal hStatic As LongPtr, ByVal name As String, ByVal cols As Long, ByVal rows As Long, _
                             ByRef colors() As Long, ByVal cellW As Long, ByVal cellH As Long, ByVal gp As Long, _
                             ByVal rowGap As Long) As LongPtr
    Dim p As String, hb As LongPtr
    p = JoinPath(TempDirX(), "panel_" & name & "_" & cellW & ".bmp")
    CellsBmp p, cols, rows, colors, cellW, cellH, gp, rowGap
    hb = LoadImageW(0, StrPtr(p), IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
    If hb <> 0 Then SendMessageW hStatic, STM_SETIMAGE, IMAGE_BITMAP, hb
    PanelBitmap = hb
End Function

Private Sub PanelPreview(ByVal hStatic As LongPtr, ByVal c As Long, ByVal w As Long, ByVal h As Long)
    Dim p As String, hb As LongPtr, old As LongPtr, one(0 To 0) As Long
    one(0) = c
    p = JoinPath(TempDirX(), "panel_preview_" & w & ".bmp")
    CellsBmp p, 1, 1, one, w, h, 0, 0
    hb = LoadImageW(0, StrPtr(p), IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
    If hb = 0 Then Exit Sub
    old = SendMessageW(hStatic, STM_SETIMAGE, IMAGE_BITMAP, hb)
    If old <> 0 Then DeleteObject old
End Sub

' Writes a 24-bit BMP of colour cells (row 0 is separated from the other rows by rowGap) on the
' dialog background colour; every cell has a grey frame.
Private Sub CellsBmp(ByVal p As String, ByVal cols As Long, ByVal rows As Long, ByRef colors() As Long, _
                     ByVal cellW As Long, ByVal cellH As Long, ByVal gp As Long, ByVal rowGap As Long)
    Dim b() As Byte, w As Long, h As Long, rowBytes As Long, n As Long, x As Long, y As Long, i As Long
    Dim col As Long, row As Long, x0 As Long, y0 As Long, c As Long, bg As Long, raw As String
    w = cols * cellW + (cols - 1) * gp
    h = rows * cellH
    If rows > 1 Then h = h + rowGap
    rowBytes = ((w * 3 + 3) \ 4) * 4
    n = 54 + rowBytes * h
    ReDim b(0 To n - 1)
    BmpHeader b, w, h, rowBytes
    bg = GetSysColor(COLOR_BTNFACE)
    For y = 0 To h - 1
        For x = 0 To w - 1
            i = 54 + y * rowBytes + x * 3
            b(i) = (bg \ &H10000) And &HFF&: b(i + 1) = (bg \ &H100&) And &HFF&: b(i + 2) = bg And &HFF&
        Next x
    Next y
    For row = 0 To rows - 1
        For col = 0 To cols - 1
            c = colors(row * cols + col)
            x0 = col * (cellW + gp)
            y0 = row * cellH
            If row > 0 Then y0 = y0 + rowGap
            For y = y0 To y0 + cellH - 1
                For x = x0 To x0 + cellW - 1
                    i = 54 + (h - 1 - y) * rowBytes + x * 3
                    If x = x0 Or y = y0 Or x = x0 + cellW - 1 Or y = y0 + cellH - 1 Then
                        b(i) = 170: b(i + 1) = 170: b(i + 2) = 170
                    Else
                        b(i) = (c \ &H10000) And &HFF&: b(i + 1) = (c \ &H100&) And &HFF&: b(i + 2) = c And &HFF&
                    End If
                Next x
            Next y
        Next col
    Next row
    raw = b
    WriteBytesX p, raw
End Sub
#End If

Private Sub BmpHeader(ByRef b() As Byte, ByVal w As Long, ByVal h As Long, ByVal rowBytes As Long)
    b(0) = 66: b(1) = 77                                         ' "BM"
    PutLong b, 2, 54 + rowBytes * h
    PutLong b, 10, 54
    PutLong b, 14, 40
    PutLong b, 18, w
    PutLong b, 22, h
    b(26) = 1: b(28) = 24
    PutLong b, 34, rowBytes * h
End Sub

' Last-resort colour panel when no native window can be shown: a numbered list of theme,
' standard and recent colours in an input box; a number, R,G,B or #RRGGBB can be entered.
Private Function ColorListPanel(ByRef c As Long) As Boolean
    Dim msg As String, ans As String, i As Long, n As Long, lbl As String, idx(0 To 29) As Long
    Dim rc As Variant, r As Long, g As Long, b As Long
    For i = 0 To 9
        idx(n) = i: n = n + 1                                    ' theme colours
    Next i
    For i = 0 To 9
        idx(n) = 60 + i: n = n + 1                               ' standard colours
    Next i
    rc = RecentColors()
    For i = 0 To UBound(rc)
        If i < 5 Then idx(n) = 70 + i: n = n + 1
    Next i
    msg = Tx("col.listPrompt") & vbCrLf
    For i = 0 To n - 1
        GalColor idx(i), lbl
        msg = msg & (i + 1) & ". " & lbl & vbCrLf
    Next i
    Do
        ans = Trim$(InputBox(msg, AppTitle() & " - " & Tx("col.title"), HexRGB(c)))
        If Len(ans) = 0 Then Exit Function
        If IsNumeric(ans) And InStr(ans, ",") = 0 Then
            If CLng(Val(ans)) >= 1 And CLng(Val(ans)) <= n Then
                c = GalColor(idx(CLng(Val(ans)) - 1))
                ColorListPanel = True
                Exit Function
            End If
        ElseIf ParseColor(ans, r, g, b) Then
            c = RGB(r, g, b)
            ColorListPanel = True
            Exit Function
        End If
        MsgBox Tx("col.bad"), vbExclamation, AppTitle()
    Loop
End Function

Private Sub ResetRunState()
    gSvgFallback = 0
    gSvgNative = 0
    gSvgXpsOff = False
    gBigPngOk = False
End Sub

' What to export: the selection, or (after asking) every shape on the current slide.
Private Function GetTarget(ByRef win As Object, ByRef sr As Object, ByVal allowWholeSlide As Boolean) As Boolean
    Dim sld As Object, selType As Long
    On Error Resume Next
    Set win = Application.ActiveWindow
    selType = -1
    selType = win.Selection.Type
    On Error GoTo 0
    If win Is Nothing Or selType < 0 Then
        MsgBox Tx("msg.openAndSelect"), vbInformation, AppTitle()
        Exit Function
    End If

    If selType = PP_SEL_SHAPES Or selType = PP_SEL_TEXT Then
        If win.Selection.HasChildShapeRange Then
            Set sr = win.Selection.ChildShapeRange
        Else
            Set sr = win.Selection.ShapeRange
        End If
    Else
        If Not allowWholeSlide Then
            MsgBox Tx("msg.selectFirst"), vbInformation, AppTitle()
            Exit Function
        End If
        On Error Resume Next
        Set sld = win.View.Slide
        On Error GoTo 0
        If sld Is Nothing Then
            MsgBox Tx("msg.normalView"), vbInformation, AppTitle()
            Exit Function
        End If
        Set sr = ExportableRange(sld)
        If sr Is Nothing Then
            MsgBox Tx("msg.noShapesSlide"), vbInformation, AppTitle()
            Exit Function
        End If
        If MsgBox(Tx("msg.askWholeSlide"), vbQuestion + vbYesNo, AppTitle()) <> vbYes Then Exit Function
    End If
    GetTarget = Not (sr Is Nothing)
End Function

' All visible shapes on a slide except empty placeholders ("click to add title" boxes).
Private Function ExportableRange(ByVal sld As Object) As Object
    Dim i As Long, n As Long, shp As Object, keep As Boolean
    Dim idx() As Variant
    If sld.Shapes.Count = 0 Then Exit Function
    ReDim idx(1 To sld.Shapes.Count)
    For i = 1 To sld.Shapes.Count
        Set shp = sld.Shapes.Item(i)
        keep = True
        On Error Resume Next
        If shp.Visible = MSO_FALSE Then keep = False
        If shp.Type = MSO_PLACEHOLDER Then
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText = MSO_FALSE Then keep = False
            End If
        End If
        On Error GoTo 0
        If keep Then
            n = n + 1
            idx(n) = i
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve idx(1 To n)
    Set ExportableRange = sld.Shapes.Range(idx)
End Function

' Exports one shape or a range of shapes. Returns the file written ("" if skipped).
' Several shapes, a background or a margin are composed on a temporary slide.
Private Function ExportItem(ByVal pres As Object, ByVal src As Object, _
                            ByVal path As String, ByVal fmt As Long) As String
    Dim n As Long, margin As Single
    Dim tmp As Object, target As Object, wasSaved As Long
    Dim errNum As Long, errDesc As String, result As String

    n = 1
    If TypeName(src) = "ShapeRange" Then n = src.Count
    margin = GetMargin()
    ProgressAt 0.03, Tx("step.prepare", n)
    CheckCancel

    If n = 1 And BgTypeIdx() = BG_NONE And margin <= 0 Then
        If TypeName(src) = "ShapeRange" Then
            ExportItem = ExportShape(pres, src.Item(1), path, fmt)
        Else
            ExportItem = ExportShape(pres, src, path, fmt)
        End If
        Exit Function
    End If

    wasSaved = pres.Saved
    On Error GoTo Failed
    ProgressAt 0.08, Tx("step.compose")
    Set tmp = AddTempSlide(pres, src)
    Set target = BuildComposite(tmp, src, margin)
    ProgressAt 0.18, Tx("step.composed")
    CheckCancel
    result = ExportShape(pres, target, path, fmt)

Cleanup:
    On Error Resume Next
    If Not tmp Is Nothing Then tmp.Delete
    pres.Saved = wasSaved
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, , errDesc
    ExportItem = result
    Exit Function
Failed:
    errNum = Err.Number
    errDesc = Err.Description
    Resume Cleanup
End Function

' Writes one shape in the requested format. Returns the path written ("" if skipped).
Private Function ExportShape(ByVal pres As Object, ByVal shp As Object, _
                             ByVal path As String, ByVal fmt As Long) As String
    Dim t As String, dpi As Long, px As Double
    Select Case fmt
        Case FMT_EMF
            ProgressAt 0.25, Tx("step.ppEmf")
            ExportRaw shp, path, PP_FMT_EMF
        Case FMT_WMF
            ProgressAt 0.25, Tx("step.ppWmf")
            ExportRaw shp, path, PP_FMT_WMF
        Case FMT_EMZ
            ProgressAt 0.25, Tx("step.ppEmf")
            t = TempFileX("emf")
            ExportToTemp shp, t, PP_FMT_EMF, 0, 0
            ProgressAt 0.65, Tx("step.emz")
            If Not GzipX(t, path) Then Err.Raise vbObjectError + 514, , Tx("err.emz")
        Case FMT_PNG
            dpi = DpiValue()
            px = (shp.Width / 72# * dpi) * (shp.Height / 72# * dpi)
            If px > PNG_WARN_PIXELS And Not gBigPngOk Then
                If AskUser(Tx("msg.bigPng", dpi, CLng(shp.Width / 72# * dpi), CLng(shp.Height / 72# * dpi), _
                              Format$(px / 1000000#, "0")), vbQuestion + vbYesNo) <> vbYes Then
                    Err.Raise vbObjectError + 599, , "cancelled"
                End If
                gBigPngOk = True
            End If
            ProgressAt 0.25, Tx("step.png", CLng(shp.Width / 72# * dpi), CLng(shp.Height / 72# * dpi), dpi)
            ExportRaw shp, path, PP_FMT_PNG, _
                      CLng(pres.PageSetup.SlideWidth * dpi / 96), CLng(pres.PageSetup.SlideHeight * dpi / 96)
        Case FMT_SVG
            path = ExportSvg(pres, shp, path)
        Case FMT_PDF
            ExportPdf shp, path
        Case FMT_VSDX
            ProgressAt 0.2, Tx("step.analyze")
            ExportVsdx pres, shp, path
    End Select
    If Len(path) > 0 Then
        If Not FileExistsX(path) Then Err.Raise vbObjectError + 515, , Tx("err.noFile", path)
    End If
    ProgressAt 1, Tx("step.doneFile", FileNameOf(path))
    ExportShape = path
End Function

' SVG, best method first:
'  1. PowerPoint's own SVG export (2302+ / Mac 16.82+), checked and repaired;
'  2. Windows, any PowerPoint 2013+: PowerPoint prints the drawing to XPS and the built-in
'     converter turns it into SVG (paths, gradients, images, real text, embedded fonts);
'  3. EMF converted with Inkscape or LibreOffice, when installed;
'  4. otherwise offer EMF instead.
Private Function ExportSvg(ByVal pres As Object, ByVal shp As Object, ByVal path As String) As String
    Dim ok As Boolean, t As String, e As String, p2 As String, st As Long

    If gSvgNative <> 2 Then
        ProgressAt 0.22, Tx("step.ppSvg")
        t = TempFileX("svg")
        On Error Resume Next
        ExportToTemp shp, t, PP_FMT_SVG, 0, 0
        ok = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
        ProgressAt 0.8, Tx("step.fixSvg")
        If ok Then ok = FixSvgFile(t)
        If ok Then
            gSvgNative = 1
            MoveTempTo t, path
            ExportSvg = path
            Exit Function
        End If
        gSvgNative = 2                              ' this PowerPoint cannot write SVG
        DeleteTempX t
        gPgFrac = 0.2                               ' the next method starts over
    End If

    st = RUN_FAILED
    t = TempFileX("svg")
    If Not gSvgXpsOff Then st = XpsToSvg(pres, shp, t)
    If st = RUN_CANCELLED Then
        DeleteTempX t
        Err.Raise ERR_CANCEL, , "cancelled"
    End If
    ok = (st = 0)
    If ok Then
        ProgressAt 0.96, Tx("step.fixSvg")
        ok = FixSvgFile(t)
    End If
    If ok Then
        MoveTempTo t, path
        ExportSvg = path
        Exit Function
    End If
    DeleteTempX t
    If st = RUN_TIMEOUT Then
        If AskUser(Tx("msg.svgTimeout", SVG_TIMEOUT), vbQuestion + vbYesNo) <> vbYes Then
            Err.Raise ERR_CANCEL, , "cancelled"
        End If
        gSvgXpsOff = True                           ' do not wait again for the next shapes of this run
    End If

    ProgressAt 0.5, Tx("step.emf2svg")
    e = TempFileX("emf")
    t = TempFileX("svg")
    ExportToTemp shp, e, PP_FMT_EMF, 0, 0
    st = ConvertEmfToSvg(e, t)
    If st = RUN_CANCELLED Then
        DeleteTempX e
        DeleteTempX t
        Err.Raise ERR_CANCEL, , "cancelled"
    End If
    ok = (st = 0)
    If ok Then ok = FixSvgFile(t)
    DeleteTempX e
    If ok Then
        MoveTempTo t, path
        ExportSvg = path
        Exit Function
    End If
    DeleteTempX t

    If gSvgFallback = 0 Then
        If AskUser(Tx("msg.noSvg"), vbQuestion + vbYesNo) = vbYes Then
            gSvgFallback = 1
        Else
            gSvgFallback = 2
        End If
    End If
    If gSvgFallback = 1 Then
        p2 = UniquePath(ParentFolder(path), BaseNameOf(path), "emf")
        ExportRaw shp, p2, PP_FMT_EMF
        ExportSvg = p2
    End If
End Function

' Makes sure an exported file really is SVG that browsers display as a picture:
' the root element must be <svg> in the SVG namespace (missing xmlns is added).
' Returns False when the file is not SVG at all.
Private Function FixSvgFile(ByVal p As String) As Boolean
    Dim data As String, k As Long, pos As Long, tagEnd As Long, nm As String
    Dim tag As String, ins As String, nextB As String

    data = ReadBytesX(p)
    If LenB(data) < 5 Then Exit Function
    If AscB(MidB(data, 1, 1)) = &HFF And AscB(MidB(data, 2, 1)) = &HFE Then
        FixSvgFile = (InStr(1, data, "<svg", vbTextCompare) > 0)      ' UTF-16 text: leave untouched
        Exit Function
    End If

    pos = 1
    Do
        k = InStrB(pos, data, AsciiB("<"))
        If k = 0 Or k > 65536 Then Exit Function
        nextB = MidB(data, k + 1, 1)
        If nextB = AsciiB("?") Or nextB = AsciiB("!") Then
            If MidB(data, k, 4) = AsciiB("<!--") Then
                pos = InStrB(k, data, AsciiB("-->"))
            Else
                pos = InStrB(k, data, AsciiB(">"))
            End If
            If pos = 0 Then Exit Function
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop

    tagEnd = InStrB(k, data, AsciiB(">"))
    If tagEnd = 0 Then Exit Function
    tag = MidB(data, k, tagEnd - k + 1)
    nm = LCase$(TagNameOf(tag))
    If nm = "svg" Then
        If InStrB(1, tag, AsciiB("http://www.w3.org/2000/svg")) = 0 Then
            ins = ins & AsciiB(" xmlns=""http://www.w3.org/2000/svg""")
        End If
        If InStrB(1, data, AsciiB("xlink:")) > 0 And InStrB(1, tag, AsciiB("xmlns:xlink")) = 0 Then
            ins = ins & AsciiB(" xmlns:xlink=""http://www.w3.org/1999/xlink""")
        End If
        If LenB(ins) > 0 Then
            data = LeftB(data, k + 3) & ins & MidB(data, k + 4)
            WriteBytesX p, data
        End If
        FixSvgFile = True
    ElseIf Right$(nm, 4) = ":svg" Then
        FixSvgFile = True
    End If
End Function

' Element name of a start tag held as bytes, e.g. "<svg width=..." -> "svg".
Private Function TagNameOf(ByVal tagBytes As String) As String
    Dim i As Long, c As Long, s As String
    For i = 2 To LenB(tagBytes)
        c = AscB(MidB(tagBytes, i, 1))
        If c = 32 Or c = 9 Or c = 10 Or c = 13 Or c = 62 Or c = 47 Then Exit For
        s = s & Chr$(c)
    Next i
    TagNameOf = s
End Function

' Vector PDF whose page is the size of the drawing (PowerPoint pages are 1 to 56 inches,
' so very small or very large drawings are scaled to fit those limits).
Private Sub ExportPdf(ByVal shp As Object, ByVal path As String)
    Dim tp As Object, sld As Object, pic As Object, t As String
    Dim w As Single, h As Single, k As Single, pw As Single, ph As Single
    Dim errNum As Long, errDesc As String

    w = shp.Width
    h = shp.Height
    If w < 1 Then w = 1
    If h < 1 Then h = 1
    k = 1
    If w < 72 Or h < 72 Then k = 72 / MinS(w, h)
    If MaxS(w, h) * k > 4032 Then k = 4032 / MaxS(w, h)
    pw = MinS(MaxS(w * k, 72), 4032)
    ph = MinS(MaxS(h * k, 72), 4032)
    t = TempFileX("pdf")
    ProgressAt 0.25, Tx("step.pdfDoc")

    On Error GoTo Failed
    Set tp = Application.Presentations.Add(MSO_FALSE)
    tp.PageSetup.SlideWidth = pw
    tp.PageSetup.SlideHeight = ph
    Set sld = tp.Slides.Add(1, PP_LAYOUT_BLANK)
    ProgressAt 0.45, Tx("step.pdfPaste")
    Set pic = PasteFromClipboard(sld, shp, True).Item(1)
    pic.LockAspectRatio = MSO_FALSE
    pic.Width = w * k
    pic.Height = h * k
    pic.Left = (pw - w * k) / 2
    pic.Top = (ph - h * k) / 2
    ProgressAt 0.7, Tx("step.pdfWrite")
    tp.SaveAs t, PP_SAVE_AS_PDF

Cleanup:
    On Error Resume Next
    If Not tp Is Nothing Then
        tp.Saved = MSO_TRUE
        tp.Close
    End If
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, , Tx("err.pdf", errDesc)
    MoveTempTo t, path
    Exit Sub
Failed:
    errNum = Err.Number
    errDesc = Err.Description
    Resume Cleanup
End Sub

' Visio drawing. Default: native Visio shapes built from the PowerPoint shapes (geometry,
' fill, line, arrowheads, rotation, editable rich text, groups) - no Visio needed.
' Pictures, charts, tables and other objects without a Visio equivalent are embedded as
' EMF pictures. "Picture mode" (setting VsdxNative = 0) embeds the whole drawing as one EMF.
Private Sub ExportVsdx(ByVal pres As Object, ByVal shp As Object, ByVal path As String)
    Dim e As String
    If GetFlag("VsdxNative", "1") Then
        If VsdxNative(pres, shp, path) Then Exit Sub
    End If
    ProgressAt 0.3, Tx("step.vsdxPicture")
    e = TempFileX("emf")
    ExportToTemp shp, e, PP_FMT_EMF, 0, 0
    gVxId = 0
    Set gVxMedia = New Collection
    SbReset
    SbAdd VxForeignXml(e, 0, 0, shp.Width, shp.Height, "VecStamp")
    ProgressAt 0.7, Tx("step.vsdxWrite0")
    WriteVsdxPackage shp.Width / 72#, shp.Height / 72#, SbText(), path
    DeleteTempX e
End Sub

Private Function VsdxNative(ByVal pres As Object, ByVal shp As Object, ByVal path As String) As Boolean
    Dim wasSaved As Long, ok As Boolean, i As Long, errNum As Long
    wasSaved = pres.Saved
    Set gVxPres = pres
    gVxId = 0
    gVxNative = 0
    gVxPictures = 0
    gVxCopy = 0
    gVxDone = 0
    gVxTotal = VxCount(shp)
    If gVxTotal < 1 Then gVxTotal = 1
    Set gVxMedia = New Collection
    SbReset
    On Error GoTo Failed
    Set gVxScratch = AddTempSlide(pres, shp)
    VxShape shp, shp.Left, shp.Top + shp.Height
    If gVxNative = 0 And gVxPictures = 0 Then GoTo Cleanup
    ProgressAt 0.88, Tx("step.vsdxWrite", gVxNative, gVxPictures)
    WriteVsdxPackage shp.Width / 72#, shp.Height / 72#, SbText(), path
    ok = True
Cleanup:
    On Error Resume Next
    If Not gVxScratch Is Nothing Then gVxScratch.Delete
    Set gVxScratch = Nothing
    For i = 1 To gVxMedia.Count
        DeleteTempX CStr(gVxMedia(i))
    Next i
    pres.Saved = wasSaved
    On Error GoTo 0
    If errNum = ERR_CANCEL Then Err.Raise ERR_CANCEL, , "cancelled"
    VsdxNative = ok
    Exit Function
Failed:
    ok = False
    errNum = Err.Number
    Resume Cleanup
End Function

' Number of shapes the walker will convert (group members counted one by one).
Private Function VxCount(ByVal s As Object) As Long
    Dim i As Long, t As Long
    On Error Resume Next
    t = s.Type
    If t = MSO_GROUP Then
        For i = 1 To s.GroupItems.Count
            VxCount = VxCount + 1
        Next i
    Else
        VxCount = 1
    End If
End Function

' Writes the package: template parts + shapes XML + embedded EMF pictures, zipped.
Private Sub WriteVsdxPackage(ByVal wIn As Double, ByVal hIn As Double, ByVal shapesXml As String, ByVal path As String)
    Dim d As String, z As String, m As Variant, i As Long, parts() As String, txt As String
    Dim dst As String, sub1 As Variant, rels As String

    d = JoinPath(TempDirX(), "vsdx_" & Format$(Now, "hhnnss") & "_" & CStr(Int(Rnd * 1000000)))
    MakeDirX d
    For Each sub1 In Array("_rels", "docProps", "visio", "visio/_rels", "visio/pages", "visio/pages/_rels", "visio/media")
        MakeDirX JoinPath(d, Replace(CStr(sub1), "/", PathSep()))
    Next
    For i = 1 To gVxMedia.Count
        rels = rels & "<Relationship Id=""rId" & i & """ Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"" Target=""../media/image" & i & ".emf""/>"
        CopyFileX CStr(gVxMedia(i)), JoinPath(d, Replace("visio/media/image" & i & ".emf", "/", PathSep()))
    Next i
    m = VsdxMap()
    For i = LBound(m) To UBound(m)
        parts = Split(CStr(m(i)), "=")
        txt = VsdxPart(parts(0))
        txt = Replace(txt, "{PW}", NumXml(wIn))
        txt = Replace(txt, "{PH}", NumXml(hIn))
        txt = Replace(txt, "{W}", NumXml(wIn))
        txt = Replace(txt, "{H}", NumXml(hIn))
        txt = Replace(txt, "{CX}", NumXml(wIn / 2))
        txt = Replace(txt, "{CY}", NumXml(hIn / 2))
        txt = Replace(txt, "{LX}", NumXml(wIn / 2))
        txt = Replace(txt, "{LY}", NumXml(hIn / 2))
        txt = Replace(txt, "{VERSION}", VS_VERSION)
        txt = Replace(txt, "{TITLE}", "VecStamp drawing")
        txt = Replace(txt, "{CREATED}", Format$(Now, "yyyy-mm-dd") & "T" & Format$(Now, "hh:nn:ss") & "Z")
        txt = Replace(txt, "{RELS}", rels)
        txt = Replace(txt, "{SHAPES}", shapesXml)     ' last: shape text may contain braces
        dst = JoinPath(d, Replace(parts(1), "/", PathSep()))
        WriteBytesX dst, Utf8Bytes(txt)
    Next i
    z = TempFileX("vsdx")
    ZipFolderX d, z
    RemoveTempDirX d
    MoveTempTo z, path
End Sub

' ---- shape walker -------------------------------------------------------------
' ox, oy: origin of the parent (page or group) in slide points: left and bottom edge.
Private Sub VxShape(ByVal s As Object, ByVal ox As Double, ByVal oy As Double)
    Dim t As Long, i As Long, x As String, nm As String
    t = -1
    On Error Resume Next
    t = s.Type
    If s.Visible = MSO_FALSE Then Exit Sub
    On Error GoTo 0
    If t = MSO_GROUP Then
        gVxId = gVxId + 1
        nm = VxName(s)
        x = "<Shape ID='" & gVxId & "' NameU='" & nm & "' Name='" & nm & "' Type='Group' LineStyle='0' FillStyle='0' TextStyle='0'>" & _
            VxXform(s.Left - ox, oy - (s.Top + s.Height), s.Width, s.Height, s.Rotation) & "<Shapes>"
        SbAdd x
        For i = 1 To s.GroupItems.Count
            VxShape s.GroupItems.Item(i), s.Left, s.Top + s.Height
        Next i
        SbAdd "</Shapes></Shape>"
        Exit Sub
    End If
    gVxDone = gVxDone + 1
    ProgressAt 0.22 + 0.64 * (gVxDone - 1) / gVxTotal, Tx("step.vsdxShape", gVxDone, gVxTotal, ShortName(VxName(s)))
    CheckCancel
    If Not VxNative(s, t, ox, oy) Then VxPicture s, ox, oy
End Sub

Private Function VxNative(ByVal s As Object, ByVal t As Long, ByVal ox As Double, ByVal oy As Double) As Boolean
    Dim w As Double, h As Double, x As String, cells As String, sects As String, txt As String, nm As String
    Select Case t
        Case MSO_AUTOSHAPE, MSO_CALLOUT, MSO_FREEFORM, MSO_LINE, MSO_TEXTBOX, MSO_TEXTEFFECT
        Case MSO_PLACEHOLDER
            If Not PlaceholderIsText(s) Then Exit Function
        Case Else
            Exit Function
    End Select
    On Error GoTo Nope
    w = s.Width
    h = s.Height
    gGeoTextBox = False
    If Not VxGeometry(s, t, w, h) Then Exit Function
    txt = VxText(s, w, h, cells, sects)
    nm = VxName(s)
    x = "<Shape ID='" & (gVxId + 1) & "' NameU='" & nm & "' Name='" & nm & "' Type='Shape' LineStyle='0' FillStyle='0' TextStyle='0'>" & _
        VxXform(s.Left - ox, oy - (s.Top + h), w, h, s.Rotation) & VxStyle(s, t) & cells & VxGradient(s, t) & _
        sects & gGeo & txt & "</Shape>"
    gVxId = gVxId + 1
    SbAdd x
    gVxNative = gVxNative + 1
    VxNative = True
    Exit Function
Nope:
    VxNative = False
End Function

' Objects without a Visio equivalent: an EMF picture of the object, as a Visio foreign shape.
Private Sub VxPicture(ByVal s As Object, ByVal ox As Double, ByVal oy As Double)
    Dim e As String, c As Object, w As Double, h As Double, bw As Double, bh As Double, a As Double
    e = TempFileX("emf")
    On Error Resume Next
    s.Export e, PP_FMT_EMF
    If Err.Number <> 0 Then
        Err.Clear
        Set c = VxTopCopy(s)
        If Not c Is Nothing Then
            c.Export e, PP_FMT_EMF
            c.Delete
        End If
    End If
    On Error GoTo 0
    If Not FileExistsX(e) Then Exit Sub
    w = s.Width
    h = s.Height
    a = s.Rotation * 3.14159265358979 / 180#
    bw = Abs(w * Cos(a)) + Abs(h * Sin(a))          ' the picture shows the rotated object
    bh = Abs(w * Sin(a)) + Abs(h * Cos(a))
    SbAdd VxForeignXml(e, s.Left + w / 2 - bw / 2 - ox, oy - (s.Top + h / 2 + bh / 2), bw, bh, VxName(s))
    gVxPictures = gVxPictures + 1
End Sub

Private Function VxForeignXml(ByVal emfPath As String, ByVal xPt As Double, ByVal yPt As Double, _
                              ByVal wPt As Double, ByVal hPt As Double, ByVal nm As String) As String
    Dim w As String, h As String
    gVxMedia.Add emfPath
    gVxId = gVxId + 1
    w = NumXml(wPt / 72#)
    h = NumXml(hPt / 72#)
    VxForeignXml = "<Shape ID='" & gVxId & "' NameU='" & XmlAttr(nm) & "' Name='" & XmlAttr(nm) & "' Type='Foreign' LineStyle='0' FillStyle='0' TextStyle='0'>" & _
        VxXform(xPt, yPt, wPt, hPt, 0) & VxCell("ResizeMode", "0") & _
        "<Cell N='ImgOffsetX' V='0' F='ImgWidth*0'/><Cell N='ImgOffsetY' V='0' F='ImgHeight*0'/>" & _
        "<Cell N='ImgWidth' V='" & w & "' F='Width*1'/><Cell N='ImgHeight' V='" & h & "' F='Height*1'/>" & _
        VxCell("LinePattern", "0") & VxCell("FillPattern", "0") & _
        "<Section N='Geometry' IX='0'><Cell N='NoFill' V='0'/><Cell N='NoLine' V='1'/><Cell N='NoShow' V='0'/>" & _
        "<Row T='RelMoveTo' IX='1'><Cell N='X' V='0'/><Cell N='Y' V='0'/></Row><Row T='RelLineTo' IX='2'><Cell N='X' V='1'/><Cell N='Y' V='0'/></Row>" & _
        "<Row T='RelLineTo' IX='3'><Cell N='X' V='1'/><Cell N='Y' V='1'/></Row><Row T='RelLineTo' IX='4'><Cell N='X' V='0'/><Cell N='Y' V='1'/></Row>" & _
        "<Row T='RelLineTo' IX='5'><Cell N='X' V='0'/><Cell N='Y' V='0'/></Row></Section>" & _
        "<ForeignData ForeignType='EnhMetaFile' ObjectWidth='" & w & "' ObjectHeight='" & h & "'><Rel r:id='rId" & gVxMedia.Count & "'/></ForeignData></Shape>"
End Function

' Shape frame cells. x, y: bottom-left corner relative to the parent (points); rot: PowerPoint degrees (clockwise).
Private Function VxXform(ByVal xPt As Double, ByVal yPt As Double, ByVal wPt As Double, ByVal hPt As Double, _
                         ByVal rot As Double) As String
    Dim w As Double, h As Double
    w = wPt / 72#
    h = hPt / 72#
    If w < 0.0001 Then w = 0.0001
    If h < 0.0001 Then h = 0.0001
    VxXform = VxCell("PinX", NumXml(xPt / 72# + w / 2)) & VxCell("PinY", NumXml(yPt / 72# + h / 2)) & _
        VxCell("Width", NumXml(w)) & VxCell("Height", NumXml(h)) & _
        "<Cell N='LocPinX' V='" & NumXml(w / 2) & "' F='Width*0.5'/><Cell N='LocPinY' V='" & NumXml(h / 2) & "' F='Height*0.5'/>" & _
        VxCell("Angle", NumXml(-rot * 3.14159265358979 / 180#)) & VxCell("FlipX", "0") & VxCell("FlipY", "0")
End Function

Private Function VxCell(ByVal n As String, ByVal v As String, Optional ByVal ux As String = "") As String
    If Len(ux) > 0 Then
        VxCell = "<Cell N='" & n & "' V='" & v & "' U='" & ux & "'/>"
    Else
        VxCell = "<Cell N='" & n & "' V='" & v & "'/>"
    End If
End Function

Private Function VxName(ByVal s As Object) As String
    Dim nm As String
    On Error Resume Next
    nm = s.Name
    On Error GoTo 0
    If Len(nm) = 0 Then nm = "Shape " & (gVxId + 1)
    VxName = XmlAttr(nm)
End Function

Private Function PlaceholderIsText(ByVal s As Object) As Boolean
    Dim ct As Long
    On Error Resume Next
    ct = -1
    ct = s.PlaceholderFormat.ContainedType
    If ct = -1 Then
        PlaceholderIsText = (s.HasTextFrame = MSO_TRUE)
    Else
        PlaceholderIsText = (ct = MSO_AUTOSHAPE Or ct = MSO_TEXTBOX Or ct = MSO_PLACEHOLDER)
    End If
End Function

' A top-level copy of s on the scratch slide (for shapes inside groups or on other slides).
Private Function VxTopCopy(ByVal s As Object) As Object
    Dim r As Object
    On Error Resume Next
    s.Copy
    DoEvents
    Set r = gVxScratch.Shapes.Paste
    If r Is Nothing Then
        Err.Clear
        DoEvents
        Set r = gVxScratch.Shapes.Paste
    End If
    If Not r Is Nothing Then Set VxTopCopy = r.Item(1)
End Function

' ---- fill, line, arrowheads ---------------------------------------------------
Private Function VxStyle(ByVal s As Object, ByVal t As Long) As String
    Dim r As String, ft As Long, v As Long, c As Long
    On Error Resume Next
    v = MSO_FALSE
    If t <> MSO_LINE Then v = s.Fill.Visible
    If v = MSO_TRUE And Not gGeoOpenOnly Then
        ft = s.Fill.Type
        c = s.Fill.ForeColor.RGB
        If ft = 3 Then c = s.Fill.GradientStops.Item(1).Color.RGB
        If ft = 5 Then c = RGB(255, 255, 255)                     ' slide background fill
        r = VxCell("FillForegnd", HexRGB(c)) & VxCell("FillPattern", "1") & _
            VxCell("FillForegndTrans", NumXml(s.Fill.Transparency))
    Else
        r = VxCell("FillPattern", "0")
    End If
    v = MSO_FALSE
    v = s.Line.Visible
    If v = MSO_TRUE Then
        r = r & VxCell("LineWeight", NumXml(MaxS(s.Line.Weight, 0.1) / 72#), "PT") & _
            VxCell("LineColor", HexRGB(s.Line.ForeColor.RGB)) & _
            VxCell("LinePattern", CStr(VisioDash(s.Line.DashStyle))) & _
            VxCell("LineColorTrans", NumXml(s.Line.Transparency))
        r = r & VxArrow("Begin", s.Line.BeginArrowheadStyle, s.Line.BeginArrowheadLength, s.Line.BeginArrowheadWidth)
        r = r & VxArrow("End", s.Line.EndArrowheadStyle, s.Line.EndArrowheadLength, s.Line.EndArrowheadWidth)
    Else
        r = r & VxCell("LinePattern", "0")
    End If
    VxStyle = r
End Function

Private Function VxArrow(ByVal which As String, ByVal style As Long, ByVal ln As Long, ByVal wd As Long) As String
    Dim a As Long, sz As Long
    Select Case style
        Case 2: a = 4          ' triangle
        Case 3: a = 1          ' open
        Case 4: a = 5          ' stealth
        Case 5: a = 22         ' diamond
        Case 6: a = 10         ' oval
        Case Else: Exit Function
    End Select
    sz = ln
    If wd > sz Then sz = wd
    If sz < 1 Or sz > 3 Then sz = 2
    VxArrow = VxCell(which & "Arrow", CStr(a)) & VxCell(which & "ArrowSize", CStr(sz))
End Function

Private Function VisioDash(ByVal d As Long) As Long
    Select Case d
        Case 2, 11: VisioDash = 3         ' square dot / system dot
        Case 3: VisioDash = 10            ' round dot
        Case 4, 10: VisioDash = 2         ' dash
        Case 5, 12: VisioDash = 4         ' dash dot
        Case 6: VisioDash = 5             ' dash dot dot
        Case 7: VisioDash = 9             ' long dash
        Case 8: VisioDash = 7             ' long dash dot
        Case Else: VisioDash = 1
    End Select
End Function

' Linear gradient fill (Visio 2013+); the solid colour above remains the fallback.
Private Function VxGradient(ByVal s As Object, ByVal t As Long) As String
    Dim gs As Object, i As Long, rows As String, ang As Double
    On Error GoTo Done
    If t = MSO_LINE Or gGeoOpenOnly Then Exit Function
    If s.Fill.Visible <> MSO_TRUE Then Exit Function
    If s.Fill.Type <> 3 Then Exit Function
    Set gs = s.Fill.GradientStops
    If gs.Count < 2 Then Exit Function
    For i = 1 To gs.Count
        rows = rows & "<Row IX='" & (i - 1) & "'>" & VxCell("GradientStopColor", HexRGB(gs.Item(i).Color.RGB)) & _
               VxCell("GradientStopColorTrans", NumXml(gs.Item(i).Transparency)) & _
               VxCell("GradientStopPosition", NumXml(gs.Item(i).Position)) & "</Row>"
    Next i
    ang = 0
    ang = s.Fill.GradientAngle
    VxGradient = VxCell("FillGradientEnabled", "1") & VxCell("FillGradientDir", "0") & _
                 VxCell("FillGradientAngle", NumXml(-ang * 3.14159265358979 / 180#)) & _
                 VxCell("RotateGradientWithShape", "1") & "<Section N='FillGradient'>" & rows & "</Section>"
Done:
End Function

' ---- geometry ------------------------------------------------------------------
' Builds gGeo (Geometry sections). Coordinates passed to the Geo* helpers are PowerPoint-style:
' points, origin at the top-left of the unrotated frame, y downwards; flips are applied here.
Private Function VxGeometry(ByVal s As Object, ByVal t As Long, ByVal w As Double, ByVal h As Double) As Boolean
    Dim isConn As Boolean, fh As Boolean, fv As Boolean, ast As Long
    On Error Resume Next
    fh = (s.HorizontalFlip = MSO_TRUE)
    fv = (s.VerticalFlip = MSO_TRUE)
    isConn = (s.Connector = MSO_TRUE)
    ast = 1
    ast = s.AutoShapeType
    On Error GoTo 0
    GeoStart w, h, fh, fv
    If isConn Then
        VxGeometry = VxConnector(s, w, h)
    ElseIf t = MSO_LINE Then
        GeoBegin True, False
        GeoMove 0, 0
        GeoLine w, h
        GeoEnd False
        gGeoOpenOnly = True
        VxGeometry = True
    ElseIf t = MSO_FREEFORM Then
        If (fh Or fv Or s.Rotation <> 0) Then
            VxGeometry = VxMerged(s, True)
            If Not VxGeometry Then VxGeometry = VxNodes(s, s.Left, s.Top, True)
        Else
            VxGeometry = VxNodes(s, s.Left, s.Top, True)
        End If
    ElseIf t = MSO_TEXTBOX Or t = MSO_PLACEHOLDER Or t = MSO_TEXTEFFECT Then
        GeoRect 0, 0, w, h, 0, 0
        VxGeometry = True
    Else
        VxGeometry = VxPresetGeo(s, ast, w, h)
        If Not VxGeometry Then
            GeoStart w, h, fh, fv
            VxGeometry = VxMerged(s, False)
        End If
    End If
End Function

Private Sub GeoStart(ByVal w As Double, ByVal h As Double, ByVal fh As Boolean, ByVal fv As Boolean)
    gGeoW = w
    gGeoH = h
    gGeoFlipH = fh
    gGeoFlipV = fv
    gGeo = ""
    gGeoIx = 0
    gGeoOpenOnly = False
    gGeoSec = ""
End Sub

Private Sub GeoBegin(ByVal noFill As Boolean, ByVal noLine As Boolean)
    gGeoSec = ""
    gGeoRow = 0
    gGeoNoFill = noFill
    gGeoNoLine = noLine
    gGeoClosed = False
End Sub

Private Sub GeoXY(ByVal ux As Double, ByVal v As Double, ByRef x As Double, ByRef y As Double)
    If gGeoFlipH Then ux = gGeoW - ux
    If gGeoFlipV Then v = gGeoH - v
    x = ux / 72#
    y = (gGeoH - v) / 72#
End Sub

Private Sub GeoMove(ByVal ux As Double, ByVal v As Double)
    Dim x As Double, y As Double
    GeoXY ux, v, x, y
    gGeoRow = gGeoRow + 1
    gGeoSec = gGeoSec & "<Row T='MoveTo' IX='" & gGeoRow & "'>" & VxCell("X", NumXml(x)) & VxCell("Y", NumXml(y)) & "</Row>"
    gGeoX0 = x: gGeoY0 = y
    gGeoLX = x: gGeoLY = y
    gGeoCU = ux: gGeoCV = v: gGeoSU = ux: gGeoSV = v
End Sub

Private Sub GeoLine(ByVal ux As Double, ByVal v As Double)
    Dim x As Double, y As Double
    GeoXY ux, v, x, y
    gGeoRow = gGeoRow + 1
    gGeoSec = gGeoSec & "<Row T='LineTo' IX='" & gGeoRow & "'>" & VxCell("X", NumXml(x)) & VxCell("Y", NumXml(y)) & "</Row>"
    gGeoLX = x: gGeoLY = y
    gGeoCU = ux: gGeoCV = v
End Sub

' Closes the current sub-path (a later GeoMove starts the next one in the same section).
Private Sub GeoClose()
    If gGeoRow = 0 Then Exit Sub
    If Abs(gGeoLX - gGeoX0) > 0.000001 Or Abs(gGeoLY - gGeoY0) > 0.000001 Then
        gGeoRow = gGeoRow + 1
        gGeoSec = gGeoSec & "<Row T='LineTo' IX='" & gGeoRow & "'>" & VxCell("X", NumXml(gGeoX0)) & VxCell("Y", NumXml(gGeoY0)) & "</Row>"
        gGeoLX = gGeoX0: gGeoLY = gGeoY0
    End If
    gGeoCU = gGeoSU: gGeoCV = gGeoSV
    gGeoClosed = True
End Sub

' Quadratic Bezier from the current point, written as the equivalent cubic.
Private Sub GeoQuad(ByVal u1 As Double, ByVal v1 As Double, ByVal ux As Double, ByVal v As Double)
    GeoCubic gGeoCU + 2# / 3# * (u1 - gGeoCU), gGeoCV + 2# / 3# * (v1 - gGeoCV), _
             ux + 2# / 3# * (u1 - ux), v + 2# / 3# * (v1 - v), ux, v
End Sub

' DrawingML arcTo from the current point: radii wr / hr, start and swing angles in 1/60000 degree.
' The angles are visual angles; the point on the ellipse uses the parametric angle
' atan2(wr * sin a, hr * cos a). Written as Visio elliptical arcs of at most 90 degrees each,
' so rounded corners stay true arcs in Visio.
Private Sub GeoArcTo(ByVal wr As Double, ByVal hr As Double, ByVal stA As Double, ByVal swA As Double)
    Dim st As Double, sw As Double, t0 As Double, t1 As Double, dt As Double, cx As Double, cy As Double
    Dim n As Long, i As Long, a As Double, b As Double, m As Double
    If wr <= 0.0001 Or hr <= 0.0001 Or Abs(swA) < 1 Then Exit Sub
    st = stA * PDEG
    sw = swA * PDEG
    t0 = Atan2X(wr * Sin(st), hr * Cos(st))
    t1 = Atan2X(wr * Sin(st + sw), hr * Cos(st + sw))
    If Abs(sw) >= PI2 - 0.0000001 Then
        dt = Sgn(sw) * PI2
    Else
        dt = t1 - t0
        If sw > 0 Then
            Do While dt <= 0: dt = dt + PI2: Loop
        Else
            Do While dt >= 0: dt = dt - PI2: Loop
        End If
    End If
    cx = gGeoCU - wr * Cos(t0)
    cy = gGeoCV - hr * Sin(t0)
    n = Int(Abs(dt) / (PI2 / 4#) - 0.000000001) + 1
    For i = 1 To n
        a = t0 + dt * (i - 1) / n
        b = t0 + dt * i / n
        m = (a + b) / 2#
        GeoEArc cx + wr * Cos(m), cy + hr * Sin(m), cx + wr * Cos(b), cy + hr * Sin(b), wr, hr
    Next i
End Sub

' One Visio EllipticalArcTo row: end point, a point on the arc, axis angle and axis ratio.
Private Sub GeoEArc(ByVal mu As Double, ByVal mv As Double, ByVal eu As Double, ByVal ev As Double, _
                    ByVal rx As Double, ByVal ry As Double)
    Dim mx As Double, my As Double, x As Double, y As Double, c As String, d As Double
    GeoXY mu, mv, mx, my
    GeoXY eu, ev, x, y
    If rx >= ry Then
        c = "0": d = rx / ry
    Else
        c = NumXml(PI2 / 4#): d = ry / rx
    End If
    gGeoRow = gGeoRow + 1
    gGeoSec = gGeoSec & "<Row T='EllipticalArcTo' IX='" & gGeoRow & "'>" & VxCell("X", NumXml(x)) & VxCell("Y", NumXml(y)) & _
              VxCell("A", NumXml(mx)) & VxCell("B", NumXml(my)) & VxCell("C", c) & VxCell("D", NumXml(d)) & "</Row>"
    gGeoLX = x: gGeoLY = y
    gGeoCU = eu: gGeoCV = ev
End Sub

Private Function Atan2X(ByVal y As Double, ByVal x As Double) As Double
    If x > 0 Then
        Atan2X = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then Atan2X = Atn(y / x) + PI2 / 2# Else Atan2X = Atn(y / x) - PI2 / 2#
    ElseIf y > 0 Then
        Atan2X = PI2 / 4#
    ElseIf y < 0 Then
        Atan2X = -PI2 / 4#
    End If
End Function

Private Sub GeoCubic(ByVal u1 As Double, ByVal v1 As Double, ByVal u2 As Double, ByVal v2 As Double, _
                     ByVal ux As Double, ByVal v As Double)
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, x As Double, y As Double, f As String
    GeoXY u1, v1, x1, y1
    GeoXY u2, v2, x2, y2
    GeoXY ux, v, x, y
    f = "NURBS(1,3,1,1," & NumXml(x1) & "," & NumXml(y1) & ",0,1," & NumXml(x2) & "," & NumXml(y2) & ",0,1)"
    gGeoRow = gGeoRow + 1
    gGeoSec = gGeoSec & "<Row T='NURBSTo' IX='" & gGeoRow & "'>" & VxCell("X", NumXml(x)) & VxCell("Y", NumXml(y)) & _
              VxCell("A", "1") & VxCell("B", "1") & VxCell("C", "0") & VxCell("D", "1") & _
              "<Cell N='E' V='" & f & "' F='" & f & "'/></Row>"
    gGeoLX = x: gGeoLY = y
    gGeoCU = ux: gGeoCV = v
End Sub

Private Sub GeoEllipse(ByVal cu As Double, ByVal cv As Double, ByVal ru As Double, ByVal rv As Double)
    Dim x As Double, y As Double
    GeoBegin False, False
    GeoXY cu, cv, x, y
    gGeoRow = 1
    gGeoSec = "<Row T='Ellipse' IX='1'>" & VxCell("X", NumXml(x)) & VxCell("Y", NumXml(y)) & _
              VxCell("A", NumXml(x + ru / 72#)) & VxCell("B", NumXml(y)) & VxCell("C", NumXml(x)) & _
              VxCell("D", NumXml(y + rv / 72#)) & "</Row>"
    gGeoX0 = x: gGeoY0 = y: gGeoLX = x: gGeoLY = y
    GeoEnd True
End Sub

Private Sub GeoEnd(ByVal closed As Boolean)
    Dim nf As Long, nl As Long
    If gGeoRow = 0 Then Exit Sub
    If closed And (Abs(gGeoLX - gGeoX0) > 0.000001 Or Abs(gGeoLY - gGeoY0) > 0.000001) Then
        gGeoRow = gGeoRow + 1
        gGeoSec = gGeoSec & "<Row T='LineTo' IX='" & gGeoRow & "'>" & VxCell("X", NumXml(gGeoX0)) & VxCell("Y", NumXml(gGeoY0)) & "</Row>"
    End If
    If gGeoNoFill Or Not closed Then nf = 1
    If gGeoNoLine Then nl = 1
    gGeo = gGeo & "<Section N='Geometry' IX='" & gGeoIx & "'>" & VxCell("NoFill", CStr(nf)) & VxCell("NoLine", CStr(nl)) & _
           VxCell("NoShow", "0") & gGeoSec & "</Section>"
    gGeoIx = gGeoIx + 1
    gGeoSec = ""
    gGeoRow = 0
End Sub

' Closed polygon from a flat list u1, v1, u2, v2, ...
Private Sub GeoPoly(ParamArray c() As Variant)
    Dim i As Long
    GeoBegin False, False
    GeoMove CDbl(c(0)), CDbl(c(1))
    For i = 2 To UBound(c) - 1 Step 2
        GeoLine CDbl(c(i)), CDbl(c(i + 1))
    Next i
    GeoEnd True
End Sub

' Rectangle with optional elliptical corners (rx, ry).
Private Sub GeoRect(ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double, _
                    ByVal rx As Double, ByVal ry As Double)
    Dim kx As Double, ky As Double, r As Double, b As Double
    r = l + w
    b = t + h
    If rx > w / 2 Then rx = w / 2
    If ry > h / 2 Then ry = h / 2
    GeoBegin False, False
    If rx <= 0 Or ry <= 0 Then
        GeoMove l, t: GeoLine r, t: GeoLine r, b: GeoLine l, b
        GeoEnd True
        Exit Sub
    End If
    kx = rx * KAPPA
    ky = ry * KAPPA
    GeoMove l + rx, t
    GeoLine r - rx, t
    GeoCubic r - rx + kx, t, r, t + ry - ky, r, t + ry
    GeoLine r, b - ry
    GeoCubic r, b - ry + ky, r - rx + kx, b, r - rx, b
    GeoLine l + rx, b
    GeoCubic l + rx - kx, b, l, b - ry + ky, l, b - ry
    GeoLine l, t + ry
    GeoCubic l, t + ry - ky, l + rx - kx, t, l + rx, t
    GeoEnd True
End Sub

Private Function Adj(ByVal s As Object, ByVal i As Long, ByVal defVal As Double) As Double
    Adj = defVal
    On Error Resume Next
    If s.Adjustments.Count >= i Then Adj = s.Adjustments.Item(i)
End Function

' ---- preset geometry engine ---------------------------------------------------
' Every PowerPoint AutoShape is drawn from its DrawingML preset definition (ECMA-376
' presetShapeDefinitions, compiled by build/make_presets.py into PresetSource): the guide
' formulas are evaluated with the shape's size and adjustment values and the paths are
' written as Visio geometry - straight lines, Bezier curves and true elliptical arcs.
Private Function VxPresetGeo(ByVal s As Object, ByVal ast As Long, ByVal w As Double, ByVal h As Double) As Boolean
    Dim src As String, sec() As String, part() As String, i As Long, n As Long, d As String
    Dim v As Double, a As Double, isAng As Boolean, anyFill As Boolean, rc() As String
    src = PresetSource(ast)
    If Len(src) = 0 Then Exit Function
    sec = Split(src, "|")
    If UBound(sec) < 3 Then Exit Function
    gPw = w
    gPh = h
    ReDim gPa(0 To 0)
    If Len(sec(0)) > 0 Then
        part = Split(sec(0), ",")
        ReDim gPa(0 To UBound(part) + 1)
        n = 0
        On Error Resume Next
        n = s.Adjustments.Count
        On Error GoTo 0
        For i = 0 To UBound(part)
            d = part(i)
            isAng = (Left$(d, 1) = "!")
            If isAng Then d = Mid$(d, 2)
            v = Val(d)
            If n >= i + 1 Then
                ' PowerPoint reports linear adjustments as fractions and angles in degrees
                On Error Resume Next
                a = s.Adjustments.Item(i + 1)
                If Err.Number = 0 Then
                    If isAng Then
                        If Abs(a) > 400 Then v = a Else v = a * 60000#
                        If v < 0 And Val(d) >= 0 Then v = v + 21600000#
                    Else
                        If Abs(a) > 10 Then v = a Else v = a * 100000#
                    End If
                End If
                Err.Clear
                On Error GoTo 0
            End If
            gPa(i + 1) = v
        Next i
    End If
    If Len(sec(1)) > 0 Then
        part = Split(sec(1), ";")
        ReDim gPv(0 To UBound(part))
        For i = 0 To UBound(part)
            gPv(i) = PresetEval(part(i))
        Next i
    Else
        ReDim gPv(0 To 0)
    End If
    part = Split(sec(2), "~")
    For i = 0 To UBound(part)
        If PresetPath(part(i)) Then anyFill = True
    Next i
    If Not anyFill Then gGeoOpenOnly = True
    rc = Split(sec(3), ",")
    If Abs(PresetArg(rc(0))) > 0.5 Or Abs(PresetArg(rc(1))) > 0.5 Or Abs(PresetArg(rc(2)) - w) > 0.5 Or _
       Abs(PresetArg(rc(3)) - h) > 0.5 Then gGeoTextBox = True      ' text area is inset
    VxPresetGeo = (Len(gGeo) > 0)
End Function

' One path: "<fill f|n>,<stroke 1|0>,<path w>,<path h>:<commands>". Returns True when it is filled.
Private Function PresetPath(ByVal p As String) As Boolean
    Dim k As Long, hd() As String, cmds() As String, c() As String, i As Long, sx As Double, sy As Double
    Dim filled As Boolean
    k = InStr(p, ":")
    If k = 0 Then Exit Function
    hd = Split(Left$(p, k - 1), ",")
    cmds = Split(Mid$(p, k + 1), ";")
    filled = (hd(0) = "f")
    sx = 1: sy = 1
    If Val(hd(2)) > 0 Then sx = gPw / Val(hd(2))
    If Val(hd(3)) > 0 Then sy = gPh / Val(hd(3))
    GeoBegin Not filled, (hd(1) = "0")
    For i = 0 To UBound(cmds)
        c = Split(cmds(i), ",")
        Select Case c(0)
            Case "M": GeoMove PresetArg(c(1)) * sx, PresetArg(c(2)) * sy
            Case "L": GeoLine PresetArg(c(1)) * sx, PresetArg(c(2)) * sy
            Case "C": GeoCubic PresetArg(c(1)) * sx, PresetArg(c(2)) * sy, PresetArg(c(3)) * sx, PresetArg(c(4)) * sy, _
                               PresetArg(c(5)) * sx, PresetArg(c(6)) * sy
            Case "Q": GeoQuad PresetArg(c(1)) * sx, PresetArg(c(2)) * sy, PresetArg(c(3)) * sx, PresetArg(c(4)) * sy
            Case "A": GeoArcTo PresetArg(c(1)) * sx, PresetArg(c(2)) * sy, PresetArg(c(3)), PresetArg(c(4))
            Case "Z": GeoClose
        End Select
    Next i
    PresetPath = filled And gGeoClosed
    GeoEnd gGeoClosed
End Function

Private Function PresetArg(ByVal tok As String) As Double
    Dim c As String
    c = Left$(tok, 1)
    If c = "$" Then
        PresetArg = gPv(CLng(Mid$(tok, 2)))
    ElseIf c = "@" Then
        PresetArg = gPa(CLng(Mid$(tok, 2)))
    ElseIf c = "-" Or (c >= "0" And c <= "9") Then
        PresetArg = Val(tok)
    Else
        PresetArg = PresetBuiltin(tok)
    End If
End Function

Private Function PresetBuiltin(ByVal tok As String) As Double
    Dim ss As Double
    ss = gPw
    If gPh < ss Then ss = gPh
    Select Case tok
        Case "w", "r": PresetBuiltin = gPw
        Case "h", "b": PresetBuiltin = gPh
        Case "l", "t": PresetBuiltin = 0
        Case "hc": PresetBuiltin = gPw / 2
        Case "vc": PresetBuiltin = gPh / 2
        Case "ss": PresetBuiltin = ss
        Case "ls": If gPw > gPh Then PresetBuiltin = gPw Else PresetBuiltin = gPh
        Case "cd2": PresetBuiltin = 10800000#
        Case "cd4": PresetBuiltin = 5400000#
        Case "cd8": PresetBuiltin = 2700000#
        Case "k34": PresetBuiltin = 16200000#
        Case "k38": PresetBuiltin = 8100000#
        Case "k58": PresetBuiltin = 13500000#
        Case "k78": PresetBuiltin = 18900000#
        Case Else
            If Left$(tok, 3) = "ssd" Then
                PresetBuiltin = ss / Val(Mid$(tok, 4))
            ElseIf Left$(tok, 2) = "wd" Then
                PresetBuiltin = gPw / Val(Mid$(tok, 3))
            ElseIf Left$(tok, 2) = "hd" Then
                PresetBuiltin = gPh / Val(Mid$(tok, 3))
            End If
    End Select
End Function

' Guide formula "<op>,<x>,<y>,<z>" (DrawingML operators, angles in 1/60000 degree).
Private Function PresetEval(ByVal f As String) As Double
    Dim t() As String, x As Double, y As Double, z As Double, n As Long
    t = Split(f, ",")
    n = UBound(t)
    If n >= 1 Then x = PresetArg(t(1))
    If n >= 2 Then y = PresetArg(t(2))
    If n >= 3 Then z = PresetArg(t(3))
    Select Case t(0)
        Case "m": If z <> 0 Then PresetEval = x * y / z
        Case "a": PresetEval = x + y - z
        Case "d": If z <> 0 Then PresetEval = (x + y) / z
        Case "q": If x > 0 Then PresetEval = y Else PresetEval = z
        Case "b": PresetEval = Abs(x)
        Case "t": PresetEval = Atan2X(y, x) / PDEG
        Case "c": PresetEval = x * Cos(Atan2X(z, y))
        Case "o": PresetEval = x * Cos(y * PDEG)
        Case "x": If x > y Then PresetEval = x Else PresetEval = y
        Case "n": If x < y Then PresetEval = x Else PresetEval = y
        Case "h": PresetEval = Sqr(x * x + y * y + z * z)
        Case "p"
            If y < x Then
                PresetEval = x
            ElseIf y > z Then
                PresetEval = z
            Else
                PresetEval = y
            End If
        Case "s": PresetEval = x * Sin(Atan2X(z, y))
        Case "i": PresetEval = x * Sin(y * PDEG)
        Case "r": If x > 0 Then PresetEval = Sqr(x)
        Case "g": PresetEval = x * Tan(y * PDEG)
        Case "v": PresetEval = x
    End Select
End Function

' Connectors: straight, elbow (bentConnector2-5) and curved, from their adjustment values.
Private Function VxConnector(ByVal s As Object, ByVal w As Double, ByVal h As Double) As Boolean
    Dim ct As Long, n As Long, a1 As Double, a2 As Double, a3 As Double, x1 As Double, x2 As Double, x3 As Double
    On Error Resume Next
    ct = 1
    ct = s.ConnectorFormat.Type
    n = s.Adjustments.Count
    On Error GoTo 0
    gGeoOpenOnly = True
    GeoBegin True, False
    GeoMove 0, 0
    If ct = 2 Then
        a1 = Adj(s, 1, 0.5): a2 = Adj(s, 2, 0.5): a3 = Adj(s, 3, 0.5)
        Select Case n
            Case 0
                GeoLine w, 0: GeoLine w, h
            Case 1
                GeoLine a1 * w, 0: GeoLine a1 * w, h: GeoLine w, h
            Case 2
                GeoLine a1 * w, 0: GeoLine a1 * w, a2 * h: GeoLine w, a2 * h: GeoLine w, h
            Case Else
                GeoLine a1 * w, 0: GeoLine a1 * w, a2 * h: GeoLine a3 * w, a2 * h: GeoLine a3 * w, h: GeoLine w, h
        End Select
    ElseIf ct = 3 Then
        If n = 0 Then
            GeoCubic w / 2, 0, w, h / 2, w, h
        Else
            x2 = Adj(s, 1, 0.5) * w
            x1 = x2 / 2
            x3 = (w + x2) / 2
            GeoCubic x1, 0, x2, h / 4, x2, h / 2
            GeoCubic x2, h * 3 / 4, x3, h, w, h
        End If
    Else
        GeoLine w, h
    End If
    GeoEnd False
    VxConnector = True
End Function

' Freeform nodes -> geometry. (l, t): origin of the frame the node coordinates are measured from.
' PowerPoint stores a Bezier segment as two control nodes followed by its end vertex; the
' SegmentType of a vertex describes the segment that follows it (ShapeNodes.SetSegmentType),
' control nodes report msoSegmentCurve. Older builds were seen to flag the segment that ends at
' a node instead, so both readings are scored by tangent continuity and the better one is used.
Private Function VxNodes(ByVal s As Object, ByVal l As Double, ByVal t As Double, ByVal useFill As Boolean) As Boolean
    Dim n As Long, i As Long, p As Variant, filled As Boolean, started As Boolean, after As Boolean
    Dim px() As Double, py() As Double, sg() As Long, scoreF As Long, scoreA As Long
    Dim su As Double, sv As Double, ux As Double, v As Double
    On Error GoTo Nope
    n = s.Nodes.Count
    If n < 2 Then Exit Function
    ReDim px(1 To n): ReDim py(1 To n): ReDim sg(1 To n)
    For i = 1 To n
        p = s.Nodes.Item(i).Points
        px(i) = p(1, 1) - l
        py(i) = p(1, 2) - t
        sg(i) = s.Nodes.Item(i).SegmentType
    Next i
    filled = False
    If useFill Then filled = (s.Fill.Visible = MSO_TRUE)
    For i = 2 To n - 1
        If sg(i) = 1 And sg(i - 1) = 0 Then
            If Collinear3(px(i - 1), py(i - 1), px(i), py(i), px(i + 1), py(i + 1)) Then scoreF = scoreF + 1
            If i >= 3 Then
                If Collinear3(px(i - 2), py(i - 2), px(i - 1), py(i - 1), px(i), py(i)) Then scoreA = scoreA + 1
            End If
        End If
    Next i
    after = (scoreA > scoreF)
    i = 1
    Do While i <= n
        If Not started Then
            GeoBegin False, False
            GeoMove px(i), py(i)
            su = px(i): sv = py(i)
            ux = su: v = sv
            started = True
            i = i + 1
        Else
            If after Then
                ' flag on the node where a segment ends: node i is the first node of the next segment
                If sg(i) = 1 And i + 2 <= n Then
                    GeoCubic px(i), py(i), px(i + 1), py(i + 1), px(i + 2), py(i + 2)
                    ux = px(i + 2): v = py(i + 2)
                    i = i + 3
                Else
                    GeoLine px(i), py(i)
                    ux = px(i): v = py(i)
                    i = i + 1
                End If
            Else
                ' flag on the vertex where a segment starts (node i - 1 is the current vertex)
                If sg(i - 1) = 1 And i + 2 <= n Then
                    GeoCubic px(i), py(i), px(i + 1), py(i + 1), px(i + 2), py(i + 2)
                    ux = px(i + 2): v = py(i + 2)
                    i = i + 3
                Else
                    GeoLine px(i), py(i)
                    ux = px(i): v = py(i)
                    i = i + 1
                End If
            End If
            If i <= n And gGeoRow > 2 And Abs(ux - su) < 0.01 And Abs(v - sv) < 0.01 Then
                GeoEnd True                          ' sub-path closed: the next node starts a new one
                started = False
            End If
        End If
    Loop
    If started Then GeoEnd filled Or (Abs(ux - su) < 0.01 And Abs(v - sv) < 0.01)
    If Not filled Then gGeoOpenOnly = True
    VxNodes = (Len(gGeo) > 0)
    Exit Function
Nope:
    VxNodes = False
End Function

' True when b lies on the straight continuation from a to c (smooth join).
Private Function Collinear3(ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double, _
                            ByVal cx As Double, ByVal cy As Double) As Boolean
    Dim ux As Double, uy As Double, vx As Double, vy As Double, lu As Double, lv As Double
    ux = bx - ax: uy = by - ay
    vx = cx - bx: vy = cy - by
    lu = Sqr(ux * ux + uy * uy)
    lv = Sqr(vx * vx + vy * vy)
    If lu < 0.01 Or lv < 0.01 Then Exit Function
    Collinear3 = (Abs(ux * vy - uy * vx) / (lu * lv) < 0.02) And (ux * vx + uy * vy > 0)
End Function

' Any other closed shape: PowerPoint 2013+ "Merge Shapes - Union" of an unrotated copy with
' its duplicate gives a freeform with the shape's exact outline; its nodes are the geometry.
Private Function VxMerged(ByVal s As Object, ByVal keepOpen As Boolean) As Boolean
    Dim a As Object, b As Object, m As Object, k As Long, nmA As String, nmB As String
    Dim l As Double, t As Double, i As Long, names As String
    On Error GoTo Nope
    If gVxScratch Is Nothing Then Exit Function
    Set a = VxTopCopy(s)
    If a Is Nothing Then Exit Function
    a.Rotation = 0
    gVxCopy = gVxCopy + 1
    nmA = "vsA" & gVxCopy
    nmB = "vsB" & gVxCopy
    a.Name = nmA
    l = a.Left
    t = a.Top
    Set b = a.Duplicate.Item(1)
    b.Name = nmB
    b.Left = l
    b.Top = t
    For i = 1 To gVxScratch.Shapes.Count
        names = names & "|" & gVxScratch.Shapes.Item(i).Name & "|"
    Next i
    gVxScratch.Shapes.Range(Array(nmA, nmB)).MergeShapes 1, a      ' msoMergeUnion
    For i = gVxScratch.Shapes.Count To 1 Step -1
        If InStr(names, "|" & gVxScratch.Shapes.Item(i).Name & "|") = 0 Or gVxScratch.Shapes.Item(i).Name = nmA Then
            Set m = gVxScratch.Shapes.Item(i)
            Exit For
        End If
    Next i
    If m Is Nothing Then GoTo NopeCleanup
    gGeoFlipH = False                                ' flips are part of the merged outline
    gGeoFlipV = False
    VxMerged = VxNodes(m, l, t, Not keepOpen)
    m.Delete
    Exit Function
Nope:
    Resume NopeCleanup
NopeCleanup:
    On Error Resume Next
    If Not a Is Nothing Then a.Delete
    If Not b Is Nothing Then b.Delete
    VxMerged = False
End Function

' ---- text ----------------------------------------------------------------------
' Returns the <Text> element; cells and sects receive the text block cells and the
' Character / Paragraph sections (one row per run / paragraph, in text order).
Private Function VxText(ByVal s As Object, ByVal w As Double, ByVal h As Double, _
                        ByRef cells As String, ByRef sects As String) As String
    Dim tf As Object, tr As Object, para As Object, rn As Object
    Dim pi As Long, ri As Long, nc As Long, body As String, chars As String, paras As String
    Dim txt As String, fnt As String, fea As String, st As Long, al As Long, va As Long
    Dim sp As String, sb As Double, sa As Double, nowrap As Boolean, orient As Long, ang As Double
    Dim bx As Double, by As Double, bw As Double, bh As Double, tw As Double, th As Double, px As Double
    Dim bul As String

    cells = ""
    sects = ""
    On Error GoTo NoText
    If s.HasTextFrame <> MSO_TRUE Then Exit Function
    Set tf = s.TextFrame
    If tf.HasText <> MSO_TRUE Then Exit Function
    Set tr = tf.TextRange
    For pi = 1 To tr.Paragraphs.Count
        Set para = tr.Paragraphs(pi)
        al = 0
        Select Case para.ParagraphFormat.Alignment
            Case 2: al = 1
            Case 3: al = 2
            Case 4, 7: al = 3
            Case 5, 6: al = 4
        End Select
        sp = "-1.2"
        On Error Resume Next
        If para.ParagraphFormat.LineRuleWithin = MSO_TRUE Then
            sp = NumXml(-1.2 * para.ParagraphFormat.SpaceWithin)
        Else
            sp = NumXml(para.ParagraphFormat.SpaceWithin / 72#)
        End If
        sb = 0: sa = 0
        sb = para.ParagraphFormat.SpaceBefore
        sa = para.ParagraphFormat.SpaceAfter
        If para.ParagraphFormat.LineRuleBefore = MSO_TRUE Then sb = sb * para.Font.Size * 1.2
        If para.ParagraphFormat.LineRuleAfter = MSO_TRUE Then sa = sa * para.Font.Size * 1.2
        bul = ""
        If para.ParagraphFormat.Bullet.Visible = MSO_TRUE Then
            If para.ParagraphFormat.Bullet.Type = 2 Then
                bul = CStr(para.ParagraphFormat.Bullet.Number) & ". "
            Else
                bul = ChrW(para.ParagraphFormat.Bullet.Character) & " "
            End If
            If Len(bul) = 0 Or bul = ChrW(0) & " " Then bul = ChrW(&H2022) & " "
        End If
        On Error GoTo NoText
        If Left$(sp, 1) = "-" Then
            paras = paras & "<Row IX='" & (pi - 1) & "'>" & VxCell("HorzAlign", CStr(al)) & VxCell("SpLine", sp)
        Else
            paras = paras & "<Row IX='" & (pi - 1) & "'>" & VxCell("HorzAlign", CStr(al)) & VxCell("SpLine", sp, "PT")
        End If
        paras = paras & VxCell("SpBefore", NumXml(sb / 72#), "PT") & VxCell("SpAfter", NumXml(sa / 72#), "PT") & _
                VxCell("Bullet", "0") & "</Row>"
        body = body & "<pp IX='" & (pi - 1) & "'/>"
        For ri = 1 To para.Runs.Count
            Set rn = para.Runs(ri)
            txt = rn.Text
            Do While Len(txt) > 0 And (Right$(txt, 1) = vbCr Or Right$(txt, 1) = vbLf)
                txt = Left$(txt, Len(txt) - 1)
            Loop
            If ri = 1 Then txt = bul & txt
            If Len(txt) > 0 Then
                fnt = ThemeFontName(rn.Font.Name, False)
                fea = fnt
                On Error Resume Next
                fea = ThemeFontName(rn.Font.NameFarEast, True)
                If Len(fea) = 0 Then fea = fnt
                On Error GoTo NoText
                st = 0
                If rn.Font.Bold = MSO_TRUE Then st = st Or 1
                If rn.Font.Italic = MSO_TRUE Then st = st Or 2
                If rn.Font.Underline = MSO_TRUE Then st = st Or 4
                chars = chars & "<Row IX='" & nc & "'>" & VxCell("Font", XmlAttr(fnt)) & VxCell("AsianFont", XmlAttr(fea)) & _
                        VxCell("Color", HexRGB(rn.Font.Color.RGB)) & VxCell("Style", CStr(st)) & _
                        VxCell("Size", NumXml(rn.Font.Size / 72#), "PT") & VxCell("Pos", CStr(SuperSub(rn))) & _
                        VxCell("LangID", LangOf(txt)) & "</Row>"
                body = body & "<cp IX='" & nc & "'/>" & XmlText(Replace(txt, Chr$(11), ChrW(&H2028)))
                nc = nc + 1
            End If
        Next ri
        If pi < tr.Paragraphs.Count Then body = body & vbLf
    Next pi
    If nc = 0 Then Exit Function

    va = 1
    Select Case tf.VerticalAnchor
        Case 1, 2: va = 0
        Case 4, 5: va = 2
    End Select
    cells = VxCell("LeftMargin", NumXml(tf.MarginLeft / 72#), "PT") & VxCell("RightMargin", NumXml(tf.MarginRight / 72#), "PT") & _
            VxCell("TopMargin", NumXml(tf.MarginTop / 72#), "PT") & VxCell("BottomMargin", NumXml(tf.MarginBottom / 72#), "PT") & _
            VxCell("VerticalAlign", CStr(va))
    On Error Resume Next
    nowrap = (tf.WordWrap = MSO_FALSE)
    orient = 1
    orient = tf.Orientation
    On Error GoTo NoText
    If orient = 4 Or orient = 5 Then cells = cells & VxCell("TextDirection", "1")
    If orient = 2 Then ang = 1.5707963267949
    If orient = 3 Then ang = -1.5707963267949

    bx = 0: by = 0: bw = w: bh = h
    If gGeoTextBox And s.Rotation = 0 And ang = 0 Then
        On Error Resume Next
        bw = 0
        bx = tr.BoundLeft - s.Left
        by = tr.BoundTop - s.Top
        bw = tr.BoundWidth
        bh = tr.BoundHeight
        On Error GoTo NoText
        If bw <= 0 Or bh <= 0 Then
            bx = 0: by = 0: bw = w: bh = h
            gGeoTextBox = False
        Else
            bx = bx - bw * 0.05 - tf.MarginLeft
            bw = bw * 1.1 + tf.MarginLeft + tf.MarginRight
            by = by - tf.MarginTop
            bh = bh + tf.MarginTop + tf.MarginBottom
        End If
    End If
    If gGeoTextBox Or ang <> 0 Or nowrap Then
        tw = bw: th = bh
        If ang <> 0 Then tw = bh: th = bw
        If nowrap Then tw = MaxS(CSng(tw * 2), CSng(tw + 72))
        px = bx + bw / 2
        If nowrap And Not gGeoTextBox Then
            If al = 0 Then px = bx + tw / 2
            If al = 2 Then px = bx + bw - tw / 2
        End If
        cells = cells & VxCell("TxtPinX", NumXml(px / 72#)) & VxCell("TxtPinY", NumXml((h - by - bh / 2) / 72#)) & _
                VxCell("TxtWidth", NumXml(tw / 72#)) & VxCell("TxtHeight", NumXml(th / 72#)) & _
                VxCell("TxtLocPinX", NumXml(tw / 144#)) & VxCell("TxtLocPinY", NumXml(th / 144#)) & _
                VxCell("TxtAngle", NumXml(ang))
    End If
    sects = "<Section N='Character'>" & chars & "</Section><Section N='Paragraph'>" & paras & "</Section>"
    VxText = "<Text>" & body & "</Text>"
    Exit Function
NoText:
    cells = ""
    sects = ""
    VxText = ""
End Function

Private Function SuperSub(ByVal rn As Object) As Long
    Dim b As Double
    On Error Resume Next
    b = rn.Font.BaselineOffset
    If b > 0 Then SuperSub = 1
    If b < 0 Then SuperSub = 2
End Function

Private Function LangOf(ByVal s As String) As String
    Dim i As Long, c As Long
    LangOf = "en-US"
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c >= &H2E80& And c <= &H9FFF& Then LangOf = "zh-CN": Exit Function
    Next i
End Function

' Theme font placeholders (+mn-lt, +mj-ea, ...) -> the real font name of the theme.
Private Function ThemeFontName(ByVal nm As String, ByVal asian As Boolean) As String
    Dim scheme As Object, f As Object
    ThemeFontName = nm
    If Left$(nm, 1) <> "+" Then Exit Function
    On Error Resume Next
    Set scheme = gVxPres.SlideMaster.Theme.ThemeFontScheme
    If InStr(nm, "mj") > 0 Then Set f = scheme.MajorFont Else Set f = scheme.MinorFont
    If InStr(nm, "-ea") > 0 Or asian Then
        ThemeFontName = f.Item(2).Name
    Else
        ThemeFontName = f.Item(1).Name
    End If
    If Len(ThemeFontName) = 0 Or Left$(ThemeFontName, 1) = "+" Then ThemeFontName = "Calibri"
End Function

' ---- small helpers -------------------------------------------------------------
Private Sub SbReset()
    ReDim gVxSb(0 To 255)
    gVxSbN = 0
End Sub

Private Sub SbAdd(ByVal s As String)
    If gVxSbN > UBound(gVxSb) Then ReDim Preserve gVxSb(0 To UBound(gVxSb) * 2 + 1)
    gVxSb(gVxSbN) = s
    gVxSbN = gVxSbN + 1
End Sub

Private Function SbText() As String
    If gVxSbN = 0 Then Exit Function
    ReDim Preserve gVxSb(0 To gVxSbN - 1)
    SbText = Join(gVxSb, "")
End Function

Private Function HexRGB(ByVal c As Long) As String
    HexRGB = "#" & Right$("0" & Hex$(c And &HFF&), 2) & Right$("0" & Hex$((c \ &H100&) And &HFF&), 2) & _
             Right$("0" & Hex$((c \ &H10000) And &HFF&), 2)
End Function

Private Function XmlAttr(ByVal s As String) As String
    XmlAttr = Replace(Replace(XmlText(s), "'", "&apos;"), """", "&quot;")
End Function

' Escapes text for XML and drops characters XML 1.0 does not allow.
Private Function XmlText(ByVal s As String) As String
    Dim i As Long, c As Long, r As String
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c < 32 And c <> 9 And c <> 10 Then
            ' skip
        ElseIf c = 38 Then
            r = r & "&amp;"
        ElseIf c = 60 Then
            r = r & "&lt;"
        ElseIf c = 62 Then
            r = r & "&gt;"
        ElseIf c = &HFFFE& Or c = &HFFFF& Then
            ' skip
        Else
            r = r & Mid$(s, i, 1)
        End If
    Next i
    XmlText = r
End Function

' UTF-8 bytes of a VBA (UTF-16) string, packed into a byte string for WriteBytesX.
Private Function Utf8Bytes(ByVal s As String) As String
    Dim b() As Byte, i As Long, n As Long, c As Long, c2 As Long, k As Long
    n = Len(s)
    If n = 0 Then Exit Function
    ReDim b(0 To n * 3)
    For i = 1 To n
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c >= &HD800& And c <= &HDBFF& And i < n Then
            c2 = AscW(Mid$(s, i + 1, 1)) And &HFFFF&
            If c2 >= &HDC00& And c2 <= &HDFFF& Then
                c = &H10000 + (c - &HD800&) * &H400& + (c2 - &HDC00&)
                i = i + 1
            End If
        End If
        If c < &H80& Then
            b(k) = c: k = k + 1
        ElseIf c < &H800& Then
            b(k) = &HC0& Or (c \ &H40&): b(k + 1) = &H80& Or (c And &H3F&): k = k + 2
        ElseIf c < &H10000 Then
            b(k) = &HE0& Or (c \ &H1000&): b(k + 1) = &H80& Or ((c \ &H40&) And &H3F&)
            b(k + 2) = &H80& Or (c And &H3F&): k = k + 3
        Else
            If k + 4 > UBound(b) Then ReDim Preserve b(0 To UBound(b) + 16)
            b(k) = &HF0& Or (c \ &H40000): b(k + 1) = &H80& Or ((c \ &H1000&) And &H3F&)
            b(k + 2) = &H80& Or ((c \ &H40&) And &H3F&): b(k + 3) = &H80& Or (c And &H3F&): k = k + 4
        End If
    Next i
    ReDim Preserve b(0 To k - 1)
    Utf8Bytes = b
End Function

' SVG for PowerPoint versions without SVG export (Windows, PowerPoint 2013 - 2021 / LTSC):
' PowerPoint prints a copy of the drawing (alone on a temporary slide) to XPS, and the
' built-in converter turns the XPS page into SVG - paths, gradients, images and real text.
' The converter runs as a separate hidden process: the progress window shows its steps,
' Esc stops it and it is stopped after SVG_TIMEOUT seconds.
' Returns 0 on success, RUN_CANCELLED, RUN_TIMEOUT or another non-zero value on failure.
Private Function XpsToSvg(ByVal pres As Object, ByVal shp As Object, ByVal dst As String) As Long
#If Mac Then
    XpsToSvg = RUN_FAILED
#Else
    Dim tmp As Object, c As Object, rng As Object, x As String, sc As String, cmd As String, pf As String
    Dim l As Double, t As Double, w As Double, h As Double, i As Long, wasSaved As Long, st As Long
    st = RUN_FAILED
    wasSaved = pres.Saved
    x = TempFileX("xps")
    pf = TempFileX("txt")
    On Error GoTo Failed
    sc = HelperScriptPath()
    If Len(sc) = 0 Then GoTo Cleanup
    ProgressAt 0.24, Tx("step.xpsPrep")
    Set tmp = AddTempSlide(pres, shp)
    On Error Resume Next
    tmp.FollowMasterBackground = MSO_FALSE
    tmp.Background.Fill.Visible = MSO_FALSE
    tmp.DisplayMasterShapes = MSO_FALSE
    On Error GoTo Failed
    Set c = PasteFromClipboard(tmp, shp, False)
    l = 1E+30: t = 1E+30: w = -1E+30: h = -1E+30
    For i = 1 To c.Count
        If c.Item(i).Left < l Then l = c.Item(i).Left
        If c.Item(i).Top < t Then t = c.Item(i).Top
        If c.Item(i).Left + c.Item(i).Width > w Then w = c.Item(i).Left + c.Item(i).Width
        If c.Item(i).Top + c.Item(i).Height > h Then h = c.Item(i).Top + c.Item(i).Height
    Next i
    w = w - l
    h = h - t
    CheckCancel
    ProgressAt 0.3, Tx("step.xpsPrint")
    pres.PrintOptions.Ranges.ClearAll
    Set rng = pres.PrintOptions.Ranges.Add(tmp.SlideIndex, tmp.SlideIndex)
    ' Path, XPS, print intent, no frame, handout order, slides, hidden slides, range, slide range
    pres.ExportAsFixedFormat x, 1, 2, MSO_FALSE, 1, 1, MSO_TRUE, rng, 4
    If Not tmp Is Nothing Then tmp.Delete
    Set tmp = Nothing
    ProgressAt 0.4, Tx("step.xps", Tx("xps.start"))
    cmd = "& ([scriptblock]::Create([IO.File]::ReadAllText('" & PsQ(sc) & "',[Text.Encoding]::UTF8))) " & _
          "-Xps '" & PsQ(x) & "' -Out '" & PsQ(dst) & "' -Frame '" & NumXml(l) & "," & NumXml(t) & "," & _
          NumXml(w) & "," & NumXml(h) & "' -Slide '" & NumXml(pres.PageSetup.SlideWidth) & "," & _
          NumXml(pres.PageSetup.SlideHeight) & "' -Progress '" & PsQ(pf) & "' -MaxSeconds " & (SVG_TIMEOUT - 5)
    DeleteIfExists dst
    st = RunPowerShell(cmd, SVG_TIMEOUT, pf, 0.4, 0.95)
    If st = 3 Then st = RUN_TIMEOUT
    If st = 0 Then
        If Not Fso().FileExists(dst) Then st = RUN_FAILED
    End If
Cleanup:
    On Error Resume Next
    If Not tmp Is Nothing Then tmp.Delete
    pres.PrintOptions.Ranges.ClearAll
    DeleteTempX x
    DeleteTempX pf
    pres.Saved = wasSaved
    On Error GoTo 0
    XpsToSvg = st
    Exit Function
Failed:
    If Err.Number = ERR_CANCEL Then st = RUN_CANCELLED Else st = RUN_FAILED
    Resume Cleanup
#End If
End Function

#If Mac Then
#Else
' Writes the XPS -> SVG converter (embedded in this module) to the temp folder.
Private Function HelperScriptPath() As String
    Dim p As String, b() As Byte
    On Error GoTo Failed
    p = JoinPath(TempDirX(), "xps2svg-" & VS_VERSION & ".ps1")
    b = Utf8Bytes(HelperScript())
    If Fso().FileExists(p) Then
        If Fso().GetFile(p).Size <> UBound(b) + 1 Then Fso().DeleteFile p, True
    End If
    If Not Fso().FileExists(p) Then WriteBytesX p, b
    HelperScriptPath = p
    Exit Function
Failed:
    HelperScriptPath = ""
End Function
#End If


' =====================================================================
'  Temporary slide used to combine shapes and add background / margin / grid
' =====================================================================
Private Function AddTempSlide(ByVal pres As Object, ByVal src As Object) As Object
    Dim lay As Object, sld As Object, i As Long
    On Error Resume Next
    Set lay = src.Parent.CustomLayout              ' same master and theme as the source slide
    If lay Is Nothing Then Set lay = pres.SlideMaster.CustomLayouts.Item(pres.SlideMaster.CustomLayouts.Count)
    On Error GoTo 0
    Set sld = pres.Slides.AddSlide(pres.Slides.Count + 1, lay)
    For i = sld.Shapes.Count To 1 Step -1
        sld.Shapes.Item(i).Delete                  ' drop the empty layout placeholders
    Next i
    Set AddTempSlide = sld
End Function

Private Function BuildComposite(ByVal tmp As Object, ByVal src As Object, ByVal margin As Single) As Object
    Dim s As Object
    Set s = TryComposite(tmp, src, margin, False)
    If s Is Nothing Then
        ' Some objects (e.g. placeholders) cannot be grouped: use a vector EMF picture instead.
        ClearSlide tmp
        Set s = TryComposite(tmp, src, margin, True)
    End If
    If s Is Nothing Then Err.Raise vbObjectError + 513, , Tx("err.group")
    Set BuildComposite = s
End Function

Private Function TryComposite(ByVal tmp As Object, ByVal src As Object, _
                              ByVal margin As Single, ByVal asPicture As Boolean) As Object
    Dim content As Object, rect As Object, bg As Long
    PasteFromClipboard tmp, src, asPicture
    If tmp.Shapes.Count > 1 Then
        Set content = GroupAll(tmp)
        If content Is Nothing Then Exit Function
    Else
        Set content = tmp.Shapes.Item(1)
    End If

    bg = BgTypeIdx()
    If bg <> BG_NONE Or margin > 0 Then
        Set rect = tmp.Shapes.AddShape(MSO_RECTANGLE, content.Left - margin, content.Top - margin, _
                                       content.Width + 2 * margin, content.Height + 2 * margin)
        rect.Line.Visible = MSO_FALSE
        rect.Shadow.Visible = MSO_FALSE
        rect.Fill.Visible = MSO_TRUE
        rect.Fill.Solid
        If bg = BG_NONE Then
            rect.Fill.ForeColor.RGB = RGB(255, 255, 255)
            rect.Fill.Transparency = 1                  ' invisible: only reserves the margin
        Else
            rect.Fill.ForeColor.RGB = BgColorValue(bg)
            rect.Fill.Transparency = BgAlpha() / 100
        End If
        rect.ZOrder MSO_SEND_TO_BACK
        If bg = BG_GRID Then AddGrid tmp, rect
        Set content = GroupAll(tmp)
    End If
    Set TryComposite = content
End Function

' Squared-paper background: light lines every GridStep points, a darker line every 5th.
Private Sub AddGrid(ByVal tmp As Object, ByVal rect As Object)
    Dim stepPt As Single, x As Single, y As Single, i As Long, ln As Object
    Dim l As Single, t As Single, r As Single, b As Single
    l = rect.Left: t = rect.Top
    r = l + rect.Width: b = t + rect.Height
    stepPt = GridStep()
    Do While (rect.Width / stepPt + rect.Height / stepPt) > 600
        stepPt = stepPt * 2
    Loop
    i = 1
    x = l + stepPt
    Do While x < r - 0.01
        Set ln = tmp.Shapes.AddLine(x, t, x, b)
        StyleGridLine ln, (i Mod 5 = 0)
        i = i + 1
        x = l + i * stepPt
    Loop
    i = 1
    y = t + stepPt
    Do While y < b - 0.01
        Set ln = tmp.Shapes.AddLine(l, y, r, y)
        StyleGridLine ln, (i Mod 5 = 0)
        i = i + 1
        y = t + i * stepPt
    Loop
    rect.ZOrder MSO_SEND_TO_BACK
End Sub

Private Sub StyleGridLine(ByVal ln As Object, ByVal major As Boolean)
    On Error Resume Next
    ln.Shadow.Visible = MSO_FALSE
    ln.Line.Visible = MSO_TRUE
    If major Then
        ln.Line.ForeColor.RGB = RGB(176, 184, 196)
        ln.Line.Weight = 0.75
    Else
        ln.Line.ForeColor.RGB = RGB(222, 226, 232)
        ln.Line.Weight = 0.5
    End If
    ln.ZOrder MSO_SEND_TO_BACK
End Sub

' Copies src and pastes it onto sld (as shapes or as an EMF picture). Returns the pasted range.
Private Function PasteFromClipboard(ByVal sld As Object, ByVal src As Object, ByVal asPicture As Boolean) As Object
    Dim k As Long, ok As Boolean, t As Single, r As Object
    For k = 1 To 8
        src.Copy
        DoEvents
        On Error Resume Next
        Err.Clear
        If asPicture Then
            Set r = sld.Shapes.PasteSpecial(PP_PASTE_EMF)
        Else
            Set r = sld.Shapes.Paste
        End If
        ok = (Err.Number = 0)
        On Error GoTo 0
        If ok Then
            Set PasteFromClipboard = r
            Exit Function
        End If
        t = Timer
        Do While Timer - t < 0.25 And Timer >= t
            DoEvents
        Loop
    Next k
    Err.Raise vbObjectError + 512, , Tx("err.clipboard")
End Function

Private Function GroupAll(ByVal sld As Object) As Object
    On Error Resume Next
    Set GroupAll = sld.Shapes.Range().Group
    On Error GoTo 0
End Function

Private Sub ClearSlide(ByVal sld As Object)
    Dim i As Long
    For i = sld.Shapes.Count To 1 Step -1
        sld.Shapes.Item(i).Delete
    Next i
End Sub


' =====================================================================
'  Platform layer: export, files, dialogs (Windows API / macOS AppleScript)
' =====================================================================

' Shape.Export straight to the final path (macOS: via the sandbox, then moved).
Private Sub ExportRaw(ByVal shp As Object, ByVal path As String, ByVal ppFmt As Long, _
                      Optional ByVal scaleW As Long = 0, Optional ByVal scaleH As Long = 0)
#If Mac Then
    Dim t As String
    t = TempFileX(ExtOf(path))
    ExportToTemp shp, t, ppFmt, scaleW, scaleH
    MacMove t, path
#Else
    DeleteIfExists path
    ExportToTemp shp, path, ppFmt, scaleW, scaleH
#End If
End Sub

' Shape.Export to a path PowerPoint is allowed to write (on macOS: inside the sandbox).
Private Sub ExportToTemp(ByVal shp As Object, ByVal path As String, ByVal ppFmt As Long, _
                         ByVal scaleW As Long, ByVal scaleH As Long)
    On Error GoTo EH
    If scaleW > 0 And scaleH > 0 Then
        shp.Export path, ppFmt, scaleW, scaleH, PP_RELATIVE_TO_SLIDE
    Else
        shp.Export path, ppFmt
    End If
    Exit Sub
EH:
#If Mac Then
    Err.Raise Err.Number, , Err.Description & vbCrLf & Tx("err.macVersion")
#Else
    Err.Raise Err.Number, , Err.Description
#End If
End Sub

' Moves a finished temporary file to its destination (overwriting).
Private Sub MoveTempTo(ByVal t As String, ByVal dst As String)
#If Mac Then
    MacMove t, dst
#Else
    DeleteIfExists dst
    Fso().MoveFile t, dst
#End If
End Sub

Private Function TempDirX() As String
    Dim d As String
#If Mac Then
    d = Environ("HOME") & "/VecStampTemp"
    If Len(Dir(d, vbDirectory)) = 0 Then MkDir d
#Else
    d = JoinPath(Fso().GetSpecialFolder(2).Path, "VecStamp")
    If Not Fso().FolderExists(d) Then Fso().CreateFolder d
#End If
    TempDirX = d
End Function

Private Function TempFileX(ByVal ext As String) As String
    Randomize
    TempFileX = JoinPath(TempDirX(), "vs_" & Format$(Now, "hhnnss") & "_" & CStr(Int(Rnd * 1000000)) & "." & ext)
End Function

Private Sub DeleteTempX(ByVal p As String)
    On Error Resume Next
#If Mac Then
    If Len(Dir(p)) > 0 Then Kill p
#Else
    If Fso().FileExists(p) Then Fso().DeleteFile p, True
#End If
End Sub

Private Sub MakeDirX(ByVal d As String)
#If Mac Then
    If Len(Dir(d, vbDirectory)) = 0 Then MkDir d
#Else
    If Not Fso().FolderExists(d) Then Fso().CreateFolder d
#End If
End Sub

Private Sub RemoveTempDirX(ByVal d As String)
    On Error Resume Next
    If InStr(d, "VecStamp") = 0 Then Exit Sub        ' only ever delete our own temp folders
#If Mac Then
    MacScript2 "removePath", d
#Else
    Fso().DeleteFolder d, True
#End If
End Sub

' Writes ASCII text (the VSDX XML parts are pure ASCII).
Private Sub WriteTextX(ByVal p As String, ByVal txt As String)
#If Mac Then
    Dim f As Integer
    f = FreeFile
    Open p For Output As #f
    Print #f, txt;
    Close #f
#Else
    With Fso().CreateTextFile(p, True, False)
        .Write txt
        .Close
    End With
#End If
End Sub

Private Sub CopyFileX(ByVal src As String, ByVal dst As String)
#If Mac Then
    FileCopy src, dst
#Else
    Fso().CopyFile src, dst, True
#End If
End Sub

' Raw file bytes packed into a String (use LenB / MidB / InStrB on the result).
Private Function ReadBytesX(ByVal p As String) As String
    Dim b() As Byte
#If Mac Then
    Dim f As Integer, n As Long
    f = FreeFile
    Open p For Binary Access Read As #f
    n = LOF(f)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #f, , b
    End If
    Close #f
    If n > 0 Then ReadBytesX = b
#Else
    With CreateObject("ADODB.Stream")
        .Type = 1
        .Open
        .LoadFromFile p
        If .Size > 0 Then
            b = .Read
            ReadBytesX = b
        End If
        .Close
    End With
#End If
End Function

Private Sub WriteBytesX(ByVal p As String, ByVal data As String)
    Dim b() As Byte
    b = data
#If Mac Then
    Dim f As Integer
    If Len(Dir(p)) > 0 Then Kill p
    f = FreeFile
    Open p For Binary Access Write As #f
    Put #f, , b
    Close #f
#Else
    With CreateObject("ADODB.Stream")
        .Type = 1
        .Open
        .Write b
        .SaveToFile p, 2
        .Close
    End With
#End If
End Sub

' ASCII text as a byte string, for comparisons with InStrB.
Private Function AsciiB(ByVal s As String) As String
    Dim i As Long, r As String
    For i = 1 To Len(s)
        r = r & ChrB(Asc(Mid$(s, i, 1)))
    Next i
    AsciiB = r
End Function

Private Function FileExistsX(ByVal p As String) As Boolean
#If Mac Then
    FileExistsX = (MacScript2("pathExists", p) = "1")
#Else
    FileExistsX = Fso().FileExists(p)
#End If
End Function

Private Function FolderExistsX(ByVal p As String) As Boolean
#If Mac Then
    FolderExistsX = (MacScript2("isFolder", p) = "1")
#Else
    FolderExistsX = Fso().FolderExists(p)
#End If
End Function

Private Function DesktopPathX() As String
    Dim d As String
#If Mac Then
    d = MacScript2("desktopFolder", "")
    If Len(d) > 1 And Right$(d, 1) = "/" Then d = Left$(d, Len(d) - 1)
#Else
    On Error Resume Next
    d = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    On Error GoTo 0
    If Len(d) = 0 Then d = Environ$("USERPROFILE") & "\Desktop"
#End If
    DesktopPathX = d
End Function

Private Function PickFolderX(ByVal initDir As String, ByVal promptText As String) As String
    Dim f As String
#If Mac Then
    f = MacScript2("chooseFolder", promptText & MAC_SEP & initDir)
#Else
    With Application.FileDialog(MSO_FOLDER_PICKER)
        .Title = promptText
        .AllowMultiSelect = False
        If Len(initDir) > 0 Then .InitialFileName = JoinPath(initDir, "")
        If .Show = -1 Then f = .SelectedItems(1)
    End With
#End If
    If Len(f) > 3 And Right$(f, 1) = PathSep() Then f = Left$(f, Len(f) - 1)
    PickFolderX = f
End Function

Private Sub OpenPathX(ByVal p As String)
#If Mac Then
    MacScript2 "openPath", p
#Else
    CreateObject("WScript.Shell").Run "explorer.exe """ & p & """", 1, False
#End If
End Sub

Private Sub RevealFileX(ByVal p As String)
#If Mac Then
    MacScript2 "revealFile", p
#Else
    CreateObject("WScript.Shell").Run "explorer.exe /select,""" & p & """", 1, False
#End If
End Sub

Private Function FileManagerName() As String
#If Mac Then
    FileManagerName = Tx("fm.mac")
#Else
    FileManagerName = Tx("fm.win")
#End If
End Function

' Save dialog. fmt is updated when the user picks another type / types another extension.
Private Function AskSavePath(ByVal initDir As String, ByVal baseName As String, ByRef fmt As Long) As String
    Dim p As String, k As Long
#If Mac Then
    p = MacScript2("chooseSaveFile", Tx("dlg.saveAs", UCase$(FmtExt(fmt))) & MAC_SEP & _
                   baseName & "." & FmtExt(fmt) & MAC_SEP & initDir)
    If Len(p) = 0 Then Exit Function
#Else
    Dim ofn As OPENFILENAMEW
    Dim fileBuf As String, filterStr As String, dlgTitle As String, defExt As String
    Dim r As Long, i As Long, folder As String, nm As String

    For i = 0 To FMT_LAST
        filterStr = filterStr & FilterLabel(i) & vbNullChar & "*." & FmtExt(i) & vbNullChar
    Next i
    filterStr = filterStr & vbNullChar
    fileBuf = baseName & String$(2048, vbNullChar)
    dlgTitle = AppTitle() & " - " & Tx("dlg.saveTitle")
    defExt = FmtExt(fmt)

    With ofn
        .lStructSize = LenB(ofn)
        .hwndOwner = GetActiveWindow()
        .lpstrFilter = StrPtr(filterStr)
        .nFilterIndex = fmt + 1
        .lpstrFile = StrPtr(fileBuf)
        .nMaxFile = Len(fileBuf)
        .lpstrInitialDir = StrPtr(initDir)
        .lpstrTitle = StrPtr(dlgTitle)
        .lpstrDefExt = StrPtr(defExt)
        .Flags = OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_PATHMUSTEXIST Or OFN_EXPLORER
    End With

    On Error GoTo Fallback
    r = GetSaveFileNameW(ofn)
    On Error GoTo 0
    If r = 0 Then Exit Function                     ' cancelled
    k = InStr(fileBuf, vbNullChar)
    If k > 0 Then p = Left$(fileBuf, k - 1) Else p = fileBuf
    k = FmtFromExt(LCase$(ExtOf(p)))
    If k < 0 Then
        fmt = ofn.nFilterIndex - 1
        If fmt < 0 Or fmt > FMT_LAST Then fmt = FMT_EMF
    End If
#End If
    k = FmtFromExt(LCase$(ExtOf(p)))
    If k >= 0 Then
        fmt = k                                     ' the extension decides the format
    Else
        p = p & "." & FmtExt(fmt)
    End If
    AskSavePath = p
    Exit Function

#If Mac Then
#Else
Fallback:
    ' Save dialog API unavailable: folder picker + file name prompt.
    folder = PickFolderX(initDir, Tx("dlg.pickFolder"))
    If Len(folder) = 0 Then Exit Function
    nm = InputBox(Tx("dlg.fileName"), AppTitle(), baseName)
    If Len(nm) = 0 Then Exit Function
    AskSavePath = JoinPath(folder, SafeName(nm) & "." & FmtExt(fmt))
#End If
End Function

Private Function GzipX(ByVal src As String, ByVal dst As String) As Boolean
    Dim ok As Boolean
#If Mac Then
    ok = (MacScript2("gzipFile", src & MAC_SEP & dst) = "ok")
#Else
    Dim cmd As String
    DeleteIfExists dst
    cmd = "$i=[IO.File]::OpenRead('" & PsQ(src) & "');" & _
          "$o=[IO.File]::Create('" & PsQ(dst) & "');" & _
          "$g=New-Object IO.Compression.GZipStream($o,[IO.Compression.CompressionMode]::Compress);" & _
          "$i.CopyTo($g);$g.Dispose();$o.Dispose();$i.Dispose()"
    ok = (RunPowerShell(cmd, 120) = 0)
    DeleteIfExists src
#End If
    If ok Then ok = FileExistsX(dst)
    GzipX = ok
End Function

' Zips a folder into an OPC package; [Content_Types].xml goes first, entry names use "/".
Private Sub ZipFolderX(ByVal dirPath As String, ByVal zipPath As String)
#If Mac Then
    If MacScript2("zipFolder", dirPath & MAC_SEP & zipPath) <> "ok" Then
        Err.Raise vbObjectError + 516, , Tx("err.zip")
    End If
#Else
    Dim cmd As String
    ProgressAt 0.94, Tx("step.zip")
    DeleteIfExists zipPath
    cmd = "Add-Type -AssemblyName System.IO.Compression;Add-Type -AssemblyName System.IO.Compression.FileSystem;" & _
          "$src='" & PsQ(dirPath) & "';$dst='" & PsQ(zipPath) & "';" & _
          "$z=[IO.Compression.ZipFile]::Open($dst,'Create');" & _
          "$fs=@(Get-ChildItem -LiteralPath $src -Recurse -File -Force|Sort-Object @{e={if($_.Name -eq '[Content_Types].xml'){0}else{1}}},FullName);" & _
          "foreach($f in $fs){$n=$f.FullName.Substring($src.Length).TrimStart('\').Replace('\','/');" & _
          "[void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($z,$f.FullName,$n)};$z.Dispose()"
    Select Case RunPowerShell(cmd, 120)
        Case 0
            If Not Fso().FileExists(zipPath) Then Err.Raise vbObjectError + 516, , Tx("err.zipPs")
        Case RUN_CANCELLED
            Err.Raise ERR_CANCEL, , "cancelled"
        Case Else
            Err.Raise vbObjectError + 516, , Tx("err.zipPs")
    End Select
#End If
End Sub

' EMF -> SVG with Inkscape or LibreOffice, when one of them is installed.
' Returns 0 on success, RUN_CANCELLED, or another non-zero value.
Private Function ConvertEmfToSvg(ByVal src As String, ByVal dst As String) As Long
    ConvertEmfToSvg = RUN_FAILED
#If Mac Then
    If MacScript2("inkscapeSvg", src & MAC_SEP & dst) = "ok" Then
        ConvertEmfToSvg = 0
    ElseIf MacScript2("libreofficeSvg", src & MAC_SEP & dst) = "ok" Then
        ConvertEmfToSvg = 0
    End If
#Else
    Dim c As Variant, exe As String, outDir As String, produced As String, st As Long
    For Each c In Array(Environ$("ProgramFiles") & "\Inkscape\bin\inkscape.exe", _
                        Environ$("ProgramW6432") & "\Inkscape\bin\inkscape.exe", _
                        Environ$("ProgramFiles(x86)") & "\Inkscape\bin\inkscape.exe", _
                        Environ$("LOCALAPPDATA") & "\Programs\Inkscape\bin\inkscape.exe")
        If Fso().FileExists(CStr(c)) Then exe = CStr(c): Exit For
    Next
    If Len(exe) > 0 Then
        ProgressAt 0.55, Tx("step.inkscape")
        st = RunHidden("""" & exe & """ """ & src & """ --export-type=svg ""--export-filename=" & dst & """", 120)
        If st = RUN_CANCELLED Then ConvertEmfToSvg = st: Exit Function
        If Fso().FileExists(dst) Then ConvertEmfToSvg = 0: Exit Function
    End If

    exe = ""
    For Each c In Array(Environ$("ProgramFiles") & "\LibreOffice\program\soffice.exe", _
                        Environ$("ProgramW6432") & "\LibreOffice\program\soffice.exe", _
                        Environ$("ProgramFiles(x86)") & "\LibreOffice\program\soffice.exe")
        If Fso().FileExists(CStr(c)) Then exe = CStr(c): Exit For
    Next
    If Len(exe) > 0 Then
        outDir = JoinPath(TempDirX(), "lo_" & Format$(Now, "hhnnss"))
        MakeDirX outDir
        ProgressAt 0.6, Tx("step.libreoffice")
        st = RunHidden("""" & exe & """ -env:UserInstallation=" & FileUrlX(JoinPath(TempDirX(), "lo_profile")) & _
                " --headless --norestore --convert-to svg --outdir """ & outDir & """ """ & src & """", 180)
        produced = JoinPath(outDir, BaseNameOf(src) & ".svg")
        If Fso().FileExists(produced) Then
            DeleteIfExists dst
            Fso().MoveFile produced, dst
            ConvertEmfToSvg = 0
        ElseIf st = RUN_CANCELLED Then
            ConvertEmfToSvg = st
        End If
        RemoveTempDirX outDir
    End If
#End If
End Function

#If Mac Then
' Calls a handler of the AppleScript helper (outside the PowerPoint sandbox).
Private Function MacScript2(ByVal handler As String, ByVal param As String) As String
    On Error GoTo EH
    MacScript2 = AppleScriptTask(MAC_SCRIPT, handler, param)
    Exit Function
EH:
    Err.Raise vbObjectError + 520, , Tx("err.macHelper", MAC_SCRIPT)
End Function

Private Sub MacMove(ByVal src As String, ByVal dst As String)
    If MacScript2("moveFile", src & MAC_SEP & dst) <> "ok" Then
        Err.Raise vbObjectError + 521, , Tx("err.macWrite", dst)
    End If
End Sub
#Else
Private Function Fso() As Object
    Static f As Object
    If f Is Nothing Then Set f = CreateObject("Scripting.FileSystemObject")
    Set Fso = f
End Function

Private Sub DeleteIfExists(ByVal p As String)
    If Fso().FileExists(p) Then Fso().DeleteFile p, True
End Sub

' Runs a program without a window and without freezing PowerPoint: while it works, the
' progress window keeps updating (from progressFile, when given), Esc ends the program and
' so does the time limit. Returns the program's exit code, RUN_FAILED (could not start),
' RUN_CANCELLED or RUN_TIMEOUT.
Private Function RunHidden(ByVal cmdLine As String, ByVal timeoutSec As Double, _
                           Optional ByVal progressFile As String = "", Optional ByVal frac0 As Double = 0, _
                           Optional ByVal frac1 As Double = 0) As Long
    Dim si As STARTUPINFOW, pi As PROCESS_INFORMATION, buf As String, code As Long, t0 As Double, el As Double
    Dim lastSec As Long
    buf = cmdLine & vbNullChar
    si.cb = LenB(si)
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = SW_HIDE
    If CreateProcessW(0, StrPtr(buf), 0, 0, 0, CREATE_NO_WINDOW, 0, 0, si, pi) = 0 Then
        RunHidden = RUN_FAILED
        Exit Function
    End If
    t0 = Timer
    lastSec = -1
    Do
        If WaitForSingleObject(pi.hProcess, 120) <> WAIT_TIMEOUT Then Exit Do
        el = Timer - t0
        If el < 0 Then el = el + 86400#
        If Len(progressFile) > 0 Then ProgressFromFile progressFile, frac0, frac1, el
        If Int(el) <> lastSec Then
            lastSec = Int(el)
            ProgressElapsed lastSec
        End If
        DoEvents
        If ProgressCancelled() Then
            TerminateProcess pi.hProcess, 1
            WaitForSingleObject pi.hProcess, 3000
            code = RUN_CANCELLED
            GoTo Done
        End If
        If el > timeoutSec Then
            TerminateProcess pi.hProcess, 1
            WaitForSingleObject pi.hProcess, 3000
            code = RUN_TIMEOUT
            GoTo Done
        End If
    Loop
    If GetExitCodeProcess(pi.hProcess, code) = 0 Then code = RUN_FAILED
Done:
    CloseHandle pi.hThread
    CloseHandle pi.hProcess
    RunHidden = code
End Function

' Runs a PowerShell snippet (no window). The script is passed as plain text with -Command - no
' encoded command and no execution-policy switch, so security software sees ordinary use.
Private Function RunPowerShell(ByVal script As String, ByVal timeoutSec As Double, _
                               Optional ByVal progressFile As String = "", Optional ByVal frac0 As Double = 0, _
                               Optional ByVal frac1 As Double = 0) As Long
    Dim exe As String
    exe = Environ$("SystemRoot") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
    If Not Fso().FileExists(exe) Then exe = "powershell.exe"
    RunPowerShell = RunHidden("""" & exe & """ -NoLogo -NoProfile -NonInteractive -Command """ & _
                              "$ErrorActionPreference='Stop';" & Replace(script, """", "'") & """", _
                              timeoutSec, progressFile, frac0, frac1)
End Function

' Shows the step the XPS -> SVG converter reports in its progress file ("percent|step|a|b").
Private Sub ProgressFromFile(ByVal path As String, ByVal frac0 As Double, ByVal frac1 As Double, ByVal el As Double)
    Dim txt As String, f() As String, pct As Double, d As String
    On Error Resume Next
    If Not Fso().FileExists(path) Then Exit Sub
    With Fso().OpenTextFile(path, 1)
        txt = .ReadAll
        .Close
    End With
    If Err.Number <> 0 Or Len(txt) = 0 Or txt = gPgFile Then Exit Sub
    On Error GoTo 0
    gPgFile = txt
    f = Split(txt, "|")
    If UBound(f) < 3 Then Exit Sub
    pct = Val(f(0)) / 100#
    Select Case f(1)
        Case "open", "write", "done": d = Tx("xps." & f(1))
        Case "page": d = Tx("xps.page", f(3))
        Case "elem": d = Tx("xps.elem", f(2), f(3))
        Case "font": d = Tx("xps.font", f(2))
        Case Else: d = f(1)
    End Select
    ProgressAt frac0 + (frac1 - frac0) * pct, Tx("step.xps", d)
End Sub

Private Function PsQ(ByVal s As String) As String
    ' Escape for a PowerShell single-quoted string (PowerShell also treats curly quotes as quotes).
    s = Replace(s, "'", "''")
    s = Replace(s, ChrW(&H2018), ChrW(&H2018) & ChrW(&H2018))
    s = Replace(s, ChrW(&H2019), ChrW(&H2019) & ChrW(&H2019))
    PsQ = s
End Function

' file:/// URL with UTF-8 percent-encoding (for LibreOffice's -env:UserInstallation).
Private Function FileUrlX(ByVal p As String) As String
    Dim b() As Byte, i As Long, c As Long, s As String
    With CreateObject("ADODB.Stream")
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText Replace(p, "\", "/")
        .Position = 0
        .Type = 1
        .Position = 3                               ' skip the UTF-8 byte order mark
        b = .Read
        .Close
    End With
    For i = LBound(b) To UBound(b)
        c = b(i)
        If (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or _
           c = 45 Or c = 46 Or c = 95 Or c = 126 Or c = 47 Or c = 58 Then
            s = s & Chr$(c)
        Else
            s = s & "%" & Right$("0" & Hex$(c), 2)
        End If
    Next i
    FileUrlX = "file:///" & s
End Function
#End If


' =====================================================================
'  Progress window (Windows: a small native window; macOS: a dialog from the helper)
'  A run has items (shapes, files or slides); ProgressAt sets the position inside the current
'  item and the step being done, and the window shows the overall percentage.
' =====================================================================
Private Sub ProgressBegin(ByVal title As String, ByVal items As Long)
    ProgressEnd
    gPgItems = items
    If gPgItems < 1 Then gPgItems = 1
    gPgItem = 1
    gPgFrac = 0
    gPgActive = True
    gPgCancel = False
    gPgTitle = title
    gPgLabel = ""
    gPgDetail = Tx("prog.preparing")
    gPgFile = ""
#If Mac Then
    gPgMacPid = ""
    gPgMacPct = -100
    gPgMacTime = 0
    ProgressShow
#Else
    GetAsyncKeyState VK_ESCAPE                     ' forget an Esc pressed before the run started
    ProgressCreate title
#End If
End Sub

' Starts item i of the run, with a short label such as "2 / 5: Rectangle 3".
Private Sub ProgressItem(ByVal i As Long, ByVal label As String)
    If Not gPgActive Then Exit Sub
    gPgItem = i
    If gPgItem > gPgItems Then gPgItems = gPgItem
    gPgFrac = 0
    gPgLabel = label
    ProgressShow
End Sub

' Position inside the current item (0..1, never backwards) and the step being done.
Private Sub ProgressAt(ByVal frac As Double, ByVal detail As String)
    If Not gPgActive Then Exit Sub
    If frac < gPgFrac Then frac = gPgFrac
    If frac > 1 Then frac = 1
    gPgFrac = frac
    If Len(detail) > 0 Then gPgDetail = detail
    ProgressShow
End Sub

' While an external program runs: show how long it has been working.
Private Sub ProgressElapsed(ByVal sec As Long)
    If Not gPgActive Or sec < 2 Then Exit Sub
#If Mac Then
#Else
    ProgressShow sec
#End If
End Sub

Private Function ProgressPercent() As Long
    Dim p As Long
    p = Int(((gPgItem - 1) + gPgFrac) / gPgItems * 100# + 0.5)
    If p < 0 Then p = 0
    If p > 100 Then p = 100
    ProgressPercent = p
End Function

Private Sub ProgressEnd()
    gPgActive = False
#If Mac Then
    If Len(gPgMacPid) > 0 Then
        On Error Resume Next
        AppleScriptTask MAC_SCRIPT, "progressHide", gPgMacPid
        gPgMacPid = ""
    End If
#Else
    On Error Resume Next
    If gPgWnd <> 0 Then DestroyWindow gPgWnd
    gPgWnd = 0
    If gPgFont <> 0 Then DeleteObject gPgFont
    If gPgFontB <> 0 Then DeleteObject gPgFontB
    gPgFont = 0
    gPgFontB = 0
#End If
End Sub

' True once the user pressed Esc (Windows) or closed the progress dialog with its button
' (macOS) during the run; it stays True until the run ends.
Private Function ProgressCancelled() As Boolean
#If Mac Then
    If gPgActive And Not gPgCancel And Len(gPgMacPid) > 0 Then
        If Abs(Timer - gPgMacCheck) >= 1.5 Then
            gPgMacCheck = Timer
            On Error Resume Next
            If AppleScriptTask(MAC_SCRIPT, "progressAlive", gPgMacPid) = "0" Then gPgCancel = True
        End If
    End If
#Else
    If gPgActive And Not gPgCancel Then
        If (GetAsyncKeyState(VK_ESCAPE) And &H8001) <> 0 Then gPgCancel = True
    End If
#End If
    ProgressCancelled = gPgCancel
End Function

' Stops the run (error ERR_CANCEL, handled by the export flows) when Esc was pressed.
Private Sub CheckCancel()
    If ProgressCancelled() Then Err.Raise ERR_CANCEL, , "cancelled"
End Sub

' MsgBox that hides the progress window while it is shown.
Private Function AskUser(ByVal prompt As String, ByVal buttons As VbMsgBoxStyle) As VbMsgBoxResult
#If Mac Then
    AskUser = MsgBox(prompt, buttons, AppTitle())
#Else
    If gPgWnd <> 0 Then ShowWindow gPgWnd, SW_HIDE
    AskUser = MsgBox(prompt, buttons, AppTitle())
    If gPgWnd <> 0 And gPgActive Then ShowWindow gPgWnd, SW_SHOWNOACTIVATE
    GetAsyncKeyState VK_ESCAPE                     ' an Esc that closed the message box does not cancel
#End If
End Function

#If Mac Then
' macOS: the helper shows a small dialog; it is replaced when the text changes noticeably
' (at most every 1.5 seconds, the dialog cannot be updated in place).
Private Sub ProgressShow(Optional ByVal sec As Long = -1)
    Dim pct As Long, txt As String, now As Double
    If Not gPgActive Then Exit Sub
    pct = ProgressPercent()
    txt = gPgTitle
    If Len(gPgLabel) > 0 Then txt = txt & " - " & gPgLabel
    txt = txt & vbLf & gPgDetail & vbLf & pct & "%"
    now = Timer
    If Len(gPgMacPid) > 0 Then
        If Abs(pct - gPgMacPct) < 10 And Abs(now - gPgMacTime) < 1.5 Then Exit Sub
        If txt = gPgMacText Then Exit Sub
    End If
    On Error Resume Next
    If Len(gPgMacPid) > 0 Then
        gPgMacPid = AppleScriptTask(MAC_SCRIPT, "progressUpdate", gPgMacPid & MAC_SEP & txt & MAC_SEP & Tx("btn.cancel"))
    Else
        gPgMacPid = AppleScriptTask(MAC_SCRIPT, "progressShow", txt & MAC_SEP & Tx("btn.cancel"))
    End If
    gPgMacPct = pct
    gPgMacTime = now
    gPgMacText = txt
    gPgMacCheck = now
End Sub
#Else
Private Sub ProgressCreate(ByVal title As String)
    Dim r As RECT, dpi As Long, k As Double, w As Long, h As Long, x As Long, y As Long
    Dim cls As String, cap As String, face As String, icc As INITCOMMONCONTROLSEX_T
    Dim owner As LongPtr, hdc As LongPtr
    On Error GoTo Done
    icc.dwSize = LenB(icc)
    icc.dwICC = ICC_PROGRESS_CLASS
    InitCommonControlsEx icc
    owner = GetActiveWindow()
    hdc = GetDC(0)
    dpi = GetDeviceCaps(hdc, LOGPIXELSY)
    ReleaseDC 0, hdc
    If dpi < 72 Then dpi = 96
    k = dpi / 96#
    w = CLng(480 * k)
    h = CLng(158 * k)
    If owner <> 0 And GetWindowRect(owner, r) <> 0 Then
        x = r.Left + (r.Right - r.Left - w) \ 2
        y = r.Top + (r.Bottom - r.Top - h) \ 2
    Else
        x = (GetSystemMetrics(0) - w) \ 2
        y = (GetSystemMetrics(1) - h) \ 2
    End If
    cls = "#32770"
    cap = AppTitle()
    gPgWnd = CreateWindowExW(WS_EX_DLGMODALFRAME Or WS_EX_TOOLWINDOW, StrPtr(cls), StrPtr(cap), WS_POPUP Or WS_CAPTION, _
                             x, y, w, h, owner, 0, 0, 0)
    If gPgWnd = 0 Then Exit Sub
    face = UiFontName()
    gPgFontB = CreateFontW(-CLng(10.5 * dpi / 72), 0, 0, 0, 700, 0, 0, 0, 1, 0, 0, 5, 0, StrPtr(face))
    gPgFont = CreateFontW(-CLng(9 * dpi / 72), 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 5, 0, StrPtr(face))
    cls = "Static"
    gPgMsg = CreateWindowExW(0, StrPtr(cls), StrPtr(title), WS_CHILD Or WS_VISIBLE Or SS_NOPREFIX, _
                             CLng(20 * k), CLng(14 * k), CLng(360 * k), CLng(24 * k), gPgWnd, 0, 0, 0)
    cap = "0%"
    gPgPct = CreateWindowExW(0, StrPtr(cls), StrPtr(cap), WS_CHILD Or WS_VISIBLE Or SS_NOPREFIX Or SS_RIGHT, _
                             CLng(380 * k), CLng(14 * k), CLng(72 * k), CLng(24 * k), gPgWnd, 0, 0, 0)
    cls = "msctls_progress32"
    gPgBar = CreateWindowExW(0, StrPtr(cls), 0, WS_CHILD Or WS_VISIBLE, _
                             CLng(20 * k), CLng(44 * k), CLng(432 * k), CLng(16 * k), gPgWnd, 0, 0, 0)
    cls = "Static"
    cap = gPgDetail
    gPgSub = CreateWindowExW(0, StrPtr(cls), StrPtr(cap), WS_CHILD Or WS_VISIBLE Or SS_NOPREFIX, _
                             CLng(20 * k), CLng(68 * k), CLng(432 * k), CLng(42 * k), gPgWnd, 0, 0, 0)
    SendMessageW gPgMsg, WM_SETFONT, gPgFontB, 1
    SendMessageW gPgPct, WM_SETFONT, gPgFontB, 1
    SendMessageW gPgSub, WM_SETFONT, gPgFont, 1
    If gPgBar <> 0 Then SendMessageW gPgBar, PBM_SETRANGE32, 0, 100
    ShowWindow gPgWnd, SW_SHOWNOACTIVATE
    UpdateWindow gPgWnd
    DoEvents
Done:
End Sub

Private Sub ProgressShow(Optional ByVal sec As Long = -1)
    Dim head As String, detail As String, pct As Long
    On Error Resume Next
    If gPgWnd = 0 Then Exit Sub
    If IsWindow(gPgWnd) = 0 Then gPgWnd = 0: Exit Sub
    pct = ProgressPercent()
    head = gPgTitle
    If Len(gPgLabel) > 0 Then head = head & " - " & gPgLabel
    detail = gPgDetail
    If sec >= 2 Then detail = detail & Tx("prog.elapsed", sec)
    detail = detail & vbCrLf & Tx("prog.escHint")
    SetWindowTextW gPgMsg, StrPtr(head)
    SetWindowTextW gPgPct, StrPtr(CStr(pct) & "%")
    SetWindowTextW gPgSub, StrPtr(detail)
    If gPgBar <> 0 Then SendMessageW gPgBar, PBM_SETPOS, pct, 0
    UpdateWindow gPgWnd
    DoEvents
End Sub
#End If


' Font for native windows in the current interface language.
Private Function UiFontName() As String
    Select Case CurLang()
        Case "zh-TW": UiFontName = "Microsoft JhengHei UI"
        Case "ja": UiFontName = "Yu Gothic UI"
        Case "en": UiFontName = "Segoe UI"
        Case Else: UiFontName = "Microsoft YaHei UI"
    End Select
End Function

' A name shortened for one-line labels.
Private Function ShortName(ByVal nm As String) As String
    If Len(nm) > 28 Then nm = Left$(nm, 26) & "…"
    ShortName = nm
End Function


' =====================================================================
'  Feedback, About and donations
' =====================================================================
Private Sub DoFeedback()
    Dim ans As VbMsgBoxResult
    ans = MsgBox(Tx("fb.ask", VS_EMAIL), vbQuestion + vbYesNoCancel, AppTitle() & " - " & Tx("fb.title"))
    Select Case ans
        Case vbYes: FeedbackMail
        Case vbNo: FeedbackIssue
    End Select
End Sub

Private Sub FeedbackMail()
    Dim copied As Boolean, url As String
    copied = SetClipboardText(VS_EMAIL)
    If MsgBox(Tx(IIf(copied, "fb.mailCopied", "fb.mailPlain"), VS_EMAIL), _
              vbQuestion + vbYesNo + vbDefaultButton2, AppTitle() & " - " & Tx("fb.mailTitle")) <> vbYes Then Exit Sub
    url = "mailto:" & VS_EMAIL & "?subject=" & UrlEncode(Tx("fb.subject", VS_VERSION)) & _
          "&body=" & UrlEncode(vbCrLf & vbCrLf & "----" & vbCrLf & EnvironmentInfo())
    If Not OpenUrlX(url) Then
        MsgBox Tx("fb.noMailApp", VS_EMAIL), vbInformation, AppTitle()
    End If
End Sub

Private Sub FeedbackIssue()
    Dim url As String, body As String
    body = Replace(Tx("fb.issueBody", EnvironmentInfo()), vbCrLf, vbLf)
    url = VS_REPO & "/issues/new?title=" & UrlEncode(Tx("fb.issueTitle")) & "&body=" & UrlEncode(body)
    If Not OpenUrlX(url) Then
        MsgBox Tx("fb.noBrowser", VS_REPO & "/issues"), vbInformation, AppTitle()
    End If
End Sub

Private Function EnvironmentInfo() As String
    Dim s As String
    s = "VecStamp: " & VS_VERSION
    On Error Resume Next
    s = s & vbCrLf & "PowerPoint: " & Application.Version & " (build " & Application.Build & ")"
    s = s & vbCrLf & "OS: " & Application.OperatingSystem
    EnvironmentInfo = s
End Function

Private Sub DoAbout()
    MsgBox Tx("about.text", VS_VERSION, VS_AUTHOR, VS_EMAIL, VS_REPO, VS_YEAR) & vbCrLf & vbCrLf & _
           Tx("about.lang", LangName(CurLang())), vbInformation, Tx("about.title", AppTitle())
End Sub

Private Sub DoDonateThanks(ByVal which As String)
    MsgBox Tx("don.thanks", which), vbInformation, AppTitle() & " - " & Tx("don.title")
End Sub

Private Function OpenUrlX(ByVal url As String) As Boolean
    On Error GoTo Failed
#If Mac Then
    MacScript2 "openPath", url
#Else
    CreateObject("Shell.Application").ShellExecute url
#End If
    OpenUrlX = True
    Exit Function
Failed:
    OpenUrlX = False
End Function

Private Function SetClipboardText(ByVal s As String) As Boolean
    On Error GoTo Failed
#If Mac Then
    SetClipboardText = (MacScript2("setClipboard", s) = "ok")
#Else
    With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")    ' MSForms.DataObject
        .SetText s
        .PutInClipboard
    End With
    SetClipboardText = True
#End If
    Exit Function
Failed:
    SetClipboardText = False
End Function

' Percent-encoding of the UTF-8 bytes (for mailto: and https: query strings).
Private Function UrlEncode(ByVal s As String) As String
    Dim b() As Byte, i As Long, c As Long, r As String, raw As String
    raw = Utf8Bytes(s)
    If LenB(raw) = 0 Then Exit Function
    b = raw
    For i = LBound(b) To UBound(b)
        c = b(i)
        If (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or c = 45 Or c = 46 Or c = 95 Or c = 126 Then
            r = r & Chr$(c)
        Else
            r = r & "%" & Right$("0" & Hex$(c), 2)
        End If
    Next i
    UrlEncode = r
End Function


' =====================================================================
'  Background colour gallery: theme colours (with tints and shades), standard
'  colours, recent colours, "More Colors" and the eyedropper - like PowerPoint's own
' =====================================================================
Private Function ThemeBase(ByVal col As Long) As Long
    ' column order as in PowerPoint: Background 1, Text 1, Background 2, Text 2, Accent 1-6
    Dim idx As Variant, defs As Variant
    idx = Array(2, 1, 4, 3, 5, 6, 7, 8, 9, 10)
    defs = Array(&HFFFFFF, &H0&, &HE6E6E7, &H6A5444, &HC47244, &H317DED, &HA5A5A5, &HC0FF&, &HD59B5B, &H47AD70)
    ThemeBase = defs(col)
    On Error Resume Next
    ThemeBase = ActivePresentation.SlideMaster.Theme.ThemeColorScheme.Colors(idx(col)).RGB
End Function

' Office's tint / shade rows: lighter 80/60/40 %, darker 25/50 % (other steps for very light / dark colours).
Private Function ThemeTint(ByVal base As Long, ByVal row As Long, ByRef label As String) As Long
    Dim h As Double, s As Double, l As Double, p As Variant, f As Double, lighter As Boolean
    RgbToHsl base, h, s, l
    If l >= 0.999 Then
        p = Array(-0.05, -0.15, -0.25, -0.35, -0.5)
    ElseIf l <= 0.001 Then
        p = Array(0.5, 0.35, 0.25, 0.15, 0.05)
    ElseIf l > 0.8 Then
        p = Array(-0.1, -0.25, -0.5, -0.75, -0.9)
    ElseIf l < 0.2 Then
        p = Array(0.9, 0.75, 0.5, 0.25, 0.1)
    Else
        p = Array(0.8, 0.6, 0.4, -0.25, -0.5)
    End If
    f = p(row - 1)
    lighter = (f > 0)
    If lighter Then
        l = l * (1 - f) + f
        label = Tx("col.lighter", CLng(f * 100))
    Else
        l = l * (1 + f)
        label = Tx("col.darker", CLng(-f * 100))
    End If
    ThemeTint = HslToRgb(h, s, l)
End Function

Private Function GalColor(ByVal index As Long, Optional ByRef label As String) As Long
    Dim names As Variant, std As Variant, stdNames As Variant, rc As Variant, col As Long, row As Long, t As String
    names = Split(Tx("col.themeNames"), "|")
    std = Array(&HC0&, &HFF&, &HC0FF&, &HFFFF&, &H50D092, &H50B000, &HF0B000, &HC07000, &H602000, &HA03070)
    stdNames = Split(Tx("col.stdNames"), "|")
    If index < 60 Then
        col = index Mod 10
        row = index \ 10
        If row = 0 Then
            GalColor = ThemeBase(col)
            label = names(col)
        Else
            GalColor = ThemeTint(ThemeBase(col), row, t)
            label = Tx("col.tint", names(col), t)
        End If
    ElseIf index < 70 Then
        GalColor = std(index - 60)
        label = Tx("col.stdItem", stdNames(index - 60))
    Else
        rc = RecentColors()
        If index - 70 <= UBound(rc) Then
            GalColor = rc(index - 70)
            label = Tx("col.recentItem", ColorText(rc(index - 70)))
        End If
    End If
    If Len(label) > 0 And index < 70 Then label = Tx("col.withValue", label, ColorText(GalColor))
End Function

Private Function ColorText(ByVal c As Long) As String
    ColorText = "RGB " & (c And &HFF&) & "," & ((c \ &H100&) And &HFF&) & "," & ((c \ &H10000) And &HFF&)
End Function

Private Function RecentColors() As Variant
    Dim parts() As String, i As Long, n As Long, r() As Long, s As String
    s = GetSet("RecentColors", "")
    If Len(s) = 0 Then RecentColors = Array(): Exit Function
    parts = Split(s, ";")
    ReDim r(0 To UBound(parts))
    For i = 0 To UBound(parts)
        If Len(parts(i)) = 6 Then
            r(n) = RGB(CLng("&H" & Mid$(parts(i), 1, 2)), CLng("&H" & Mid$(parts(i), 3, 2)), CLng("&H" & Mid$(parts(i), 5, 2)))
            n = n + 1
        End If
    Next i
    If n = 0 Then RecentColors = Array(): Exit Function
    ReDim Preserve r(0 To n - 1)
    RecentColors = r
End Function

Private Sub AddRecentColor(ByVal c As Long)
    Dim rc As Variant, i As Long, s As String, n As Long, hx As String
    hx = Mid$(HexRGB(c), 2)
    s = hx
    n = 1
    rc = RecentColors()
    For i = 0 To UBound(rc)
        If rc(i) <> c And n < 10 Then
            s = s & ";" & Mid$(HexRGB(CLng(rc(i))), 2)
            n = n + 1
        End If
    Next i
    PutSet "RecentColors", s
End Sub

' Uses c as the custom background colour (and switches the background to "custom colour").
Private Sub SetBgColor(ByVal c As Long)
    PutSet "BgColor", (c And &HFF&) & "," & ((c \ &H100&) And &HFF&) & "," & ((c \ &H10000) And &HFF&)
    PutSet "BgType", CStr(BG_CUSTOM)
    AddRecentColor c
    RefreshRibbon
End Sub

' "More Colors": PowerPoint's own colour dialog (Windows: run on a temporary off-slide shape whose
' fill colour is read afterwards; macOS: the macOS colour panel, which is what PowerPoint uses there).
' When it cannot be opened, VecStamp's own colour panel takes over (then the system colour dialog,
' then a colour list as the last resort).
Private Sub DoMoreColors()
    Dim c As Long, trans As Double
    On Error GoTo EH
    c = BgColorValue(BG_CUSTOM)
#If Mac Then
    If MacColorPicker(c) Then SetBgColor c
#Else
    trans = BgAlpha() / 100#
    Select Case PickWithPowerPoint("ObjectFillMoreColorsDialog", False, c, trans)
        Case 1
            SetBgColor c
            PutSet "BgAlpha", CStr(CLng(trans * 100))
            RefreshRibbon
        Case -1
            Select Case ColorPanel(c)
                Case 1: SetBgColor c
                Case -1
                    If SysColorDialog(c, GetActiveWindow()) Then SetBgColor c
            End Select
    End Select
#End If
    Exit Sub
EH:
    If ColorListPanel(c) Then SetBgColor c
End Sub

Private Sub DoEyedropper()
    Dim c As Long, trans As Double
    c = BgColorValue(BG_CUSTOM)
    trans = 0
    Select Case PickWithPowerPoint("EyedropperFill", True, c, trans)
        Case 1
            SetBgColor c
        Case -1
            MsgBox Tx("msg.eyedropperView"), vbInformation, AppTitle()
    End Select
End Sub

' Returns 1 when a colour was picked, 0 when the user cancelled, -1 when the tool is not available.
Private Function PickWithPowerPoint(ByVal idMso As String, ByVal waitForPick As Boolean, _
                                    ByRef c As Long, ByRef trans As Double) As Long
    Dim win As Object, sld As Object, tmp As Object, prev As Object, pres As Object
    Dim c0 As Long, t0 As Double, started As Double, wasSaved As Long, changed As Boolean
    On Error Resume Next
    Set win = Application.ActiveWindow
    Set sld = win.View.Slide
    Set pres = win.Presentation
    PickWithPowerPoint = -1
    If sld Is Nothing Or pres Is Nothing Then Exit Function
    If win.Selection.Type = PP_SEL_SHAPES Then Set prev = win.Selection.ShapeRange
    Err.Clear
    On Error GoTo Failed
    wasSaved = pres.Saved
    Set tmp = sld.Shapes.AddShape(MSO_RECTANGLE, -120, 10, 60, 60)
    tmp.Name = "VecStamp colour"
    tmp.Line.Visible = MSO_FALSE
    tmp.Fill.Visible = MSO_TRUE
    tmp.Fill.Solid
    tmp.Fill.ForeColor.RGB = c
    tmp.Fill.Transparency = trans
    tmp.Select
    c0 = tmp.Fill.ForeColor.RGB
    t0 = tmp.Fill.Transparency
    If waitForPick Then
        ProgressBegin Tx("prog.eyedropper"), 1
        ProgressItem 1, ""
        ProgressAt 0, Tx("step.eyedropper")
    End If
    Application.CommandBars.ExecuteMso idMso
    If waitForPick Then
        started = Timer
        Do
            DoEvents
            If tmp.Fill.ForeColor.RGB <> c0 Then Exit Do
            If ProgressCancelled() Then Exit Do
            If Timer - started > 60 Or Timer < started Then Exit Do
            ProgressAt CDbl(Timer - started) / 60#, ""
        Loop
        ProgressEnd
    End If
    changed = (tmp.Fill.ForeColor.RGB <> c0) Or (Abs(tmp.Fill.Transparency - t0) > 0.001)
    If changed Then
        c = tmp.Fill.ForeColor.RGB
        trans = tmp.Fill.Transparency
    End If
    tmp.Delete
    pres.Saved = wasSaved
    On Error Resume Next
    If Not prev Is Nothing Then prev.Select
    PickWithPowerPoint = IIf(changed, 1, 0)
    Exit Function
Failed:
    Resume FailCleanup
FailCleanup:
    On Error Resume Next
    ProgressEnd
    If Not tmp Is Nothing Then
        tmp.Delete
        pres.Saved = wasSaved
    End If
    If Not prev Is Nothing Then prev.Select
    PickWithPowerPoint = -1
End Function

Private Sub RgbToHsl(ByVal c As Long, ByRef h As Double, ByRef s As Double, ByRef l As Double)
    Dim r As Double, g As Double, b As Double, mx As Double, mn As Double, d As Double
    r = (c And &HFF&) / 255#
    g = ((c \ &H100&) And &HFF&) / 255#
    b = ((c \ &H10000) And &HFF&) / 255#
    mx = r: If g > mx Then mx = g
    If b > mx Then mx = b
    mn = r: If g < mn Then mn = g
    If b < mn Then mn = b
    l = (mx + mn) / 2
    d = mx - mn
    If d < 0.000001 Then
        h = 0: s = 0
        Exit Sub
    End If
    If l > 0.5 Then s = d / (2 - mx - mn) Else s = d / (mx + mn)
    If mx = r Then
        h = (g - b) / d + IIf(g < b, 6, 0)
    ElseIf mx = g Then
        h = (b - r) / d + 2
    Else
        h = (r - g) / d + 4
    End If
    h = h / 6
End Sub

Private Function HslToRgb(ByVal h As Double, ByVal s As Double, ByVal l As Double) As Long
    Dim q As Double, p As Double
    If l < 0 Then l = 0
    If l > 1 Then l = 1
    If s < 0.000001 Then
        HslToRgb = RGB(CLng(l * 255), CLng(l * 255), CLng(l * 255))
        Exit Function
    End If
    If l < 0.5 Then q = l * (1 + s) Else q = l + s - l * s
    p = 2 * l - q
    HslToRgb = RGB(CLng(HueChannel(p, q, h + 1 / 3) * 255), CLng(HueChannel(p, q, h) * 255), CLng(HueChannel(p, q, h - 1 / 3) * 255))
End Function

Private Function HueChannel(ByVal p As Double, ByVal q As Double, ByVal t As Double) As Double
    If t < 0 Then t = t + 1
    If t > 1 Then t = t - 1
    If t < 1 / 6 Then
        HueChannel = p + (q - p) * 6 * t
    ElseIf t < 0.5 Then
        HueChannel = q
    ElseIf t < 2 / 3 Then
        HueChannel = p + (q - p) * (2 / 3 - t) * 6
    Else
        HueChannel = p
    End If
End Function

' Swatches for the ribbon (Windows: a small BMP file loaded as a picture).
Private Function SwatchPicture(ByVal c As Long, ByVal size As Long) As Object
    Set SwatchPicture = PatternPicture(0, c, size)
End Function

' Preview square of a background: transparent = checkerboard, grid = graph paper, else the colour.
Private Function BgPicture(ByVal bg As Long, ByVal size As Long) As Object
    Select Case bg
        Case BG_NONE: Set BgPicture = PatternPicture(1, RGB(255, 255, 255), size)
        Case BG_GRID: Set BgPicture = PatternPicture(2, RGB(255, 255, 255), size)
        Case Else: Set BgPicture = PatternPicture(0, BgColorValue(bg), size)
    End Select
End Function

' kind: 0 = solid colour c, 1 = checkerboard, 2 = grid on c; grey frame.
Private Function PatternPicture(ByVal kind As Long, ByVal c As Long, ByVal size As Long) As Object
#If Mac Then
#Else
    Dim p As String, b() As Byte, i As Long, x As Long, y As Long, rowBytes As Long, n As Long
    Dim r As Long, g As Long, bl As Long, cell As Long, raw As String
    On Error GoTo Done
    p = JoinPath(TempDirX(), "swatch" & kind & "_" & Mid$(HexRGB(c), 2) & "_" & size & ".bmp")
    If Not Fso().FileExists(p) Then
        rowBytes = ((size * 3 + 3) \ 4) * 4
        n = 54 + rowBytes * size
        ReDim b(0 To n - 1)
        BmpHeader b, size, size, rowBytes
        cell = size \ 4
        If cell < 2 Then cell = 2
        For y = 0 To size - 1
            For x = 0 To size - 1
                i = 54 + y * rowBytes + x * 3
                r = c And &HFF&: g = (c \ &H100&) And &HFF&: bl = (c \ &H10000) And &HFF&
                If kind = 1 And ((x \ cell) + (y \ cell)) Mod 2 = 0 Then r = 201: g = 206: bl = 214
                If kind = 2 And (x Mod cell = 0 Or y Mod cell = 0) Then r = 154: g = 165: bl = 180
                If x = 0 Or y = 0 Or x = size - 1 Or y = size - 1 Then r = 128: g = 128: bl = 128
                b(i) = bl: b(i + 1) = g: b(i + 2) = r
            Next x
        Next y
        raw = b
        WriteBytesX p, raw
    End If
    Set PatternPicture = LoadPicture(p)
Done:
#End If
End Function

Private Sub PutLong(ByRef b() As Byte, ByVal pos As Long, ByVal v As Long)
    b(pos) = v And &HFF&
    b(pos + 1) = (v \ &H100&) And &HFF&
    b(pos + 2) = (v \ &H10000) And &HFF&
    b(pos + 3) = (v \ &H1000000) And &HFF&
End Sub


' =====================================================================
'  Paths and names (pure string helpers, same on both platforms)
' =====================================================================
Private Function PathSep() As String
#If Mac Then
    PathSep = "/"
#Else
    PathSep = "\"
#End If
End Function

Private Function JoinPath(ByVal folder As String, ByVal fileName As String) As String
    If Right$(folder, 1) = PathSep() Then JoinPath = folder & fileName Else JoinPath = folder & PathSep() & fileName
End Function

Private Function ParentFolder(ByVal p As String) As String
    Dim k As Long
    k = InStrRev(p, PathSep())
    If k > 0 Then ParentFolder = Left$(p, k - 1)
End Function

Private Function FileNameOf(ByVal p As String) As String
    FileNameOf = Mid$(p, InStrRev(p, PathSep()) + 1)
End Function

Private Function ExtOf(ByVal p As String) As String
    Dim nm As String, k As Long
    nm = FileNameOf(p)
    k = InStrRev(nm, ".")
    If k > 0 Then ExtOf = Mid$(nm, k + 1)
End Function

Private Function BaseNameOf(ByVal p As String) As String
    Dim nm As String, k As Long
    nm = FileNameOf(p)
    k = InStrRev(nm, ".")
    If k > 1 Then BaseNameOf = Left$(nm, k - 1) Else BaseNameOf = nm
End Function

Private Function BaseNameFor(ByVal pres As Object, ByVal sr As Object) As String
    Dim s As String, idx As Long
    s = SafeName(BaseNameOf(pres.Name))
    On Error Resume Next
    idx = sr.Parent.SlideIndex
    On Error GoTo 0
    If idx > 0 Then s = s & "_" & Tx("file.page", idx)
    BaseNameFor = s
End Function

Private Function UniquePath(ByVal folder As String, ByVal baseStr As String, ByVal ext As String) As String
    Dim p As String, k As Long
    p = JoinPath(folder, baseStr & "." & ext)
    k = 2
    Do While FileExistsX(p)
        p = JoinPath(folder, baseStr & "_" & k & "." & ext)
        k = k + 1
    Loop
    UniquePath = p
End Function

Private Function SafeName(ByVal s As String) As String
    Dim bad As Variant
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
        s = Replace(s, CStr(bad), "_")
    Next
    s = Trim$(s)
    Do While Len(s) > 0 And Right$(s, 1) = "."
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) > 80 Then s = Left$(s, 80)
    If Len(s) = 0 Then s = "shape"
    SafeName = s
End Function

Private Function DefaultFolder(ByVal pres As Object) As String
    Dim f As String
    f = GetSet("Folder", "")
    If Len(f) > 0 Then
        If FolderExistsX(f) Then DefaultFolder = f: Exit Function
    End If
    f = ""
    On Error Resume Next
    f = pres.Path                                   ' may be a URL for cloud files
    On Error GoTo 0
    If Len(f) > 0 And InStr(f, "://") = 0 Then
        If FolderExistsX(f) Then DefaultFolder = f: Exit Function
    End If
    DefaultFolder = DesktopPathX()
End Function

Private Function InitialFolder(ByVal pres As Object) As String
    Dim f As String
    f = GetSet("LastFolder", "")
    If Len(f) > 0 Then
        If FolderExistsX(f) Then InitialFolder = f: Exit Function
    End If
    InitialFolder = DefaultFolder(pres)
End Function

Private Function FolderSettingText() As String
    Dim f As String
    f = GetSet("Folder", "")
    If Len(f) = 0 Then
        FolderSettingText = Tx("set.folderNone")
    Else
        FolderSettingText = f
    End If
End Function


' =====================================================================
'  Settings (Windows: HKCU\Software\VecStamp, macOS: VBA settings store)
' =====================================================================
Private Function GetSet(ByVal key As String, ByVal defVal As String) As String
    Dim v As Variant
    On Error Resume Next
#If Mac Then
    v = GetSetting(VS_NAME, "Settings", key, defVal)
#Else
    v = CreateObject("WScript.Shell").RegRead(REG_ROOT & key)
#End If
    If Err.Number <> 0 Then v = defVal
    On Error GoTo 0
    GetSet = CStr(v)
End Function

Private Sub PutSet(ByVal key As String, ByVal v As String)
    On Error Resume Next
#If Mac Then
    SaveSetting VS_NAME, "Settings", key, v
#Else
    CreateObject("WScript.Shell").RegWrite REG_ROOT & key, v, "REG_SZ"
#End If
End Sub

Private Sub ResetSettings()
    Dim k As Variant
    On Error Resume Next
    For Each k In Array("Format", "Mode", "Naming", "Dpi", "DpiCustom", "Background", "BgType", "BgColor", _
                        "BgAlpha", "GridStep", "Margin", "OpenFolder", "VisioEdit", "VsdxNative", "RecentColors", _
                        "Folder", "LastFolder")
#If Mac Then
        DeleteSetting VS_NAME, "Settings", CStr(k)
#Else
        CreateObject("WScript.Shell").RegDelete REG_ROOT & CStr(k)
#End If
    Next
End Sub

Private Function GetIdx(ByVal key As String, ByVal maxIdx As Long, Optional ByVal defIdx As Long = 0) As Long
    Dim v As String, r As Long
    v = GetSet(key, CStr(defIdx))
    r = defIdx
    If IsNumeric(v) Then r = CLng(v)
    If r < 0 Or r > maxIdx Then r = defIdx
    GetIdx = r
End Function

Private Function GetFlag(ByVal key As String, ByVal defVal As String) As Boolean
    GetFlag = (GetSet(key, defVal) = "1")
End Function

Private Function GetNum(ByVal key As String, ByVal defVal As Double, ByVal lo As Double, ByVal hi As Double) As Double
    Dim v As Double
    v = Val(GetSet(key, NumText(defVal)))
    If v < lo Or v > hi Then v = defVal
    GetNum = v
End Function

Private Function GetMargin() As Single
    GetMargin = GetNum("Margin", 0, 0, 500)
End Function

Private Function CustomDpi() As Long
    CustomDpi = CLng(GetNum("DpiCustom", 1200, 1, DPI_MAX))
End Function

Private Function DpiValue() As Long
    Select Case GetIdx("Dpi", 3, 1)
        Case 0: DpiValue = 150
        Case 2: DpiValue = 600
        Case 3: DpiValue = CustomDpi()
        Case Else: DpiValue = 300
    End Select
End Function

' Background type; migrates the v1.1 "white background" checkbox.
Private Function BgTypeIdx() As Long
    If Len(GetSet("BgType", "")) = 0 Then
        If GetFlag("Background", "0") Then BgTypeIdx = BG_WHITE Else BgTypeIdx = BG_NONE
    Else
        BgTypeIdx = GetIdx("BgType", BG_CUSTOM, BG_NONE)
    End If
End Function

Private Function BgAlpha() As Long
    BgAlpha = CLng(GetNum("BgAlpha", 0, 0, 100))
End Function

Private Function GridStep() As Single
    GridStep = GetNum("GridStep", 10, 2, 200)
End Function

Private Function BgColorValue(ByVal bg As Long) As Long
    Dim r As Long, g As Long, b As Long
    Select Case bg
        Case BG_BEIGE: BgColorValue = RGB(245, 240, 225)
        Case BG_CUSTOM
            If Not ParseColor(GetSet("BgColor", DEF_CUSTOM_COLOR), r, g, b) Then ParseColor DEF_CUSTOM_COLOR, r, g, b
            BgColorValue = RGB(r, g, b)
        Case Else: BgColorValue = RGB(255, 255, 255)
    End Select
End Function

Private Function ColorSettingText() As String
    Dim r As Long, g As Long, b As Long
    If Not ParseColor(GetSet("BgColor", DEF_CUSTOM_COLOR), r, g, b) Then ParseColor DEF_CUSTOM_COLOR, r, g, b
    ColorSettingText = r & "," & g & "," & b
End Function

' Accepts "255,248,235", "255 248 235", "rgb(255,248,235)", "#FFF8EB", "FFF8EB" or "#FFF".
Private Function ParseColor(ByVal s As String, ByRef r As Long, ByRef g As Long, ByRef b As Long) As Boolean
    Dim t As String, parts() As String, i As Long, n As Long, vals(0 To 2) As Long, ch As String
    t = LCase$(Trim$(s))
    t = Replace(t, "rgb", "")
    t = Replace(t, "(", "")
    t = Replace(t, ")", "")
    t = Replace(t, ChrW(&HFF0C), ",")               ' full-width comma
    If Left$(t, 1) = "#" Then t = Mid$(t, 2)
    If (Len(t) = 6 Or Len(t) = 3) And InStr(t, ",") = 0 And InStr(t, " ") = 0 Then
        For i = 1 To Len(t)
            ch = Mid$(t, i, 1)
            If InStr("0123456789abcdef", ch) = 0 Then Exit Function
        Next i
        If Len(t) = 3 Then t = Mid$(t, 1, 1) & Mid$(t, 1, 1) & Mid$(t, 2, 1) & Mid$(t, 2, 1) & Mid$(t, 3, 1) & Mid$(t, 3, 1)
        r = CLng("&H" & Mid$(t, 1, 2))
        g = CLng("&H" & Mid$(t, 3, 2))
        b = CLng("&H" & Mid$(t, 5, 2))
        ParseColor = True
        Exit Function
    End If
    t = Replace(t, ",", " ")
    t = Replace(t, ";", " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    parts = Split(Trim$(t), " ")
    If UBound(parts) <> 2 Then Exit Function
    For i = 0 To 2
        If Not IsNumeric(parts(i)) Then Exit Function
        n = CLng(Val(parts(i)))
        If n < 0 Or n > 255 Then Exit Function
        vals(i) = n
    Next i
    r = vals(0): g = vals(1): b = vals(2)
    ParseColor = True
End Function

' Parses a number typed into a ribbon edit box (accepts "," as decimal separator).
Private Function ReadNumber(ByVal s As String, ByVal lo As Double, ByVal hi As Double, ByRef v As Double) As Boolean
    s = Trim$(Replace(s, ",", "."))
    If Len(s) = 0 Then s = CStr(lo)
    If Not IsNumeric(Replace(s, ".", "")) Then Exit Function
    v = Val(s)
    If v < lo Or v > hi Then Exit Function
    ReadNumber = True
End Function

Private Function NumText(ByVal v As Double) As String
    NumText = Trim$(Str$(v))
    If Left$(NumText, 1) = "." Then NumText = "0" & NumText
End Function

' Number for XML (always "." as decimal separator).
Private Function NumXml(ByVal v As Double) As String
    Dim neg As Boolean, ip As Double, fp As Long, r As String
    v = Round(v, 6)
    If Abs(v) < 0.0000005 Then NumXml = "0": Exit Function
    neg = (v < 0)
    v = Abs(v)
    ip = Int(v)
    fp = CLng((v - ip) * 1000000#)
    If fp >= 1000000 Then ip = ip + 1: fp = fp - 1000000
    r = Trim$(Str$(ip))
    If fp > 0 Then
        r = r & "." & Right$("000000" & Trim$(Str$(fp)), 6)
        Do While Right$(r, 1) = "0"
            r = Left$(r, Len(r) - 1)
        Loop
    End If
    If neg Then r = "-" & r
    NumXml = r
End Function

Private Function MinS(ByVal a As Single, ByVal b As Single) As Single
    If a < b Then MinS = a Else MinS = b
End Function

Private Function MaxS(ByVal a As Single, ByVal b As Single) As Single
    If a > b Then MaxS = a Else MaxS = b
End Function


' =====================================================================
'  Formats and small UI helpers
' =====================================================================
Private Function FmtExt(ByVal fmt As Long) As String
    Select Case fmt
        Case FMT_EMZ: FmtExt = "emz"
        Case FMT_SVG: FmtExt = "svg"
        Case FMT_WMF: FmtExt = "wmf"
        Case FMT_PNG: FmtExt = "png"
        Case FMT_PDF: FmtExt = "pdf"
        Case FMT_VSDX: FmtExt = "vsdx"
        Case Else: FmtExt = "emf"
    End Select
End Function

Private Function FmtFromExt(ByVal ext As String) As Long
    Select Case ext
        Case "emf": FmtFromExt = FMT_EMF
        Case "emz": FmtFromExt = FMT_EMZ
        Case "svg": FmtFromExt = FMT_SVG
        Case "wmf": FmtFromExt = FMT_WMF
        Case "png": FmtFromExt = FMT_PNG
        Case "pdf": FmtFromExt = FMT_PDF
        Case "vsdx": FmtFromExt = FMT_VSDX
        Case Else: FmtFromExt = -1
    End Select
End Function

Private Function FilterLabel(ByVal fmt As Long) As String
    Select Case fmt
        Case 0 To FMT_LAST: FilterLabel = Tx("fmt." & FmtExt(fmt)) & " (*." & FmtExt(fmt) & ")"
        Case Else: FilterLabel = Tx("fmt.emf") & " (*.emf)"
    End Select
End Function

Private Function BgShortName(ByVal bg As Long) As String
    Select Case bg
        Case BG_WHITE: BgShortName = Tx("bg.white")
        Case BG_BEIGE: BgShortName = Tx("bg.beige")
        Case BG_GRID: BgShortName = Tx("bg.grid")
        Case BG_CUSTOM: BgShortName = Tx("bg.custom")
        Case Else: BgShortName = Tx("bg.none")
    End Select
End Function

Private Function BgName(ByVal bg As Long) As String
    Select Case bg
        Case BG_WHITE: BgName = Tx("bg.white")
        Case BG_BEIGE: BgName = Tx("bg.beige")
        Case BG_GRID: BgName = Tx("bg.gridValue", NumText(GridStep()))
        Case BG_CUSTOM: BgName = Tx("bg.customValue", ColorSettingText())
        Case Else: BgName = Tx("bg.none")
    End Select
End Function

Private Function SettingsSummary() As String
    Dim s As String, bg As Long
    s = Tx("sum.head") & vbCrLf & vbCrLf
    s = s & Tx("sum.format", UCase$(FmtExt(GetIdx("Format", FMT_LAST)))) & vbCrLf
    s = s & Tx("sum.mode", Tx(IIf(GetIdx("Mode", 1) = 1, "mode.split", "mode.merge"))) & vbCrLf
    s = s & Tx("sum.naming", Tx(IIf(GetIdx("Naming", 1) = 1, "naming.auto", "naming.ask"))) & vbCrLf
    s = s & Tx("sum.dpi", DpiValue()) & vbCrLf
    bg = BgTypeIdx()
    s = s & Tx("sum.bg", BgName(bg))
    If bg <> BG_NONE Then s = s & Tx("sum.alpha", BgAlpha())
    s = s & vbCrLf & Tx("sum.margin", NumText(GetMargin()))
    s = s & vbCrLf & Tx(IIf(GetFlag("VsdxNative", "1"), "sum.vsdxNative", "sum.vsdxPicture"))
    s = s & vbCrLf & Tx("sum.lang", LangName(CurLang()))
    SettingsSummary = s
End Function

Private Sub Report(ByVal files As Collection, ByVal showList As Boolean)
    Dim msg As String, i As Long
    If files.Count = 0 Then Exit Sub
    PutSet "LastFolder", ParentFolder(CStr(files(1)))
    RefreshRibbon
    If GetFlag("OpenFolder", "0") Then
        RevealFileX CStr(files(1))
        Exit Sub
    End If
    If Not showList Then Exit Sub
    msg = Tx("msg.exported", files.Count) & vbCrLf & vbCrLf
    For i = 1 To files.Count
        If i > 8 Then msg = msg & "……" & vbCrLf: Exit For
        msg = msg & CStr(files(i)) & vbCrLf
    Next i
    MsgBox msg, vbInformation, AppTitle()
End Sub

' After Esc / Cancel: says so and lists the files finished before.
Private Sub ReportCancelled(ByVal files As Collection)
    Dim msg As String, i As Long
    If files.Count = 0 Then
        MsgBox Tx("msg.cancelled"), vbInformation, AppTitle()
        Exit Sub
    End If
    PutSet "LastFolder", ParentFolder(CStr(files(1)))
    msg = Tx("msg.cancelledPartial", files.Count) & vbCrLf & vbCrLf
    For i = 1 To files.Count
        If i > 8 Then msg = msg & "……" & vbCrLf: Exit For
        msg = msg & CStr(files(i)) & vbCrLf
    Next i
    MsgBox msg, vbInformation, AppTitle()
End Sub

Private Sub RefreshRibbon()
    On Error Resume Next
    If Not gRibbon Is Nothing Then gRibbon.Invalidate
End Sub

Private Function AppTitle() As String
    AppTitle = Tx("app.title")
End Function


' =====================================================================
'  Interface languages: Simplified Chinese (default), Traditional Chinese, English, Japanese.
'  The texts live in src/i18n/strings.json; the build turns them into LoadStrings (generated
'  block below), which fills gStr with the texts of one language.
' =====================================================================
Private Function CurLang() As String
    If Len(gLang) = 0 Then
        gLang = GetSet("Lang", "zh-CN")
        If InStr("|zh-CN|zh-TW|en|ja|", "|" & gLang & "|") = 0 Then gLang = "zh-CN"
    End If
    CurLang = gLang
End Function

' Switches the interface language at once (the caller refreshes the ribbon).
Private Sub SetLang(ByVal code As String)
    If InStr("|zh-CN|zh-TW|en|ja|", "|" & code & "|") = 0 Then code = "zh-CN"
    PutSet "Lang", code
    gLang = code
    Set gStr = Nothing
End Sub

' Each language under its own name (not translated).
Private Function LangName(ByVal code As String) As String
    Select Case code
        Case "zh-TW": LangName = "繁體中文"
        Case "en": LangName = "English"
        Case "ja": LangName = "日本語"
        Case Else: LangName = "简体中文"
    End Select
End Function

Private Sub EnsureStrings()
    If gStr Is Nothing Then
        Set gStr = New Collection
        LoadStrings CurLang()
    End If
End Sub

' The text of key in the current language with {0}, {1}, ... replaced by args (the key itself
' when it is missing). Err is preserved: Tx is often called inside error handlers.
Private Function Tx(ByVal key As String, ParamArray args() As Variant) As String
    Dim s As String, i As Long, en As Long, ed As String, es As String
    en = Err.Number: ed = Err.Description: es = Err.Source
    EnsureStrings
    On Error Resume Next
    s = gStr.Item(key)
    If Err.Number <> 0 Then s = key
    On Error GoTo 0
    For i = 0 To UBound(args)
        s = Replace(s, "{" & i & "}", CStr(args(i)))
    Next i
    If en <> 0 Then
        Err.Number = en
        Err.Description = ed
        Err.Source = es
    End If
    Tx = s
End Function

Private Function TxHas(ByVal key As String) As Boolean
    Dim s As String, en As Long, ed As String, es As String
    en = Err.Number: ed = Err.Description: es = Err.Source
    EnsureStrings
    On Error Resume Next
    s = gStr.Item(key)
    TxHas = (Err.Number = 0)
    On Error GoTo 0
    If en <> 0 Then
        Err.Number = en
        Err.Description = ed
        Err.Source = es
    End If
End Function

' Decodes UTF-16 code units written as 4-digit hex groups (used by the ASCII-only build).
Private Function U(ByVal h As String) As String
    Dim i As Long, s As String
    For i = 1 To Len(h) Step 4
        s = s & ChrW$(CLng("&H" & Mid$(h, i, 4)))
    Next i
    U = s
End Function


' >>> GENERATED: VSDX TEMPLATE PARTS, HELPER SCRIPT AND PRESET SHAPES (build/build_bas.py fills
'     this block from src/vsdx, src/helpers and src/presets)
Private Function VsdxPart(ByVal partName As String) As String
    VsdxPart = ""
End Function

Private Function VsdxMap() As Variant
    VsdxMap = Array()
End Function

Private Function HelperScript() As String
    HelperScript = ""
End Function

Private Function PresetSource(ByVal ast As Long) As String
    PresetSource = ""
End Function

Private Sub LoadStrings(ByVal lang As String)
End Sub
' <<< GENERATED

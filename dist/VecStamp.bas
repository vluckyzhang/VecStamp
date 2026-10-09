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
Private gPgWnd As LongPtr, gPgMsg As LongPtr, gPgBar As LongPtr, gPgSub As LongPtr, gPgFont As LongPtr, gPgFontB As LongPtr
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
Private gPgWnd As Long, gPgMsg As Long, gPgBar As Long, gPgSub As Long, gPgFont As Long, gPgFontB As Long
#End If
#End If

' ---- product information
Private Const VS_NAME As String = "VecStamp"
Private Const VS_VERSION As String = "1.2.1"
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
#End If

Private gRibbon As Object
Private gSvgFallback As Long        ' per run: 0 = not asked yet, 1 = save EMF instead, 2 = skip
Private gSvgNative As Long          ' per run: 0 = untested, 1 = PowerPoint writes valid SVG, 2 = it does not
Private gBigPngOk As Boolean        ' per run: user accepted a very large PNG
Private gCustColors(0 To 15) As Long
Private gBusy As Boolean            ' an export is running (ribbon clicks during DoEvents are ignored)
' progress window
Private gPgActive As Boolean, gPgPos As Long, gPgTotal As Long, gPgTitle As String
#If Mac Then
Private gPgMacPid As String
#End If
' native VSDX writer
Private gVxSb() As String, gVxSbN As Long, gVxId As Long, gVxMedia As Collection
Private gVxScratch As Object, gVxPres As Object, gVxNative As Long, gVxPictures As Long, gVxCopy As Long
Private gGeo As String, gGeoSec As String, gGeoIx As Long, gGeoRow As Long
Private gGeoW As Double, gGeoH As Double, gGeoFlipH As Boolean, gGeoFlipV As Boolean
Private gGeoX0 As Double, gGeoY0 As Double, gGeoLX As Double, gGeoLY As Double
Private gGeoNoFill As Boolean, gGeoNoLine As Boolean, gGeoOpenOnly As Boolean, gGeoTextBox As Boolean


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

' ---- background colour gallery (theme / standard / recent colours)
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

Public Sub VS_GalColorMainImage(control As Object, ByRef image)
#If Mac Then
#Else
    Set image = SwatchPicture(BgColorValue(BG_CUSTOM), 32)
#End If
End Sub

' ---- donations, project page
Public Sub VS_OnDonate(control As Object, id As String, index As Integer)
    If index = 0 Then DoDonateThanks U("652F4ED85B9D") Else DoDonateThanks U("5FAE4FE1")
End Sub

Public Sub VS_OnDonateInfo(control As Object)
    DoAbout
End Sub

Public Sub VS_OnHome(control As Object)
    If Not OpenUrlX(VS_REPO) Then MsgBox U("987976EE4E3B9875FF1A") & VS_REPO, vbInformation, AppTitle()
End Sub

Public Sub VS_OnOpenFolder(control As Object)
    Dim f As String
    f = GetSet("LastFolder", "")
    If Len(f) = 0 Or Not FolderExistsX(f) Then f = GetSet("Folder", "")
    If Len(f) = 0 Or Not FolderExistsX(f) Then
        MsgBox U("8FD86CA167095BFC51FA8FC765874EF630025BFC51FA4E006B21540EFF0C8FD991CC4F1A62535F0067008FD14F7F7528768465874EF659393002"), vbInformation, AppTitle()
        Exit Sub
    End If
    OpenPathX f
End Sub

Public Sub VS_GetExportLabel(control As Object, ByRef label)
    label = U("5BFC51FA0020") & UCase$(FmtExt(GetIdx("Format", FMT_LAST)))
End Sub

Public Sub VS_GetIndex(control As Object, ByRef index)
    Select Case control.Id
        Case "vsDdFormat": index = GetIdx("Format", FMT_LAST)
        Case "vsDdMode": index = GetIdx("Mode", 1)
        Case "vsDdNaming": index = GetIdx("Naming", 1)
        Case "vsDdDpi": index = GetIdx("Dpi", 3, 1)
        Case "vsDdBg": index = BgTypeIdx()
        Case Else: index = 0
    End Select
End Sub

Public Sub VS_OnDropDown(control As Object, id As String, index As Integer)
    Select Case control.Id
        Case "vsDdFormat": PutSet "Format", CStr(index)
        Case "vsDdMode": PutSet "Mode", CStr(index)
        Case "vsDdNaming": PutSet "Naming", CStr(index)
        Case "vsDdDpi": PutSet "Dpi", CStr(index)
        Case "vsDdBg": PutSet "BgType", CStr(index)
    End Select
    RefreshRibbon
End Sub

Public Sub VS_GetPressed(control As Object, ByRef returnedVal)
    Select Case control.Id
        Case "vsChkOpen": returnedVal = GetFlag("OpenFolder", "0")
        Case "vsChkVsdx": returnedVal = GetFlag("VsdxNative", "1")
        Case Else: returnedVal = False
    End Select
End Sub

Public Sub VS_OnCheck(control As Object, pressed As Boolean)
    Dim v As String
    If pressed Then v = "1" Else v = "0"
    Select Case control.Id
        Case "vsChkOpen": PutSet "OpenFolder", v
        Case "vsChkVsdx": PutSet "VsdxNative", v
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
                Else MsgBox U("8FB98DDD8BF78F93516500200030FF5E00350030003000204E4B95F4768465705B57FF0853554F4DFF1A78C5FF093002"), vbExclamation, AppTitle()
        Case "vsEbDpi"
            If ReadNumber(strText, 1, DPI_MAX, v) Then
                PutSet "DpiCustom", CStr(CLng(v))
            Else
                MsgBox U("81EA5B9A4E4952068FA873878BF78F93516500200031FF5E") & DPI_MAX & U("00204E4B95F4768465746570FF08004400500049FF093002"), vbExclamation, AppTitle()
            End If
        Case "vsEbAlpha"
            If ReadNumber(strText, 0, 100, v) Then PutSet "BgAlpha", CStr(CLng(v)) _
                Else MsgBox U("80CC666F900F660E5EA68BF78F93516500200030FF5E00310030003000204E4B95F4768465705B57FF080025FF093002"), vbExclamation, AppTitle()
        Case "vsEbGrid"
            If ReadNumber(strText, 2, 200, v) Then PutSet "GridStep", NumText(v) _
                Else MsgBox U("7F51683C95F48DDD8BF78F93516500200032FF5E00320030003000204E4B95F4768465705B57FF0853554F4DFF1A78C5FF093002"), vbExclamation, AppTitle()
        Case "vsEbRgb"
            If ParseColor(strText, r, g, b) Then
                PutSet "BgColor", r & "," & g & "," & b
                PutSet "BgType", CStr(BG_CUSTOM)
            Else
                MsgBox U("989C8272683C5F0F793A4F8BFF1A003200350035002C003200340038002C0032003300350020621600200023004600460046003800450042"), vbExclamation, AppTitle()
            End If
    End Select
    RefreshRibbon
End Sub

Public Sub VS_GetSupertip(control As Object, ByRef tip)
    Select Case control.Id
        Case "vsBtnFolder"
            tip = U("201C81EA52A84FDD5B58201D548C201C901098755BFC51FA201D9ED88BA44F7F7528768465874EF65939FF1A") & vbCrLf & FolderSettingText()
        Case "vsBtnOpenFolder"
            tip = U("5728") & FileManagerName() & U("4E2D62535F0067008FD14E006B215BFC51FA62405728768465874EF659393002")
        Case "vsGalColor", "vsEbRgb"
            tip = U("5F53524D81EA5B9A4E4980CC666F8272FF1A") & ColorSettingText() & vbCrLf & _
                  U("4ECE4E3B9898989C82723001680751C68272300167008FD14F7F75287684989C82724E2D900962E9FF0C62167528201C51764ED6989C8272201D201C53D682725668201DFF1B") & _
                  U("4E5F53EF4EE557280020005200470042002068464E2D8F93516500200052002C0047002C0042FF0859820020003200350035002C003200340038002C003200330035FF0962165341516D8FDB5236FF08598200200023004600460046003800450042FF09300290098272540E80CC666F81EA52A8520763624E3A201C81EA5B9A4E49989C8272201D3002")
        Case "vsChkVsdx"
            tip = U("52FE9009FF1A005600530044005800204E2D768456FE5F62662F00200056006900730069006F0020539F751F5F6272B6FF0C65875B5753EF76F463A57F168F91FF0865E097005B8988C500200056006900730069006FFF093002") & vbCrLf & _
                  U("4E0D52FE9009FF1A65744E2A56FE5F624F5C4E3A4E005F2000200045004D0046002077E291CF56FE72475D4C516500200056005300440058FF0C591689C24E0E00200050006F0077006500720050006F0069006E007400205B8C51684E0081F43002")
        Case Else
            tip = SettingsSummary()
    End Select
End Sub

Public Sub VS_OnFolder(control As Object)
    Dim cur As String, f As String
    cur = GetSet("Folder", "")
    If Len(cur) = 0 Or Not FolderExistsX(cur) Then cur = DesktopPathX()
    f = PickFolderX(cur, U("900962E9201C81EA52A84FDD5B58201D4F7F752876849ED88BA465874EF65939"))
    If Len(f) > 0 Then
        PutSet "Folder", f
        RefreshRibbon
    End If
End Sub

Public Sub VS_OnReset(control As Object)
    If MsgBox(U("786E5B9A628A624067095BFC51FA8BBE7F6E6062590D4E3A9ED88BA4503C5417FF1F"), vbQuestion + vbYesNo, AppTitle()) <> vbYes Then Exit Sub
    ResetSettings
    RefreshRibbon
End Sub

Public Sub VS_OnSettings(control As Object)
    MsgBox SettingsSummary(), vbInformation, AppTitle() & " - " & U("5F53524D8BBE7F6E")
End Sub

Public Sub VS_OnAbout(control As Object)
    DoAbout
End Sub

Public Sub VS_OnFeedback(control As Object)
    DoFeedback
End Sub

Public Sub VS_OnHelp(control As Object)
    MsgBox U("77E253700020005600650063005300740061006D007000204F7F75288BF4660E") & vbCrLf & vbCrLf & _
        U("2460002090094E2D4E004E2A6216591A4E2A56FE5F62FF0C70B9201C5BFC51FA201DFF0C621676F463A570B95404683C5F0F630994AE30026CA1670990094E2D56FE5F6265F6FF0C53EF4EE55BFC51FA657498753002") & vbCrLf & _
        U("24610020901098755BFC51FAFF1A6BCF98755E7B706F72474E004E2A65874EF6FF1B57285DE64FA77F29756556FE4E2D591A900998759762521953EA5BFC51FA6240900998753002") & vbCrLf & _
        U("246200200045004D0046002056FE7247526F672CFF1A628A6240900956FE5F62590D5236621077E291CF56FE7247653E572865C18FB9FF0C539F56FE5F624E0D53D83002") & vbCrLf & vbCrLf & _
        U("00B700200045004D00460020002F00200045004D005A0020002F00200057004D00460020002F0020005300560047FF1A77E291CF56FE30020050006F0077006500720050006F0069006E007400200032003000310033002053CA4EE54E0A90FD80FD5BFC51FA0020005300560047FF08003200330030003200204E4B524D76847248672C7ECF0020005800500053002051857F6E8F6C6362FF093002") & vbCrLf & _
        U("00B70020005000440046FF1A9875976259275C0F4E0E56FE5F624E0081F4768477E291CF0020005000440046FF0C900254080020004C00610054006500583002") & vbCrLf & _
        U("00B700200056005300440058FF1A0056006900730069006F002065874EF6FF0C56FE5F624E3A00200056006900730069006F0020539F751F5F6272B6300165875B5753EF7F168F91FF0C65E097005B8988C500200056006900730069006F3002") & vbCrLf & _
        U("00B700200050004E0047FF1A0031003500300020002F00200033003000300020002F002000360030003000200044005000490020621681EA5B9A4E49FF0867009AD8002000330030003000300020004400500049FF093002") & vbCrLf & _
        U("00B7002080CC666FFF1A900F660E3001767D827230017C73827230017F51683C621681EA5B9A4E49989C8272FF084E3B989882720020002F0020680751C682720020002F002051764ED6989C82720020002F002053D6827256680020002F0020005200470042FF09FF0C53EF8BBE900F660E5EA6548C8FB98DDD3002") & vbCrLf & _
        U("00B700205BFC51FA65F64F1A663E793A8FDB5EA67A9753E3FF0C63090020004500730063002053EF4EE553D66D883002") & vbCrLf & vbCrLf & _
        U("54085E765BFC51FA300152A080CC666F30010050004400460020548C901098755BFC51FA65F64F1A4E3465F6501F7528526A8D34677FFF0C5E7657286587672B4E3465F663D251654E009875FF085B8C6210540E81EA52A852209664FF093002"), _
        vbInformation, AppTitle()
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
            folder = PickFolderX(InitialFolder(pres), U("900962E95BFC51FA65874EF65939FF086BCF4E2A56FE5F624E004E2A65874EF6FF09"))
            If Len(folder) = 0 Then Exit Sub
        End If
        gBusy = True
        ProgressBegin U("6B6357285BFC51FA0020") & sr.Count & U("00204E2A0020") & UCase$(FmtExt(fmt)) & U("002065874EF6"), sr.Count * 4
        For i = 1 To sr.Count
            If ProgressCancelled() Then Exit For
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
        ProgressBegin U("6B6357285BFC51FA0020") & UCase$(FmtExt(fmt)), 4
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
    If Err.Number <> vbObjectError + 599 Then MsgBox U("5BFC51FA59318D25FF1A") & Err.Description, vbExclamation, AppTitle()
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
        MsgBox U("8BF7514862535F006F14793A65877A3F3002"), vbInformation, AppTitle()
        Exit Sub
    End If
    Set pres = win.Presentation
    fmt = GetIdx("Format", FMT_LAST)

    If win.Selection.Type = PP_SEL_SLIDES Then
        If win.Selection.SlideRange.Count > 1 Then
            ans = MsgBox(U("89815BFC51FA54EA4E9B5E7B706F7247FF1F") & vbCrLf & vbCrLf & _
                         U("662FFF1A53EA5BFC51FA6240900976840020") & win.Selection.SlideRange.Count & U("00209875") & vbCrLf & _
                         U("5426FF1A5BFC51FA516890E80020") & pres.Slides.Count & U("00209875"), _
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
        MsgBox U("6F14793A65877A3F4E2D6CA167095E7B706F72473002"), vbInformation, AppTitle()
        Exit Sub
    End If

    If GetIdx("Naming", 1) = 1 Then
        folder = DefaultFolder(pres)
    Else
        folder = PickFolderX(InitialFolder(pres), U("900962E9901098755BFC51FA768465874EF65939FF08683C5F0FFF1A") & UCase$(FmtExt(fmt)) & U("FF09"))
        If Len(folder) = 0 Then Exit Sub
    End If

    ResetRunState
    baseName = SafeName(BaseNameOf(pres.Name))
    If gBusy Then Exit Sub
    gBusy = True
    ProgressBegin U("901098755BFC51FAFF08") & slideList.Count & U("00209875FF0C") & UCase$(FmtExt(fmt)) & U("FF09"), slideList.Count * 4
    For i = 1 To slideList.Count
        If ProgressCancelled() Then Exit For
        Set sld = slideList(i)
        ProgressStep U("7B2C0020") & i & " / " & slideList.Count & U("00209875")
        Set rng = ExportableRange(sld)
        If rng Is Nothing Then
            skipped = skipped + 1
        Else
            path = UniquePath(folder, baseName & U("005F7B2C") & sld.SlideIndex & U("9875"), FmtExt(fmt))
            p = ExportItem(pres, rng, path, fmt)
            If Len(p) > 0 Then files.Add p
        End If
    Next i
    ProgressEnd
    gBusy = False

    If files.Count = 0 Then
        MsgBox U("624090095E7B706F72474E0A6CA1670953EF5BFC51FA768456FE5F623002"), vbInformation, AppTitle()
        Exit Sub
    End If
    If skipped > 0 Then
        MsgBox U("5DF28DF38FC70020") & skipped & U("00205F207A7A767D5E7B706F72473002"), vbInformation, AppTitle()
    End If
    Report files, True
    Exit Sub
EH:
    ProgressEnd
    gBusy = False
    If Err.Number <> vbObjectError + 599 Then MsgBox U("901098755BFC51FA59318D25FF1A") & Err.Description, vbExclamation, AppTitle()
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
    MsgBox U("751F621000200045004D0046002056FE7247526F672C59318D25FF1A") & Err.Description, vbExclamation, AppTitle()
End Sub

' Colour picker for the custom background (Windows: system colour dialog, macOS: system colour picker).
Private Sub DoPickColor()
    Dim r As Long, g As Long, b As Long
    On Error GoTo EH
    If Not ParseColor(GetSet("BgColor", DEF_CUSTOM_COLOR), r, g, b) Then ParseColor DEF_CUSTOM_COLOR, r, g, b
#If Mac Then
    Dim res As String
    res = MacScript2("chooseColor", r & "," & g & "," & b)
    If Len(res) = 0 Then Exit Sub
    If Not ParseColor(res, r, g, b) Then Exit Sub
#Else
    Dim cc As CHOOSECOLORW
    cc.lStructSize = LenB(cc)
    cc.hwndOwner = GetActiveWindow()
    cc.rgbResult = RGB(r, g, b)
    cc.lpCustColors = VarPtr(gCustColors(0))
    cc.Flags = CC_RGBINIT Or CC_FULLOPEN Or CC_ANYCOLOR
    If ChooseColorW(cc) = 0 Then Exit Sub
    r = cc.rgbResult And &HFF&
    g = (cc.rgbResult \ &H100&) And &HFF&
    b = (cc.rgbResult \ &H10000) And &HFF&
#End If
    PutSet "BgColor", r & "," & g & "," & b
    PutSet "BgType", CStr(BG_CUSTOM)
    RefreshRibbon
    Exit Sub
EH:
    MsgBox U("65E06CD562535F00989C8272900962E95668FF1A") & Err.Description & vbCrLf & U("4E5F53EF4EE576F463A55728201C005200470042201D68464E2D8F935165989C82723002"), vbExclamation, AppTitle()
End Sub

Private Sub ResetRunState()
    gSvgFallback = 0
    gSvgNative = 0
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
        MsgBox U("8BF7514862535F006F14793A65877A3FFF0C5E7690094E2D89815BFC51FA768456FE5F623002"), vbInformation, AppTitle()
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
            MsgBox U("8BF7514890094E2D56FE5F623002"), vbInformation, AppTitle()
            Exit Function
        End If
        On Error Resume Next
        Set sld = win.View.Slide
        On Error GoTo 0
        If sld Is Nothing Then
            MsgBox U("8BF75728666E901A89C656FE4E2D90094E2D89815BFC51FA768456FE5F623002"), vbInformation, AppTitle()
            Exit Function
        End If
        Set sr = ExportableRange(sld)
        If sr Is Nothing Then
            MsgBox U("5F53524D5E7B706F72474E0A6CA1670953EF5BFC51FA768456FE5F623002"), vbInformation, AppTitle()
            Exit Function
        End If
        If MsgBox(U("6CA1670990094E2D4EFB4F5556FE5F623002") & vbCrLf & U("662F54265BFC51FA5F53524D5E7B706F72474E0A7684516890E856FE5F62FF1F"), _
                  vbQuestion + vbYesNo, AppTitle()) <> vbYes Then Exit Function
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
    ProgressStep U("51C6590756FE5F62")

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
    ProgressStep U("5408621056FE5F62300180CC666F4E0E8FB98DDD")
    Set tmp = AddTempSlide(pres, src)
    Set target = BuildComposite(tmp, src, margin)
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
    ProgressStep U("5BFC51FA0020") & UCase$(FmtExt(fmt))
    Select Case fmt
        Case FMT_EMF
            ExportRaw shp, path, PP_FMT_EMF
        Case FMT_WMF
            ExportRaw shp, path, PP_FMT_WMF
        Case FMT_EMZ
            t = TempFileX("emf")
            ExportToTemp shp, t, PP_FMT_EMF, 0, 0
            ProgressStep U("538B7F294E3A00200045004D005A")
            If Not GzipX(t, path) Then Err.Raise vbObjectError + 514, , U("0045004D005A0020538B7F2959318D253002")
        Case FMT_PNG
            dpi = DpiValue()
            px = (shp.Width / 72# * dpi) * (shp.Height / 72# * dpi)
            If px > PNG_WARN_PIXELS And Not gBigPngOk Then
                If AskUser(U("63090020") & dpi & U("002000440050004900205BFC51FAFF0C8FD95F2056FE7EA64E3A0020") & CLng(shp.Width / 72# * dpi) & U("002000D70020") & _
                           CLng(shp.Height / 72# * dpi) & U("002050CF7D20FF087EA60020") & Format$(px / 1000000#, "0") & _
                           U("0020767E4E0750CF7D20FF09FF0C53EF80FD5F886162621651855B584E0D8DB330024ECD71367EE77EED5417FF1F"), vbQuestion + vbYesNo) <> vbYes Then
                    Err.Raise vbObjectError + 599, , "cancelled"
                End If
                gBigPngOk = True
            End If
            ExportRaw shp, path, PP_FMT_PNG, _
                      CLng(pres.PageSetup.SlideWidth * dpi / 96), CLng(pres.PageSetup.SlideHeight * dpi / 96)
        Case FMT_SVG
            path = ExportSvg(pres, shp, path)
        Case FMT_PDF
            ExportPdf shp, path
        Case FMT_VSDX
            ProgressStep U("751F621000200056006900730069006F00205F6272B6")
            ExportVsdx pres, shp, path
    End Select
    If Len(path) > 0 Then
        If Not FileExistsX(path) Then Err.Raise vbObjectError + 515, , U("6CA16709751F621065874EF6FF1A") & path
    End If
    ExportShape = path
End Function

' SVG, best method first:
'  1. PowerPoint's own SVG export (2302+ / Mac 16.82+), checked and repaired;
'  2. Windows, any PowerPoint 2013+: PowerPoint prints the drawing to XPS and the built-in
'     converter turns it into SVG (paths, gradients, images, real text, embedded fonts);
'  3. EMF converted with Inkscape or LibreOffice, when installed;
'  4. otherwise offer EMF instead.
Private Function ExportSvg(ByVal pres As Object, ByVal shp As Object, ByVal path As String) As String
    Dim ok As Boolean, t As String, e As String, p2 As String

    If gSvgNative <> 2 Then
        t = TempFileX("svg")
        On Error Resume Next
        ExportToTemp shp, t, PP_FMT_SVG, 0, 0
        ok = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
        If ok Then ok = FixSvgFile(t)
        If ok Then
            gSvgNative = 1
            MoveTempTo t, path
            ExportSvg = path
            Exit Function
        End If
        gSvgNative = 2                              ' this PowerPoint cannot write SVG
        DeleteTempX t
    End If

    ProgressStep U("901A8FC700200058005000530020751F62100020005300560047")
    t = TempFileX("svg")
    ok = XpsToSvg(pres, shp, t)
    If ok Then ok = FixSvgFile(t)
    If ok Then
        MoveTempTo t, path
        ExportSvg = path
        Exit Function
    End If
    DeleteTempX t

    ProgressStep U("4F7F752800200049006E006B007300630061007000650020002F0020004C0069006200720065004F0066006600690063006500208F6C6362")
    e = TempFileX("emf")
    t = TempFileX("svg")
    ExportToTemp shp, e, PP_FMT_EMF, 0, 0
    ok = ConvertEmfToSvg(e, t)
    If ok Then ok = FixSvgFile(t)
    DeleteTempX e
    If ok Then
        MoveTempTo t, path
        ExportSvg = path
        Exit Function
    End If
    DeleteTempX t

    If gSvgFallback = 0 Then
        If AskUser(U("6CA180FD751F62100020005300560047FF1A8FD94E2A00200050006F0077006500720050006F0069006E007400204E0D80FD76F463A55BFC51FA0020005300560047FF0C51857F6E8F6C63624E5F6CA167096210529FFF08970089817CFB7EDF81EA5E267684002000570069006E0064006F0077007300200050006F007700650072005300680065006C006CFF09FF0C") & _
                   U("753581114E0A4E5F6CA1670953EF75284E8E8F6C6362768400200049006E006B00730063006100700065002062160020004C0069006200720065004F006600660069006300653002") & _
                   vbCrLf & vbCrLf & U("662F542665394E3A5BFC51FA00200045004D0046FF1FFF084E5F53EF4EE55BFC51FA00200050004400460020540E752851764ED68F6F4EF68F6C6362FF09"), _
                   vbQuestion + vbYesNo) = vbYes Then
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
    ProgressStep U("751F62100020005000440046")

    On Error GoTo Failed
    Set tp = Application.Presentations.Add(MSO_FALSE)
    tp.PageSetup.SlideWidth = pw
    tp.PageSetup.SlideHeight = ph
    Set sld = tp.Slides.Add(1, PP_LAYOUT_BLANK)
    Set pic = PasteFromClipboard(sld, shp, True).Item(1)
    pic.LockAspectRatio = MSO_FALSE
    pic.Width = w * k
    pic.Height = h * k
    pic.Left = (pw - w * k) / 2
    pic.Top = (ph - h * k) / 2
    tp.SaveAs t, PP_SAVE_AS_PDF

Cleanup:
    On Error Resume Next
    If Not tp Is Nothing Then
        tp.Saved = MSO_TRUE
        tp.Close
    End If
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, , U("00500044004600205BFC51FA59318D25FF1A") & errDesc
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
    e = TempFileX("emf")
    ExportToTemp shp, e, PP_FMT_EMF, 0, 0
    gVxId = 0
    Set gVxMedia = New Collection
    SbReset
    SbAdd VxForeignXml(e, 0, 0, shp.Width, shp.Height, "VecStamp")
    WriteVsdxPackage shp.Width / 72#, shp.Height / 72#, SbText(), path
    DeleteTempX e
End Sub

Private Function VsdxNative(ByVal pres As Object, ByVal shp As Object, ByVal path As String) As Boolean
    Dim wasSaved As Long, ok As Boolean, i As Long
    wasSaved = pres.Saved
    Set gVxPres = pres
    gVxId = 0
    gVxNative = 0
    gVxPictures = 0
    gVxCopy = 0
    Set gVxMedia = New Collection
    SbReset
    On Error GoTo Failed
    Set gVxScratch = AddTempSlide(pres, shp)
    VxShape shp, shp.Left, shp.Top + shp.Height
    If gVxNative = 0 And gVxPictures = 0 Then GoTo Cleanup
    ProgressStep U("519951FA00200056006900730069006F002065874EF6")
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
    VsdxNative = ok
    Exit Function
Failed:
    ok = False
    Resume Cleanup
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
        VxGeometry = VxPreset(s, ast, w, h)
        If Not VxGeometry Then VxGeometry = VxMerged(s, False)
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
End Sub

Private Sub GeoLine(ByVal ux As Double, ByVal v As Double)
    Dim x As Double, y As Double
    GeoXY ux, v, x, y
    gGeoRow = gGeoRow + 1
    gGeoSec = gGeoSec & "<Row T='LineTo' IX='" & gGeoRow & "'>" & VxCell("X", NumXml(x)) & VxCell("Y", NumXml(y)) & "</Row>"
    gGeoLX = x: gGeoLY = y
End Sub

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

' Elliptic arc as cubic Bezier segments. Angles in radians, y downwards (positive sweep = clockwise).
Private Sub GeoArc(ByVal cu As Double, ByVal cv As Double, ByVal ru As Double, ByVal rv As Double, _
                   ByVal t0 As Double, ByVal sweep As Double, ByVal startNew As Boolean)
    Dim n As Long, d As Double, k As Double, i As Long, a As Double, b As Double
    If startNew Then
        GeoMove cu + ru * Cos(t0), cv + rv * Sin(t0)
    Else
        GeoLine cu + ru * Cos(t0), cv + rv * Sin(t0)
    End If
    n = Int(Abs(sweep) / 1.5707963267949 - 0.000001) + 1
    d = sweep / n
    k = 4# / 3# * Tan(d / 4#)
    a = t0
    For i = 1 To n
        b = a + d
        GeoCubic cu + ru * (Cos(a) - k * Sin(a)), cv + rv * (Sin(a) + k * Cos(a)), _
                 cu + ru * (Cos(b) + k * Sin(b)), cv + rv * (Sin(b) - k * Cos(b)), _
                 cu + ru * Cos(b), cv + rv * Sin(b)
        a = b
    Next i
End Sub

Private Function Adj(ByVal s As Object, ByVal i As Long, ByVal defVal As Double) As Double
    Adj = defVal
    On Error Resume Next
    If s.Adjustments.Count >= i Then Adj = s.Adjustments.Item(i)
End Function

' Common preset shapes, from the DrawingML preset definitions (adjustment values as PowerPoint reports them).
Private Function VxPreset(ByVal s As Object, ByVal ast As Long, ByVal w As Double, ByVal h As Double) As Boolean
    Dim ss As Double, a As Double, a2 As Double, x1 As Double, x2 As Double, y1 As Double, y2 As Double, d As Double
    ss = MinS(CSng(w), CSng(h))
    VxPreset = True
    Select Case ast
        Case 1, 61                                   ' rectangle, flowchart process
            GeoRect 0, 0, w, h, 0, 0
        Case 5                                       ' rounded rectangle
            a = Adj(s, 1, 0.16667)
            If a < 0 Then a = 0
            If a > 0.5 Then a = 0.5
            GeoRect 0, 0, w, h, a * ss, a * ss
        Case 62                                      ' flowchart alternate process
            GeoRect 0, 0, w, h, ss / 6, ss / 6
        Case 69                                      ' flowchart terminator
            GeoRect 0, 0, w, h, w * 3475 / 21600, h / 2
        Case 9, 73                                   ' oval, flowchart connector
            GeoEllipse w / 2, h / 2, w / 2, h / 2
        Case 4, 63                                   ' diamond, flowchart decision
            GeoPoly w / 2, 0, w, h / 2, w / 2, h, 0, h / 2
        Case 7                                       ' isosceles triangle
            GeoPoly Adj(s, 1, 0.5) * w, 0, w, h, 0, h
        Case 8                                       ' right triangle
            GeoPoly 0, h, 0, 0, w, h
        Case 2                                       ' parallelogram
            x1 = Adj(s, 1, 0.25) * ss
            GeoPoly 0, h, x1, 0, w, 0, w - x1, h
        Case 64                                      ' flowchart data
            GeoPoly w / 5, 0, w, 0, w * 4 / 5, h, 0, h
        Case 3                                       ' trapezoid
            x1 = Adj(s, 1, 0.25) * ss
            GeoPoly 0, h, x1, 0, w - x1, 0, w, h
        Case 10                                      ' hexagon
            x1 = Adj(s, 1, 0.25) * ss
            GeoPoly 0, h / 2, x1, 0, w - x1, 0, w, h / 2, w - x1, h, x1, h
        Case 6                                       ' octagon
            x1 = Adj(s, 1, 0.29289) * ss
            GeoPoly 0, x1, x1, 0, w - x1, 0, w, x1, w, h - x1, w - x1, h, x1, h, 0, h - x1
        Case 12                                      ' regular pentagon
            GeoPoly w / 2, 0, w, h * 0.38196, w * 0.80902, h, w * 0.19098, h, 0, h * 0.38196
        Case 51                                      ' pentagon (home plate)
            x1 = w - Adj(s, 1, 0.5) * ss
            GeoPoly 0, 0, x1, 0, w, h / 2, x1, h, 0, h
        Case 52                                      ' chevron
            x1 = Adj(s, 1, 0.5) * ss
            GeoPoly 0, 0, w - x1, 0, w, h / 2, w - x1, h, 0, h, x1, h / 2
        Case 33, 34                                  ' right / left arrow
            d = h * Adj(s, 1, 0.5) / 2
            x1 = w - Adj(s, 2, 0.5) * ss
            If ast = 33 Then
                GeoPoly 0, h / 2 - d, x1, h / 2 - d, x1, 0, w, h / 2, x1, h, x1, h / 2 + d, 0, h / 2 + d
            Else
                x1 = w - x1
                GeoPoly w, h / 2 - d, x1, h / 2 - d, x1, 0, 0, h / 2, x1, h, x1, h / 2 + d, w, h / 2 + d
            End If
        Case 35, 36                                  ' up / down arrow
            d = w * Adj(s, 1, 0.5) / 2
            y1 = Adj(s, 2, 0.5) * ss
            If ast = 35 Then
                GeoPoly w / 2, 0, w, y1, w / 2 + d, y1, w / 2 + d, h, w / 2 - d, h, w / 2 - d, y1, 0, y1
            Else
                y1 = h - y1
                GeoPoly w / 2, h, w, y1, w / 2 + d, y1, w / 2 + d, 0, w / 2 - d, 0, w / 2 - d, y1, 0, y1
            End If
        Case 37                                      ' left-right arrow
            d = h * Adj(s, 1, 0.5) / 2
            x1 = Adj(s, 2, 0.5) * ss
            x2 = w - x1
            GeoPoly 0, h / 2, x1, 0, x1, h / 2 - d, x2, h / 2 - d, x2, 0, w, h / 2, x2, h, x2, h / 2 + d, x1, h / 2 + d, x1, h
        Case 38                                      ' up-down arrow
            d = w * Adj(s, 1, 0.5) / 2
            y1 = Adj(s, 2, 0.5) * ss
            y2 = h - y1
            GeoPoly w / 2, 0, w, y1, w / 2 + d, y1, w / 2 + d, y2, w, y2, w / 2, h, 0, y2, w / 2 - d, y2, w / 2 - d, y1, 0, y1
        Case 65                                      ' flowchart predefined process
            GeoRect 0, 0, w, h, 0, 0
            GeoBegin True, False
            GeoMove w / 8, 0: GeoLine w / 8, h
            GeoEnd False
            GeoBegin True, False
            GeoMove w * 7 / 8, 0: GeoLine w * 7 / 8, h
            GeoEnd False
        Case 67                                      ' flowchart document
            GeoBegin False, False
            GeoMove 0, 0
            GeoLine w, 0
            GeoLine w, h * 17322 / 21600
            GeoCubic w * 10800 / 21600, h * 17322 / 21600, w * 10800 / 21600, h * 23922 / 21600, 0, h * 20172 / 21600
            GeoEnd True
        Case 29, 30                                  ' left / right bracket (outline only)
            y1 = Adj(s, 1, 0.08333) * ss
            gGeoOpenOnly = True
            GeoBegin True, False
            If ast = 29 Then
                GeoMove w, h
                GeoArc w, h - y1, w, y1, 1.5707963267949, 1.5707963267949, False
                GeoLine 0, y1
                GeoArc w, y1, w, y1, 3.14159265358979, 1.5707963267949, False
            Else
                GeoMove 0, 0
                GeoArc 0, y1, w, y1, -1.5707963267949, 1.5707963267949, False
                GeoLine w, h - y1
                GeoArc 0, h - y1, w, y1, 0, 1.5707963267949, False
            End If
            GeoEnd False
        Case 31, 32                                  ' left / right brace (outline only)
            y1 = Adj(s, 1, 0.08333) * ss
            y2 = Adj(s, 2, 0.5) * h
            gGeoOpenOnly = True
            GeoBegin True, False
            If ast = 31 Then
                GeoMove w, h
                GeoArc w, h - y1, w / 2, y1, 1.5707963267949, 1.5707963267949, False
                GeoLine w / 2, y2 + y1
                GeoArc 0, y2 + y1, w / 2, y1, 0, -1.5707963267949, False
                GeoArc 0, y2 - y1, w / 2, y1, 1.5707963267949, -1.5707963267949, False
                GeoLine w / 2, y1
                GeoArc w, y1, w / 2, y1, 3.14159265358979, 1.5707963267949, False
            Else
                GeoMove 0, 0
                GeoArc 0, y1, w / 2, y1, -1.5707963267949, 1.5707963267949, False
                GeoLine w / 2, y2 - y1
                GeoArc w, y2 - y1, w / 2, y1, 3.14159265358979, -1.5707963267949, False
                GeoArc w, y2 + y1, w / 2, y1, -1.5707963267949, -1.5707963267949, False
                GeoLine w / 2, h - y1
                GeoArc 0, h - y1, w / 2, y1, 0, 1.5707963267949, False
            End If
            GeoEnd False
        Case Else
            VxPreset = False
            Exit Function
    End Select
    Select Case ast
        Case 1, 5, 61, 62, 2, 3, 64, 65, 67
        Case Else
            gGeoTextBox = True                      ' text area is inset: place text where PowerPoint draws it
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
Private Function VxNodes(ByVal s As Object, ByVal l As Double, ByVal t As Double, ByVal useFill As Boolean) As Boolean
    Dim n As Long, i As Long, p As Variant, q As Variant, r As Variant
    Dim ux As Double, v As Double, su As Double, sv As Double, filled As Boolean, started As Boolean
    On Error GoTo Nope
    n = s.Nodes.Count
    If n < 2 Then Exit Function
    filled = False
    If useFill Then filled = (s.Fill.Visible = MSO_TRUE)
    i = 1
    Do While i <= n
        p = s.Nodes.Item(i).Points
        ux = p(1, 1) - l
        v = p(1, 2) - t
        If Not started Then
            GeoBegin False, False
            GeoMove ux, v
            su = ux: sv = v
            started = True
            i = i + 1
        ElseIf s.Nodes.Item(i).SegmentType = 1 And i + 2 <= n Then
            q = s.Nodes.Item(i + 1).Points
            r = s.Nodes.Item(i + 2).Points
            GeoCubic ux, v, q(1, 1) - l, q(1, 2) - t, r(1, 1) - l, r(1, 2) - t
            ux = r(1, 1) - l
            v = r(1, 2) - t
            i = i + 3
        Else
            GeoLine ux, v
            i = i + 1
        End If
        If started And i <= n And gGeoRow > 2 And Abs(ux - su) < 0.01 And Abs(v - sv) < 0.01 Then
            GeoEnd True                              ' sub-path closed: the next node starts a new one
            started = False
        End If
    Loop
    If started Then GeoEnd filled Or (Abs(ux - su) < 0.01 And Abs(v - sv) < 0.01)
    If Not filled Then gGeoOpenOnly = True
    VxNodes = (Len(gGeo) > 0)
    Exit Function
Nope:
    VxNodes = False
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
Private Function XpsToSvg(ByVal pres As Object, ByVal shp As Object, ByVal dst As String) As Boolean
#If Mac Then
    XpsToSvg = False
#Else
    Dim tmp As Object, c As Object, rng As Object, x As String, sc As String, cmd As String
    Dim l As Double, t As Double, w As Double, h As Double, i As Long, wasSaved As Long, ok As Boolean
    wasSaved = pres.Saved
    x = TempFileX("xps")
    On Error GoTo Failed
    sc = HelperScriptPath()
    If Len(sc) = 0 Then Exit Function
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
    pres.PrintOptions.Ranges.ClearAll
    Set rng = pres.PrintOptions.Ranges.Add(tmp.SlideIndex, tmp.SlideIndex)
    ' Path, XPS, print intent, no frame, handout order, slides, hidden slides, range, slide range
    pres.ExportAsFixedFormat x, 1, 2, MSO_FALSE, 1, 1, MSO_TRUE, rng, 4
    ProgressStep U("8F6C63624E3A0020005300560047")
    cmd = "& ([scriptblock]::Create([IO.File]::ReadAllText('" & PsQ(sc) & "',[Text.Encoding]::UTF8))) " & _
          "-Xps '" & PsQ(x) & "' -Out '" & PsQ(dst) & "' -Frame '" & NumXml(l) & "," & NumXml(t) & "," & _
          NumXml(w) & "," & NumXml(h) & "' -Slide '" & NumXml(pres.PageSetup.SlideWidth) & "," & _
          NumXml(pres.PageSetup.SlideHeight) & "'"
    DeleteIfExists dst
    ok = (RunPowerShell(cmd) = 0)
    If ok Then ok = Fso().FileExists(dst)
Cleanup:
    On Error Resume Next
    If Not tmp Is Nothing Then tmp.Delete
    pres.PrintOptions.Ranges.ClearAll
    DeleteTempX x
    pres.Saved = wasSaved
    On Error GoTo 0
    XpsToSvg = ok
    Exit Function
Failed:
    ok = False
    Resume Cleanup
#End If
End Function

#If Mac Then
#Else
' Writes the XPS -> SVG converter (embedded in this module) to the temp folder once per version.
Private Function HelperScriptPath() As String
    Dim p As String
    On Error GoTo Failed
    p = JoinPath(TempDirX(), "xps2svg-" & VS_VERSION & ".ps1")
    If Not Fso().FileExists(p) Then WriteBytesX p, Utf8Bytes(HelperScript())
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
    If s Is Nothing Then Err.Raise vbObjectError + 513, , U("624090095BF98C6165E06CD57EC454085BFC51FA3002")
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
    Err.Raise vbObjectError + 512, , U("65E06CD5901A8FC7526A8D34677F590D52366240900956FE5F62FF08526A8D34677F53EF80FD88AB51764ED67A0B5E8F53607528FF09FF0C8BF791CD8BD53002")
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
    Err.Raise Err.Number, , Err.Description & vbCrLf & U("FF080050006F0077006500720050006F0069006E007400200066006F00720020004D00610063002097008981002000310036002E003800320020621666F49AD87248672C624D80FD5BFC51FA56FE5F62FF09")
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
    FileManagerName = U("8BBF8FBE")
#Else
    FileManagerName = U("8D446E907BA174065668")
#End If
End Function

' Save dialog. fmt is updated when the user picks another type / types another extension.
Private Function AskSavePath(ByVal initDir As String, ByVal baseName As String, ByRef fmt As Long) As String
    Dim p As String, k As Long
#If Mac Then
    p = MacScript2("chooseSaveFile", U("5BFC51FA4E3A0020") & UCase$(FmtExt(fmt)) & MAC_SEP & _
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
    dlgTitle = U("77E253700020005600650063005300740061006D00700020002D00205BFC51FA90094E2D56FE5F62")
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
    folder = PickFolderX(initDir, U("900962E95BFC51FA65874EF65939"))
    If Len(folder) = 0 Then Exit Function
    nm = InputBox(U("65874EF6540DFF084E0D542B62695C55540DFF09FF1A"), AppTitle(), baseName)
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
    ok = (RunPowerShell(cmd) = 0)
    DeleteIfExists src
#End If
    If ok Then ok = FileExistsX(dst)
    GzipX = ok
End Function

' Zips a folder into an OPC package; [Content_Types].xml goes first, entry names use "/".
Private Sub ZipFolderX(ByVal dirPath As String, ByVal zipPath As String)
#If Mac Then
    If MacScript2("zipFolder", dirPath & MAC_SEP & zipPath) <> "ok" Then
        Err.Raise vbObjectError + 516, , U("6253530500200056005300440058002059318D253002")
    End If
#Else
    Dim cmd As String
    DeleteIfExists zipPath
    cmd = "Add-Type -AssemblyName System.IO.Compression;Add-Type -AssemblyName System.IO.Compression.FileSystem;" & _
          "$src='" & PsQ(dirPath) & "';$dst='" & PsQ(zipPath) & "';" & _
          "$z=[IO.Compression.ZipFile]::Open($dst,'Create');" & _
          "$fs=@(Get-ChildItem -LiteralPath $src -Recurse -File -Force|Sort-Object @{e={if($_.Name -eq '[Content_Types].xml'){0}else{1}}},FullName);" & _
          "foreach($f in $fs){$n=$f.FullName.Substring($src.Length).TrimStart('\').Replace('\','/');" & _
          "[void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($z,$f.FullName,$n)};$z.Dispose()"
    If RunPowerShell(cmd) <> 0 Or Not Fso().FileExists(zipPath) Then
        Err.Raise vbObjectError + 516, , U("6253530500200056005300440058002059318D25FF08970089817CFB7EDF81EA5E267684002000570069006E0064006F0077007300200050006F007700650072005300680065006C006CFF093002")
    End If
#End If
End Sub

' EMF -> SVG with Inkscape or LibreOffice, when one of them is installed.
Private Function ConvertEmfToSvg(ByVal src As String, ByVal dst As String) As Boolean
#If Mac Then
    If MacScript2("inkscapeSvg", src & MAC_SEP & dst) = "ok" Then
        ConvertEmfToSvg = True
    ElseIf MacScript2("libreofficeSvg", src & MAC_SEP & dst) = "ok" Then
        ConvertEmfToSvg = True
    End If
#Else
    Dim c As Variant, exe As String, outDir As String, produced As String
    For Each c In Array(Environ$("ProgramFiles") & "\Inkscape\bin\inkscape.exe", _
                        Environ$("ProgramW6432") & "\Inkscape\bin\inkscape.exe", _
                        Environ$("ProgramFiles(x86)") & "\Inkscape\bin\inkscape.exe", _
                        Environ$("LOCALAPPDATA") & "\Programs\Inkscape\bin\inkscape.exe")
        If Fso().FileExists(CStr(c)) Then exe = CStr(c): Exit For
    Next
    If Len(exe) > 0 Then
        RunWait """" & exe & """ """ & src & """ --export-type=svg ""--export-filename=" & dst & """"
        If Fso().FileExists(dst) Then ConvertEmfToSvg = True: Exit Function
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
        RunWait """" & exe & """ -env:UserInstallation=" & FileUrlX(JoinPath(TempDirX(), "lo_profile")) & _
                " --headless --norestore --convert-to svg --outdir """ & outDir & """ """ & src & """"
        produced = JoinPath(outDir, BaseNameOf(src) & ".svg")
        If Fso().FileExists(produced) Then
            DeleteIfExists dst
            Fso().MoveFile produced, dst
            ConvertEmfToSvg = True
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
    Err.Raise vbObjectError + 520, , U("627E4E0D52300020006D00610063004F005300208F8552A9811A672C0020") & MAC_SCRIPT & U("FF0C8BF791CD65B08FD0884C0020006D00610063004F005300205B8988C556683002")
End Function

Private Sub MacMove(ByVal src As String, ByVal dst As String)
    If MacScript2("moveFile", src & MAC_SEP & dst) <> "ok" Then
        Err.Raise vbObjectError + 521, , U("65E06CD5628A65874EF651995165FF1A") & dst
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

Private Function RunWait(ByVal cmd As String) As Long
    On Error Resume Next
    RunWait = CreateObject("WScript.Shell").Run(cmd, 0, True)
    If Err.Number <> 0 Then RunWait = -1
End Function

' Runs a PowerShell snippet hidden; returns its exit code (0 = success).
Private Function RunPowerShell(ByVal script As String) As Long
    RunPowerShell = RunWait("powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden " & _
                            "-EncodedCommand " & B64Utf16("$ErrorActionPreference='Stop';" & script))
End Function

Private Function PsQ(ByVal s As String) As String
    ' Escape for a PowerShell single-quoted string (PowerShell also treats curly quotes as quotes).
    s = Replace(s, "'", "''")
    s = Replace(s, ChrW(&H2018), ChrW(&H2018) & ChrW(&H2018))
    s = Replace(s, ChrW(&H2019), ChrW(&H2019) & ChrW(&H2019))
    PsQ = s
End Function

Private Function B64Utf16(ByVal s As String) As String
    Dim b() As Byte
    b = s                                           ' VBA strings are UTF-16LE
    With CreateObject("MSXML2.DOMDocument").createElement("b64")
        .DataType = "bin.base64"
        .nodeTypedValue = b
        B64Utf16 = Replace(Replace(.Text, vbLf, ""), vbCr, "")
    End With
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
' =====================================================================
Private Sub ProgressBegin(ByVal title As String, ByVal total As Long)
    ProgressEnd
    gPgTotal = total
    If gPgTotal < 1 Then gPgTotal = 1
    gPgPos = 0
    gPgActive = True
    gPgTitle = title
#If Mac Then
    gPgMacPid = ""
    On Error Resume Next
    gPgMacPid = AppleScriptTask(MAC_SCRIPT, "progressShow", title & U("FF0C8BF77A0D50192026"))
#Else
    ProgressCreate title
#End If
End Sub

' Advances one step and shows what is being done.
Private Sub ProgressStep(ByVal detail As String)
    If Not gPgActive Then Exit Sub
    gPgPos = gPgPos + 1
    If gPgPos > gPgTotal Then gPgTotal = gPgPos + 1
#If Mac Then
#Else
    ProgressUpdate detail
#End If
End Sub

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

' True when the user pressed Esc in the meantime (Windows only).
Private Function ProgressCancelled() As Boolean
#If Mac Then
#Else
    If gPgActive Then ProgressCancelled = ((GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0)
#End If
End Function

' MsgBox that hides the progress window while it is shown.
Private Function AskUser(ByVal prompt As String, ByVal buttons As VbMsgBoxStyle) As VbMsgBoxResult
#If Mac Then
    AskUser = MsgBox(prompt, buttons, AppTitle())
#Else
    If gPgWnd <> 0 Then ShowWindow gPgWnd, SW_HIDE
    AskUser = MsgBox(prompt, buttons, AppTitle())
    If gPgWnd <> 0 And gPgActive Then ShowWindow gPgWnd, SW_SHOWNOACTIVATE
#End If
End Function

#If Mac Then
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
    w = CLng(440 * k)
    h = CLng(150 * k)
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
    face = "Microsoft YaHei UI"
    gPgFontB = CreateFontW(-CLng(10.5 * dpi / 72), 0, 0, 0, 700, 0, 0, 0, 1, 0, 0, 5, 0, StrPtr(face))
    gPgFont = CreateFontW(-CLng(9 * dpi / 72), 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 5, 0, StrPtr(face))
    cls = "Static"
    gPgMsg = CreateWindowExW(0, StrPtr(cls), StrPtr(title), WS_CHILD Or WS_VISIBLE Or SS_NOPREFIX, _
                             CLng(20 * k), CLng(16 * k), CLng(392 * k), CLng(24 * k), gPgWnd, 0, 0, 0)
    cls = "msctls_progress32"
    gPgBar = CreateWindowExW(0, StrPtr(cls), 0, WS_CHILD Or WS_VISIBLE, _
                             CLng(20 * k), CLng(46 * k), CLng(392 * k), CLng(16 * k), gPgWnd, 0, 0, 0)
    cls = "Static"
    cap = U("51C659074E2D2026")
    gPgSub = CreateWindowExW(0, StrPtr(cls), StrPtr(cap), WS_CHILD Or WS_VISIBLE Or SS_NOPREFIX, _
                             CLng(20 * k), CLng(70 * k), CLng(392 * k), CLng(22 * k), gPgWnd, 0, 0, 0)
    SendMessageW gPgMsg, WM_SETFONT, gPgFontB, 1
    SendMessageW gPgSub, WM_SETFONT, gPgFont, 1
    If gPgBar <> 0 Then SendMessageW gPgBar, PBM_SETRANGE32, 0, gPgTotal
    ShowWindow gPgWnd, SW_SHOWNOACTIVATE
    UpdateWindow gPgWnd
    DoEvents
Done:
End Sub

Private Sub ProgressUpdate(ByVal detail As String)
    On Error Resume Next
    If gPgWnd = 0 Then Exit Sub
    If IsWindow(gPgWnd) = 0 Then gPgWnd = 0: Exit Sub
    SetWindowTextW gPgSub, StrPtr(detail & U("FF0863090020004500730063002053EF53D66D88FF09"))
    If gPgBar <> 0 Then
        SendMessageW gPgBar, PBM_SETRANGE32, 0, gPgTotal
        SendMessageW gPgBar, PBM_SETPOS, gPgPos, 0
    End If
    UpdateWindow gPgWnd
    DoEvents
End Sub
#End If


' =====================================================================
'  Feedback, About and donations
' =====================================================================
Private Sub DoFeedback()
    Dim ans As VbMsgBoxResult
    ans = MsgBox(U("8BF7900962E953CD998865B95F0FFF1A") & vbCrLf & vbCrLf & _
                 U("3010662F301153D1900190AE4EF67ED94F5C8005FF08") & VS_EMAIL & U("FF09") & vbCrLf & _
                 U("30105426301157280020004700690074004800750062002063D04EA4002000490073007300750065FF0897008981002000470069007400480075006200208D2653F7FF0C95EE98984F1A516C5F00663E793AFF09") & vbCrLf & _
                 U("301053D66D8830118FD456DE"), vbQuestion + vbYesNoCancel, AppTitle() & U("0020002D002053CD9988"))
    Select Case ans
        Case vbYes: FeedbackMail
        Case vbNo: FeedbackIssue
    End Select
End Sub

Private Sub FeedbackMail()
    Dim copied As Boolean, url As String
    copied = SetClipboardText(VS_EMAIL)
    If MsgBox(IIf(copied, U("4F5C800590AE7BB15DF2590D52365230526A8D34677FFF1A"), U("4F5C800590AE7BB1FF1A")) & VS_EMAIL & vbCrLf & vbCrLf & _
              U("4F6053EF4EE557284EFB4F5590AE7BB1FF087F51987590AE7BB1300100510051002090AE7BB13001004F00750074006C006F006F006B00207B49FF094E2D65B05EFA90AE4EF630017C988D3465364EF64EBA540E53D190013002") & vbCrLf & vbCrLf & _
              U("898173B0572862535F00753581114E0A76849ED88BA490AE4EF67A0B5E8F5417FF1F") & vbCrLf & U("FF085982679C6CA167098BBE7F6E90AE4EF67A0B5E8FFF0C53EF4EE5900962E9300C5426300DFF0C76F463A575287F51987590AE7BB153D19001FF09"), _
              vbQuestion + vbYesNo + vbDefaultButton2, AppTitle() & U("0020002D002090AE4EF653CD9988")) <> vbYes Then Exit Sub
    url = "mailto:" & VS_EMAIL & "?subject=" & UrlEncode(U("77E253700020005600650063005300740061006D00700020") & VS_VERSION & U("002053CD9988")) & _
          "&body=" & UrlEncode(vbCrLf & vbCrLf & "----" & vbCrLf & EnvironmentInfo())
    If Not OpenUrlX(url) Then
        MsgBox U("65E06CD562535F0090AE4EF67A0B5E8FFF0C8BF776F463A553D190AE4EF65230FF1A") & VS_EMAIL, vbInformation, AppTitle()
    End If
End Sub

Private Sub FeedbackIssue()
    Dim url As String, body As String
    body = U("002A002A95EE989863CF8FF00020002F00205EFA8BAEFF1A002A002A") & vbLf & vbLf & vbLf & U("002A002A590D73B06B659AA4FF1A002A002A") & vbLf & "1. " & vbLf & vbLf & _
           U("002A002A8FD0884C73AF5883FF1A002A002A") & vbLf & "```" & vbLf & Replace(EnvironmentInfo(), vbCrLf, vbLf) & vbLf & "```" & vbLf
    url = VS_REPO & "/issues/new?title=" & UrlEncode(U("005B53CD9988005D0020")) & "&body=" & UrlEncode(body)
    If Not OpenUrlX(url) Then
        MsgBox U("65E06CD562535F006D4F89C85668FF0C8BF7624B52A88BBF95EEFF1A") & vbCrLf & VS_REPO & "/issues", vbInformation, AppTitle()
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
    MsgBox U("77E253700020005600650063005300740061006D0070002000200076") & VS_VERSION & vbCrLf & _
           U("90094E2D53735370002020142014002056FE5F624E00952E5BFC51FA00200045004D00460020002F00200053005600470020002F00200050004400460020002F002000560053004400580020002F00200050004E004700207B49683C5F0F") & vbCrLf & vbCrLf & _
           U("4F5C8005FF1A") & VS_AUTHOR & vbCrLf & _
           U("90AE7BB1FF1A") & VS_EMAIL & vbCrLf & _
           U("4E3B9875FF1A") & VS_REPO & vbCrLf & _
           U("5F006E90534F8BAEFF1A004D004900540020004C006900630065006E00730065") & vbCrLf & _
           "Copyright (c) " & VS_YEAR & " " & VS_AUTHOR & vbCrLf & vbCrLf & _
           U("77E25370627F8BFA6C384E455F006E9030016C384E45514D8D39FF1A6CA1670965368D39529F80FDFF0C6CA167095E7F544AFF0C4E0D80547F51653696C64EFB4F554FE1606FFF0C516890E86E907801572800200047006900740048007500620020516C5F003002") & vbCrLf & vbCrLf & _
           U("5982679C5B834E3A4F60828277014E8665F695F4FF0C6B228FCE70B951FB529F80FD533A300C51734E8E300D7EC47684300C8D5E8D4F4F5C8005300DFF0C75285FAE4FE16216652F4ED85B9D626B78018BF74F5C8005559D676F549655613002") & _
           U("8D5E8D4F5B8C516881EA613FFF0C4E0D5F7154CD4EFB4F55529F80FDFF1B4F607684652F63014F1A75284E8E63017EED7EF462A4548C65398FDB77E253703002"), _
           vbInformation, U("51734E8E0020") & AppTitle()
End Sub

Private Sub DoDonateThanks(ByVal which As String)
    MsgBox U("8C228C224F60613F610F652F630177E25370FF01") & vbCrLf & vbCrLf & _
           U("8BF77528") & which & U("626B63CF300C8D5E8D4F4F5C8005300D4E2D5BF95E9476844E8C7EF47801FF0C91D1989D968F610F3002") & vbCrLf & _
           U("8D5E8D4F5B8C516881EA613FFF0C77E25370768462406709529F80FD90FD6C384E45514D8D3930015F006E903002") & vbCrLf & vbCrLf & _
           U("5982679C65B94FBFFF0C4E5F6B228FCE5728002000470069007400480075006200207ED9987976EE70B94E004E2A00200053007400610072FF0C6216628A5B8363A883507ED9970089817684540C5B66548C540C4E8B3002"), _
           vbInformation, AppTitle() & U("0020002D00208D5E8D4F")
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
        label = U("6DE182720020") & CLng(f * 100) & "%"
    Else
        l = l * (1 + f)
        label = U("6DF182720020") & CLng(-f * 100) & "%"
    End If
    ThemeTint = HslToRgb(h, s, l)
End Function

Private Function GalColor(ByVal index As Long, Optional ByRef label As String) As Long
    Dim names As Variant, std As Variant, stdNames As Variant, rc As Variant, col As Long, row As Long, t As String
    names = Array(U("767D8272FF0C80CC666F00200031"), U("9ED18272FF0C65875B5700200031"), U("80CC666F00200032"), U("65875B5700200032"), U("4E2A6027827200200031"), U("4E2A6027827200200032"), U("4E2A6027827200200033"), U("4E2A6027827200200034"), U("4E2A6027827200200035"), U("4E2A6027827200200036"))
    std = Array(&HC0&, &HFF&, &HC0FF&, &HFFFF&, &H50D092, &H50B000, &HF0B000, &HC07000, &H602000, &HA03070)
    stdNames = Array(U("6DF17EA2"), U("7EA28272"), U("6A598272"), U("9EC48272"), U("6D457EFF"), U("7EFF8272"), U("6D4584DD"), U("84DD8272"), U("6DF184DD"), U("7D2B8272"))
    If index < 60 Then
        col = index Mod 10
        row = index \ 10
        If row = 0 Then
            GalColor = ThemeBase(col)
            label = names(col)
        Else
            GalColor = ThemeTint(ThemeBase(col), row, t)
            label = names(col) & U("FF0C") & t
        End If
    ElseIf index < 70 Then
        GalColor = std(index - 60)
        label = U("680751C68272FF1A") & stdNames(index - 60)
    Else
        rc = RecentColors()
        If index - 70 <= UBound(rc) Then
            GalColor = rc(index - 70)
            label = U("67008FD14F7F75287684989C8272FF1A") & ColorText(rc(index - 70))
        End If
    End If
    If Len(label) > 0 And index < 70 Then label = label & U("FF08") & ColorText(GalColor) & U("FF09")
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

' "More Colors" and the eyedropper run PowerPoint's own tools on a temporary off-slide shape
' and read its fill colour afterwards. Falls back to the system colour dialog.
Private Sub DoMoreColors()
    Dim c As Long, trans As Double
    c = BgColorValue(BG_CUSTOM)
    trans = BgAlpha() / 100#
    Select Case PickWithPowerPoint("ObjectFillMoreColorsDialog", False, c, trans)
        Case 1
            SetBgColor c
            PutSet "BgAlpha", CStr(CLng(trans * 100))
            RefreshRibbon
        Case -1
            DoPickColor                                  ' PowerPoint's dialog unavailable: system dialog
    End Select
End Sub

Private Sub DoEyedropper()
    Dim c As Long, trans As Double
    c = BgColorValue(BG_CUSTOM)
    trans = 0
    Select Case PickWithPowerPoint("EyedropperFill", True, c, trans)
        Case 1
            SetBgColor c
        Case -1
            MsgBox U("53D682725668970089815728666E901A89C656FE4E2D62535F004E005F205E7B706F72473002") & vbCrLf & _
                   U("8BF7520763625230666E901A89C656FE540E518D70B9300C53D682725668300DFF0C7136540E57285E7B706F72474E0A70B951FB898162FE53D67684989C82723002"), vbInformation, AppTitle()
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
    If waitForPick Then ProgressBegin U("53D682725668FF1A57285E7B706F72474E0A70B951FB898162FE53D67684989C8272"), 1
    Application.CommandBars.ExecuteMso idMso
    If waitForPick Then
        started = Timer
        Do
            DoEvents
            If tmp.Fill.ForeColor.RGB <> c0 Then Exit Do
            If ProgressCancelled() Then Exit Do
            If Timer - started > 60 Or Timer < started Then Exit Do
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

' Solid-colour swatch for the gallery (Windows: a small BMP file loaded as a picture).
Private Function SwatchPicture(ByVal c As Long, ByVal size As Long) As Object
#If Mac Then
#Else
    Dim p As String, b() As Byte, i As Long, x As Long, y As Long, rowBytes As Long, n As Long
    Dim r As Long, g As Long, bl As Long, edge As Boolean, raw As String
    On Error GoTo Done
    p = JoinPath(TempDirX(), "swatch_" & Mid$(HexRGB(c), 2) & "_" & size & ".bmp")
    If Not Fso().FileExists(p) Then
        rowBytes = ((size * 3 + 3) \ 4) * 4
        n = 54 + rowBytes * size
        ReDim b(0 To n - 1)
        b(0) = 66: b(1) = 77                                     ' "BM"
        PutLong b, 2, n
        PutLong b, 10, 54
        PutLong b, 14, 40
        PutLong b, 18, size
        PutLong b, 22, size
        b(26) = 1: b(28) = 24
        PutLong b, 34, rowBytes * size
        r = c And &HFF&: g = (c \ &H100&) And &HFF&: bl = (c \ &H10000) And &HFF&
        For y = 0 To size - 1
            For x = 0 To size - 1
                i = 54 + y * rowBytes + x * 3
                edge = (x = 0 Or y = 0 Or x = size - 1 Or y = size - 1)
                If edge Then
                    b(i) = 160: b(i + 1) = 160: b(i + 2) = 160
                Else
                    b(i) = bl: b(i + 1) = g: b(i + 2) = r
                End If
            Next x
        Next y
        raw = b
        WriteBytesX p, raw
    End If
    Set SwatchPicture = LoadPicture(p)
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
    If idx > 0 Then s = s & U("005F7B2C") & idx & U("9875")
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
        FolderSettingText = U("FF08672A8BBE7F6EFF1A4F7F75286F14793A65877A3F6240572865874EF65939FF1B672A4FDD5B58621657284E917AEF65F64F7F7528684C9762FF09")
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
        Case FMT_EMZ: FilterLabel = U("0045004D005A0020538B7F297684589E5F3A578B56FE514365874EF600200028002A002E0065006D007A0029")
        Case FMT_SVG: FilterLabel = U("005300560047002053EF7F29653E77E291CF56FE5F6200200028002A002E0073007600670029")
        Case FMT_WMF: FilterLabel = U("0057004D0046002000570069006E0064006F00770073002056FE514365874EF600200028002A002E0077006D00660029")
        Case FMT_PNG: FilterLabel = U("0050004E004700209AD852068FA873874F4D56FE00200028002A002E0070006E00670029")
        Case FMT_PDF: FilterLabel = U("005000440046002077E291CF6587686300200028002A002E0070006400660029")
        Case FMT_VSDX: FilterLabel = U("005600530044005800200056006900730069006F00207ED856FE00200028002A002E00760073006400780029")
        Case Else: FilterLabel = U("0045004D00460020589E5F3A578B56FE514365874EF600200028002A002E0065006D00660029")
    End Select
End Function

Private Function BgName(ByVal bg As Long) As String
    Select Case bg
        Case BG_WHITE: BgName = U("767D8272")
        Case BG_BEIGE: BgName = U("7C738272")
        Case BG_GRID: BgName = U("7F51683CFF0895F48DDD0020") & NumText(GridStep()) & U("002078C5FF09")
        Case BG_CUSTOM: BgName = U("81EA5B9A4E49989C82720020") & ColorSettingText()
        Case Else: BgName = U("900F660E")
    End Select
End Function

Private Function SettingsSummary() As String
    Dim s As String, bg As Long
    s = U("63095F53524D8BBE7F6E5BFC51FA90094E2D56FE5F62") & vbCrLf & vbCrLf
    s = s & U("9ED88BA4683C5F0FFF1A") & UCase$(FmtExt(GetIdx("Format", FMT_LAST))) & vbCrLf
    If GetIdx("Mode", 1) = 1 Then
        s = s & U("591A4E2A56FE5F62FF1A6BCF4E2A56FE5F624E004E2A65874EF6") & vbCrLf
    Else
        s = s & U("591A4E2A56FE5F62FF1A54085E764E3A4E004E2A65874EF6") & vbCrLf
    End If
    If GetIdx("Naming", 1) = 1 Then
        s = s & U("4FDD5B5865B95F0FFF1A81EA52A84FDD5B5852309ED88BA465874EF65939") & vbCrLf
    Else
        s = s & U("4FDD5B5865B95F0FFF1A6BCF6B21900962E94FDD5B584F4D7F6E") & vbCrLf
    End If
    s = s & U("0050004E0047002052068FA87387FF1A") & DpiValue() & " DPI" & vbCrLf
    bg = BgTypeIdx()
    s = s & U("80CC666FFF1A") & BgName(bg)
    If bg <> BG_NONE Then s = s & U("FF0C900F660E5EA60020") & BgAlpha() & "%"
    s = s & vbCrLf & U("8FB98DDDFF1A") & NumText(GetMargin()) & U("002078C5")
    If GetFlag("VsdxNative", "1") Then
        s = s & vbCrLf & U("0056005300440058FF1A0056006900730069006F0020539F751F5F6272B6FF0865875B5753EF7F168F91FF09")
    Else
        s = s & vbCrLf & U("0056005300440058FF1A5D4C516500200045004D0046002077E291CF56FE7247")
    End If
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
    msg = U("5DF25BFC51FA0020") & files.Count & U("00204E2A65874EF6FF1A") & vbCrLf & vbCrLf
    For i = 1 To files.Count
        If i > 8 Then msg = msg & U("20262026") & vbCrLf: Exit For
        msg = msg & CStr(files(i)) & vbCrLf
    Next i
    MsgBox msg, vbInformation, AppTitle()
End Sub

Private Sub RefreshRibbon()
    On Error Resume Next
    If Not gRibbon Is Nothing Then gRibbon.Invalidate
End Sub

Private Function AppTitle() As String
    AppTitle = U("77E253700020005600650063005300740061006D0070")
End Function

' Decodes UTF-16 code units written as 4-digit hex groups (used by the ASCII-only build).
Private Function U(ByVal h As String) As String
    Dim i As Long, s As String
    For i = 1 To Len(h) Step 4
        s = s & ChrW$(CLng("&H" & Mid$(h, i, 4)))
    Next i
    U = s
End Function


' >>> GENERATED: VSDX TEMPLATE PARTS AND HELPER SCRIPT (generated by build/build_bas.py - do not edit)
Private Function VsdxPart(ByVal partName As String) As String
    Dim s As String
    Select Case partName
        Case "content_types.xml"
            s = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types""><Default Extension=""emf"" ContentType=""image/x-emf"
            s = s & """/><Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/><Default Extension=""xml"" ContentType=""application/xml""/><Override PartName=""/vi"
            s = s & "sio/document.xml"" ContentType=""application/vnd.ms-visio.drawing.main+xml""/><Override PartName=""/visio/pages/pages.xml"" ContentType=""application/vnd.ms-visio.pages+xml""/><Override P"
            s = s & "artName=""/visio/pages/page1.xml"" ContentType=""application/vnd.ms-visio.page+xml""/><Override PartName=""/visio/windows.xml"" ContentType=""application/vnd.ms-visio.windows+xml""/><Overr"
            s = s & "ide PartName=""/docProps/core.xml"" ContentType=""application/vnd.openxmlformats-package.core-properties+xml""/><Override PartName=""/docProps/app.xml"" ContentType=""application/vnd.open"
            s = s & "xmlformats-officedocument.extended-properties+xml""/></Types>"
        Case "root.rels"
            s = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://sche"
            s = s & "mas.microsoft.com/visio/2010/relationships/document"" Target=""visio/document.xml""/><Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/metadat"
            s = s & "a/core-properties"" Target=""docProps/core.xml""/><Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"" Target=""docPro"
            s = s & "ps/app.xml""/></Relationships>"
        Case "app.xml"
            s = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"" xmlns:vt=""http://schemas.openxm"
            s = s & "lformats.org/officeDocument/2006/docPropsVTypes""><Application>VecStamp</Application><AppVersion>{VERSION}</AppVersion><HeadingPairs><vt:vector size=""2"" baseType=""variant""><vt:varia"
            s = s & "nt><vt:lpstr>Pages</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size=""1"" baseType=""lpstr""><vt:lpstr>Page-1</"
            s = s & "vt:lpstr></vt:vector></TitlesOfParts></Properties>"
        Case "core.xml"
            s = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" xmlns:dc=""http://purl.o"
            s = s & "rg/dc/elements/1.1/"" xmlns:dcterms=""http://purl.org/dc/terms/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""><dc:title>{TITLE}</dc:title><dc:creator>VecStamp</dc:creator><d"
            s = s & "cterms:created xsi:type=""dcterms:W3CDTF"">{CREATED}</dcterms:created></cp:coreProperties>"
        Case "document.xml"
            s = "<?xml version='1.0' encoding='utf-8' ?><VisioDocument xmlns='http://schemas.microsoft.com/office/visio/2012/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/rel"
            s = s & "ationships' xml:space='preserve'><DocumentSettings TopPage='0' DefaultTextStyle='0' DefaultLineStyle='0' DefaultFillStyle='0' DefaultGuideStyle='0'><GlueSettings>9</GlueSettings><S"
            s = s & "napSettings>65847</SnapSettings><SnapExtensions>34</SnapExtensions><SnapAngles/><DynamicGridEnabled>1</DynamicGridEnabled><ProtectStyles>0</ProtectStyles><ProtectShapes>0</ProtectS"
            s = s & "hapes><ProtectMasters>0</ProtectMasters><ProtectBkgnds>0</ProtectBkgnds></DocumentSettings><Colors><ColorEntry IX='24' RGB='#7F7F7F'/><ColorEntry IX='25' RGB='#FFFFFF'/></Colors><F"
            s = s & "aceNames><FaceName NameU='Calibri' UnicodeRanges='-469750017 -1073732485 9 0' CharSets='536871423 0' Panose='2 15 5 2 2 2 4 3 2 4' Flags='325'/></FaceNames><StyleSheets><StyleSheet"
            s = s & " ID='0' NameU='No Style' IsCustomNameU='1' Name='No Style' IsCustomName='1'><Cell N='EnableLineProps' V='1'/><Cell N='EnableFillProps' V='1'/><Cell N='EnableTextProps' V='1'/><Cell"
            s = s & " N='HideForApply' V='0'/><Cell N='LineWeight' V='0.01041666666666667'/><Cell N='LineColor' V='0'/><Cell N='LinePattern' V='1'/><Cell N='Rounding' V='0'/><Cell N='EndArrowSize' V='2"
            s = s & "'/><Cell N='BeginArrow' V='0'/><Cell N='EndArrow' V='0'/><Cell N='LineCap' V='0'/><Cell N='BeginArrowSize' V='2'/><Cell N='LineColorTrans' V='0'/><Cell N='CompoundType' V='0'/><Cel"
            s = s & "l N='FillForegnd' V='1'/><Cell N='FillBkgnd' V='0'/><Cell N='FillPattern' V='1'/><Cell N='ShdwForegnd' V='0'/><Cell N='ShdwPattern' V='0'/><Cell N='FillForegndTrans' V='0'/><Cell N"
            s = s & "='FillBkgndTrans' V='0'/><Cell N='ShdwForegndTrans' V='0'/><Cell N='ShapeShdwType' V='0'/><Cell N='ShapeShdwOffsetX' V='0'/><Cell N='ShapeShdwOffsetY' V='0'/><Cell N='ShapeShdwObli"
            s = s & "queAngle' V='0'/><Cell N='ShapeShdwScaleFactor' V='1'/><Cell N='ShapeShdwBlur' V='0'/><Cell N='ShapeShdwShow' V='0'/><Cell N='LeftMargin' V='0'/><Cell N='RightMargin' V='0'/><Cell "
            s = s & "N='TopMargin' V='0'/><Cell N='BottomMargin' V='0'/><Cell N='VerticalAlign' V='1'/><Cell N='TextBkgnd' V='0'/><Cell N='DefaultTabStop' V='0.5905511811023622'/><Cell N='TextDirection"
            s = s & "' V='0'/><Cell N='TextBkgndTrans' V='0'/><Cell N='LockWidth' V='0'/><Cell N='LockHeight' V='0'/><Cell N='LockMoveX' V='0'/><Cell N='LockMoveY' V='0'/><Cell N='LockAspect' V='0'/><C"
            s = s & "ell N='LockDelete' V='0'/><Cell N='LockBegin' V='0'/><Cell N='LockEnd' V='0'/><Cell N='LockRotate' V='0'/><Cell N='LockCrop' V='0'/><Cell N='LockVtxEdit' V='0'/><Cell N='LockTextEd"
            s = s & "it' V='0'/><Cell N='LockFormat' V='0'/><Cell N='LockGroup' V='0'/><Cell N='LockCalcWH' V='0'/><Cell N='LockSelect' V='0'/><Cell N='LockCustProp' V='0'/><Cell N='LockFromGroupFormat"
            s = s & "' V='0'/><Cell N='LockThemeColors' V='0'/><Cell N='LockThemeEffects' V='0'/><Cell N='LockThemeConnectors' V='0'/><Cell N='LockThemeFonts' V='0'/><Cell N='LockThemeIndex' V='0'/><Ce"
            s = s & "ll N='LockReplace' V='0'/><Cell N='LockVariation' V='0'/><Cell N='NoObjHandles' V='0'/><Cell N='NonPrinting' V='0'/><Cell N='NoCtlHandles' V='0'/><Cell N='NoAlignBox' V='0'/><Cell "
            s = s & "N='UpdateAlignBox' V='0'/><Cell N='HideText' V='0'/><Cell N='DynFeedback' V='0'/><Cell N='GlueType' V='0'/><Cell N='WalkPreference' V='0'/><Cell N='BegTrigger' V='0' F='No Formula'"
            s = s & "/><Cell N='EndTrigger' V='0' F='No Formula'/><Cell N='ObjType' V='0'/><Cell N='Comment' V=''/><Cell N='IsDropSource' V='0'/><Cell N='NoLiveDynamics' V='0'/><Cell N='LocalizeMerge' "
            s = s & "V='0'/><Cell N='NoProofing' V='0'/><Cell N='Calendar' V='0'/><Cell N='LangID' V='en-US'/><Cell N='ShapeKeywords' V=''/><Cell N='DropOnPageScale' V='1'/><Cell N='TheData' V='0' F='N"
            s = s & "o Formula'/><Cell N='TheText' V='0' F='No Formula'/><Cell N='EventDblClick' V='0' F='No Formula'/><Cell N='EventXFMod' V='0' F='No Formula'/><Cell N='EventDrop' V='0' F='No Formula"
            s = s & "'/><Cell N='EventMultiDrop' V='0' F='No Formula'/><Cell N='HelpTopic' V=''/><Cell N='Copyright' V=''/><Cell N='LayerMember' V=''/><Cell N='XRulerDensity' V='32'/><Cell N='YRulerDen"
            s = s & "sity' V='32'/><Cell N='XRulerOrigin' V='0'/><Cell N='YRulerOrigin' V='0'/><Cell N='XGridDensity' V='8'/><Cell N='YGridDensity' V='8'/><Cell N='XGridSpacing' V='0'/><Cell N='YGridSp"
            s = s & "acing' V='0'/><Cell N='XGridOrigin' V='0'/><Cell N='YGridOrigin' V='0'/><Cell N='Gamma' V='1'/><Cell N='Contrast' V='0.5'/><Cell N='Brightness' V='0.5'/><Cell N='Sharpen' V='0'/><C"
            s = s & "ell N='Blur' V='0'/><Cell N='Denoise' V='0'/><Cell N='Transparency' V='0'/><Cell N='SelectMode' V='1'/><Cell N='DisplayMode' V='2'/><Cell N='IsDropTarget' V='0'/><Cell N='IsSnapTar"
            s = s & "get' V='1'/><Cell N='IsTextEditTarget' V='1'/><Cell N='DontMoveChildren' V='0'/><Cell N='ShapePermeableX' V='0'/><Cell N='ShapePermeableY' V='0'/><Cell N='ShapePermeablePlace' V='0"
            s = s & "'/><Cell N='Relationships' V='0'/><Cell N='ShapeFixedCode' V='0'/><Cell N='ShapePlowCode' V='0'/><Cell N='ShapeRouteStyle' V='0'/><Cell N='ShapePlaceStyle' V='0'/><Cell N='ConFixed"
            s = s & "Code' V='0'/><Cell N='ConLineJumpCode' V='0'/><Cell N='ConLineJumpStyle' V='0'/><Cell N='ConLineJumpDirX' V='0'/><Cell N='ConLineJumpDirY' V='0'/><Cell N='ShapePlaceFlip' V='0'/><C"
            s = s & "ell N='ConLineRouteExt' V='0'/><Cell N='ShapeSplit' V='0'/><Cell N='ShapeSplittable' V='0'/><Cell N='DisplayLevel' V='0'/><Cell N='ResizePage' V='0'/><Cell N='EnableGrid' V='0'/><C"
            s = s & "ell N='DynamicsOff' V='0'/><Cell N='CtrlAsInput' V='0'/><Cell N='AvoidPageBreaks' V='0'/><Cell N='PlaceStyle' V='0'/><Cell N='RouteStyle' V='0'/><Cell N='PlaceDepth' V='0'/><Cell N"
            s = s & "='PlowCode' V='0'/><Cell N='LineJumpCode' V='1'/><Cell N='LineJumpStyle' V='0'/><Cell N='PageLineJumpDirX' V='0'/><Cell N='PageLineJumpDirY' V='0'/><Cell N='LineToNodeX' V='0.09842"
            s = s & "519685039369'/><Cell N='LineToNodeY' V='0.09842519685039369'/><Cell N='BlockSizeX' V='0.1968503937007874'/><Cell N='BlockSizeY' V='0.1968503937007874'/><Cell N='AvenueSizeX' V='0.2"
            s = s & "952755905511811'/><Cell N='AvenueSizeY' V='0.2952755905511811'/><Cell N='LineToLineX' V='0.09842519685039369'/><Cell N='LineToLineY' V='0.09842519685039369'/><Cell N='LineJumpFacto"
            s = s & "rX' V='0.66666666666667'/><Cell N='LineJumpFactorY' V='0.66666666666667'/><Cell N='LineAdjustFrom' V='0'/><Cell N='LineAdjustTo' V='0'/><Cell N='PlaceFlip' V='0'/><Cell N='LineRout"
            s = s & "eExt' V='0'/><Cell N='PageShapeSplit' V='0'/><Cell N='PageLeftMargin' V='0.25'/><Cell N='PageRightMargin' V='0.25'/><Cell N='PageTopMargin' V='0.25'/><Cell N='PageBottomMargin' V='"
            s = s & "0.25'/><Cell N='ScaleX' V='1'/><Cell N='ScaleY' V='1'/><Cell N='PagesX' V='1'/><Cell N='PagesY' V='1'/><Cell N='CenterX' V='0'/><Cell N='CenterY' V='0'/><Cell N='OnPage' V='0'/><Ce"
            s = s & "ll N='PrintGrid' V='0'/><Cell N='PrintPageOrientation' V='1'/><Cell N='PaperKind' V='9'/><Cell N='PaperSource' V='7'/><Cell N='QuickStyleLineColor' V='100'/><Cell N='QuickStyleFill"
            s = s & "Color' V='100'/><Cell N='QuickStyleShadowColor' V='100'/><Cell N='QuickStyleFontColor' V='100'/><Cell N='QuickStyleLineMatrix' V='100'/><Cell N='QuickStyleFillMatrix' V='100'/><Cel"
            s = s & "l N='QuickStyleEffectsMatrix' V='100'/><Cell N='QuickStyleFontMatrix' V='100'/><Cell N='QuickStyleType' V='0'/><Cell N='QuickStyleVariation' V='0'/><Cell N='LineGradientDir' V='0'/"
            s = s & "><Cell N='LineGradientAngle' V='1.5707963267949'/><Cell N='FillGradientDir' V='0'/><Cell N='FillGradientAngle' V='1.5707963267949'/><Cell N='LineGradientEnabled' V='0'/><Cell N='Fi"
            s = s & "llGradientEnabled' V='0'/><Cell N='RotateGradientWithShape' V='1'/><Cell N='UseGroupGradient' V='0'/><Cell N='BevelTopType' V='0'/><Cell N='BevelTopWidth' V='0'/><Cell N='BevelTopH"
            s = s & "eight' V='0'/><Cell N='BevelBottomType' V='0'/><Cell N='BevelBottomWidth' V='0'/><Cell N='BevelBottomHeight' V='0'/><Cell N='BevelDepthColor' V='1'/><Cell N='BevelDepthSize' V='0'/"
            s = s & "><Cell N='BevelContourColor' V='0'/><Cell N='BevelContourSize' V='0'/><Cell N='BevelMaterialType' V='0'/><Cell N='BevelLightingType' V='0'/><Cell N='BevelLightingAngle' V='0'/><Cel"
            s = s & "l N='RotationXAngle' V='0'/><Cell N='RotationYAngle' V='0'/><Cell N='RotationZAngle' V='0'/><Cell N='RotationType' V='0'/><Cell N='Perspective' V='0'/><Cell N='DistanceFromGround' "
            s = s & "V='0'/><Cell N='KeepTextFlat' V='0'/><Cell N='ReflectionTrans' V='0'/><Cell N='ReflectionSize' V='0'/><Cell N='ReflectionDist' V='0'/><Cell N='ReflectionBlur' V='0'/><Cell N='GlowC"
            s = s & "olor' V='1'/><Cell N='GlowColorTrans' V='0'/><Cell N='GlowSize' V='0'/><Cell N='SoftEdgesSize' V='0'/><Cell N='SketchSeed' V='0'/><Cell N='SketchEnabled' V='0'/><Cell N='SketchAmou"
            s = s & "nt' V='5'/><Cell N='SketchLineWeight' V='0.04166666666666666' U='PT'/><Cell N='SketchLineChange' V='0.14'/><Cell N='SketchFillChange' V='0.1'/><Cell N='ColorSchemeIndex' V='0'/><Ce"
            s = s & "ll N='EffectSchemeIndex' V='0'/><Cell N='ConnectorSchemeIndex' V='0'/><Cell N='FontSchemeIndex' V='0'/><Cell N='ThemeIndex' V='0'/><Cell N='VariationColorIndex' V='0'/><Cell N='Var"
            s = s & "iationStyleIndex' V='0'/><Cell N='EmbellishmentIndex' V='0'/><Cell N='ReplaceLockShapeData' V='0'/><Cell N='ReplaceLockText' V='0'/><Cell N='ReplaceLockFormat' V='0'/><Cell N='Repl"
            s = s & "aceCopyCells' V='0' U='BOOL' F='No Formula'/><Cell N='PageWidth' V='0' F='No Formula'/><Cell N='PageHeight' V='0' F='No Formula'/><Cell N='ShdwOffsetX' V='0' F='No Formula'/><Cell "
            s = s & "N='ShdwOffsetY' V='0' F='No Formula'/><Cell N='PageScale' V='0' U='MM' F='No Formula'/><Cell N='DrawingScale' V='0' U='MM' F='No Formula'/><Cell N='DrawingSizeType' V='0' F='No For"
            s = s & "mula'/><Cell N='DrawingScaleType' V='0' F='No Formula'/><Cell N='InhibitSnap' V='0' F='No Formula'/><Cell N='PageLockReplace' V='0' U='BOOL' F='No Formula'/><Cell N='PageLockDuplic"
            s = s & "ate' V='0' U='BOOL' F='No Formula'/><Cell N='UIVisibility' V='0' F='No Formula'/><Cell N='ShdwType' V='0' F='No Formula'/><Cell N='ShdwObliqueAngle' V='0' F='No Formula'/><Cell N='"
            s = s & "ShdwScaleFactor' V='0' F='No Formula'/><Cell N='DrawingResizeType' V='0' F='No Formula'/><Section N='Character'><Row IX='0'><Cell N='Font' V='Calibri'/><Cell N='Color' V='0'/><Cell"
            s = s & " N='Style' V='0'/><Cell N='Case' V='0'/><Cell N='Pos' V='0'/><Cell N='FontScale' V='1'/><Cell N='Size' V='0.1666666666666667'/><Cell N='DblUnderline' V='0'/><Cell N='Overline' V='0"
            s = s & "'/><Cell N='Strikethru' V='0'/><Cell N='DoubleStrikethrough' V='0'/><Cell N='Letterspace' V='0'/><Cell N='ColorTrans' V='0'/><Cell N='AsianFont' V='0'/><Cell N='ComplexScriptFont' "
            s = s & "V='0'/><Cell N='ComplexScriptSize' V='-1'/><Cell N='LangID' V='en-US'/></Row></Section><Section N='Paragraph'><Row IX='0'><Cell N='IndFirst' V='0'/><Cell N='IndLeft' V='0'/><Cell N"
            s = s & "='IndRight' V='0'/><Cell N='SpLine' V='-1.2'/><Cell N='SpBefore' V='0'/><Cell N='SpAfter' V='0'/><Cell N='HorzAlign' V='1'/><Cell N='Bullet' V='0'/><Cell N='BulletStr' V=''/><Cell "
            s = s & "N='BulletFont' V='0'/><Cell N='BulletFontSize' V='-1'/><Cell N='TextPosAfterBullet' V='0'/><Cell N='Flags' V='0'/></Row></Section><Section N='Tabs'><Row IX='0'/></Section><Section "
            s = s & "N='LineGradient'><Row IX='0'><Cell N='GradientStopColor' V='1'/><Cell N='GradientStopColorTrans' V='0'/><Cell N='GradientStopPosition' V='0'/></Row></Section><Section N='FillGradie"
            s = s & "nt'><Row IX='0'><Cell N='GradientStopColor' V='1'/><Cell N='GradientStopColorTrans' V='0'/><Cell N='GradientStopPosition' V='0'/></Row></Section></StyleSheet></StyleSheets><Documen"
            s = s & "tSheet NameU='TheDoc' IsCustomNameU='1' Name='TheDoc' IsCustomName='1' LineStyle='0' FillStyle='0' TextStyle='0'><Cell N='OutputFormat' V='0'/><Cell N='LockPreview' V='0'/><Cell N="
            s = s & "'AddMarkup' V='0'/><Cell N='ViewMarkup' V='0'/><Cell N='DocLockReplace' V='0' U='BOOL'/><Cell N='NoCoauth' V='0' U='BOOL'/><Cell N='DocLockDuplicatePage' V='0' U='BOOL'/><Cell N='P"
            s = s & "reviewQuality' V='0'/><Cell N='PreviewScope' V='0'/><Cell N='DocLangID' V='en-US'/><Section N='User'><Row N='msvNoAutoConnect'><Cell N='Value' V='1'/><Cell N='Prompt' V='' F='No Fo"
            s = s & "rmula'/></Row></Section></DocumentSheet></VisioDocument>"
        Case "document.xml.rels"
            s = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://sche"
            s = s & "mas.microsoft.com/visio/2010/relationships/pages"" Target=""pages/pages.xml""/><Relationship Id=""rId2"" Type=""http://schemas.microsoft.com/visio/2010/relationships/windows"" Target=""win"
            s = s & "dows.xml""/></Relationships>"
        Case "windows.xml"
            s = "<?xml version='1.0' encoding='utf-8' ?><Windows ClientWidth='1280' ClientHeight='720' xmlns='http://schemas.microsoft.com/office/visio/2012/main' xmlns:r='http://schemas.openxmlfor"
            s = s & "mats.org/officeDocument/2006/relationships' xml:space='preserve'><Window ID='0' WindowType='Drawing' WindowState='1073741824' ContainerType='Page' Page='0' ViewScale='-1' ViewCente"
            s = s & "rX='{CX}' ViewCenterY='{CY}'><ShowRulers>1</ShowRulers><ShowGrid>1</ShowGrid><ShowPageBreaks>0</ShowPageBreaks><ShowGuides>1</ShowGuides><ShowConnectionPoints>1</ShowConnectionPoin"
            s = s & "ts><GlueSettings>9</GlueSettings><SnapSettings>65847</SnapSettings><SnapExtensions>34</SnapExtensions><SnapAngles/><DynamicGridEnabled>1</DynamicGridEnabled><TabSplitterPos>0.5</Ta"
            s = s & "bSplitterPos></Window></Windows>"
        Case "pages.xml"
            s = "<?xml version='1.0' encoding='utf-8' ?><Pages xmlns='http://schemas.microsoft.com/office/visio/2012/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationshi"
            s = s & "ps' xml:space='preserve'><Page ID='0' NameU='Page-1' Name='Page-1' ViewScale='-1' ViewCenterX='{CX}' ViewCenterY='{CY}'><PageSheet LineStyle='0' FillStyle='0' TextStyle='0'><Cell N"
            s = s & "='PageWidth' V='{PW}'/><Cell N='PageHeight' V='{PH}'/><Cell N='ShdwOffsetX' V='0.125'/><Cell N='ShdwOffsetY' V='-0.125'/><Cell N='PageScale' V='1' U='IN'/><Cell N='DrawingScale' V="
            s = s & "'1' U='IN'/><Cell N='DrawingSizeType' V='3'/><Cell N='DrawingScaleType' V='0'/><Cell N='InhibitSnap' V='0'/><Cell N='UIVisibility' V='0'/><Cell N='ShdwType' V='0'/><Cell N='ShdwObl"
            s = s & "iqueAngle' V='0'/><Cell N='ShdwScaleFactor' V='1'/><Cell N='DrawingResizeType' V='1'/><Cell N='PageShapeSplit' V='1'/></PageSheet><Rel r:id='rId1'/></Page></Pages>"
        Case "pages.xml.rels"
            s = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://sche"
            s = s & "mas.microsoft.com/visio/2010/relationships/page"" Target=""page1.xml""/></Relationships>"
        Case "page1.xml"
            s = "<?xml version='1.0' encoding='utf-8' ?><PageContents xmlns='http://schemas.microsoft.com/office/visio/2012/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/rela"
            s = s & "tionships' xml:space='preserve'><Shapes>{SHAPES}</Shapes></PageContents>"
        Case "page1.xml.rels"
            s = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">{RELS}</Relationships>"
    End Select
    VsdxPart = s
End Function

Private Function VsdxMap() As Variant
    VsdxMap = Array("content_types.xml=[Content_Types].xml", _
                    "root.rels=_rels/.rels", _
                    "app.xml=docProps/app.xml", _
                    "core.xml=docProps/core.xml", _
                    "document.xml=visio/document.xml", _
                    "document.xml.rels=visio/_rels/document.xml.rels", _
                    "windows.xml=visio/windows.xml", _
                    "pages.xml=visio/pages/pages.xml", _
                    "pages.xml.rels=visio/pages/_rels/pages.xml.rels", _
                    "page1.xml=visio/pages/page1.xml", _
                    "page1.xml.rels=visio/pages/_rels/page1.xml.rels")
End Function

Private Function HelperScript1() As String
    Dim s As String
    s = s & "# VecStamp (Shi Yin) - XPS page -> SVG converter." & vbLf
    s = s & "#" & vbLf
    s = s & "# Used by the PowerPoint add-in on PowerPoint versions that cannot export SVG themselves" & vbLf
    s = s & "# (Windows, PowerPoint 2013 - 2021 / LTSC): the add-in prints the drawing to XPS with" & vbLf
    s = s & "# PowerPoint (""ExportAsFixedFormat"") and this script turns the first page into SVG," & vbLf
    s = s & "# cropped to the drawing. Paths, gradients, images, clips and transforms are converted" & vbLf
    s = s & "# 1:1; text stays text (font family, size, exact glyph positions) and the subset fonts" & vbLf
    s = s & "# embedded in the XPS are embedded in the SVG as well." & vbLf
    s = s & "#" & vbLf
    s = s & "# Windows PowerShell 5.1 and PowerShell 7 compatible; no modules needed." & vbLf
    s = s & "# The add-in embeds this file (build/build_bas.py), writes it to its temp folder and runs:" & vbLf
    s = s & "#   & ([scriptblock]::Create((Get-Content -Raw xps2svg.ps1))) -Xps in.xps -Out out.svg -Frame ""l,t,w,h"" -Slide ""w,h""" & vbLf
    s = s & "# Frame and Slide are in points (slide coordinates). Exit code 0 = success." & vbLf
    s = s & "#" & vbLf
    s = s & "# Copyright (c) 2026 vluckyzhang <vluckyzhang@gmail.com>. MIT License." & vbLf
    s = s & "" & vbLf
    s = s & "param(" & vbLf
    s = s & "    [Parameter(Mandatory = $true)][string]$Xps," & vbLf
    s = s & "    [Parameter(Mandatory = $true)][string]$Out," & vbLf
    s = s & "    [string]$Frame = ''," & vbLf
    s = s & "    [string]$Slide = ''," & vbLf
    s = s & "    [switch]$NoEmbedFonts" & vbLf
    s = s & ")" & vbLf
    s = s & "" & vbLf
    s = s & "$ErrorActionPreference = 'Stop'" & vbLf
    s = s & "Add-Type -AssemblyName System.IO.Compression" & vbLf
    s = s & "Add-Type -AssemblyName System.IO.Compression.FileSystem" & vbLf
    s = s & "$Inv = [Globalization.CultureInfo]::InvariantCulture" & vbLf
    s = s & "$KeyNs = 'http://schemas.microsoft.com/xps/2005/06/resourcedictionary-key'" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ helpers" & vbLf
    s = s & "function N([double]$v) {" & vbLf
    s = s & "    if ([Math]::Abs($v) -lt 0.0005) { return '0' }" & vbLf
    s = s & "    return $v.ToString('0.###', $Inv)" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Num([string]$s) {" & vbLf
    s = s & "    return [double]::Parse($s.Trim(), [Globalization.NumberStyles]::Float, $Inv)" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Esc([string]$s) {" & vbLf
    s = s & "    $sb = New-Object System.Text.StringBuilder" & vbLf
    s = s & "    foreach ($ch in $s.ToCharArray()) {" & vbLf
    s = s & "        $c = [int]$ch" & vbLf
    s = s & "        if ($c -eq 38) { [void]$sb.Append('&amp;') }" & vbLf
    s = s & "        elseif ($c -eq 60) { [void]$sb.Append('&lt;') }" & vbLf
    s = s & "        elseif ($c -eq 62) { [void]$sb.Append('&gt;') }" & vbLf
    s = s & "        elseif ($c -eq 34) { [void]$sb.Append('&quot;') }" & vbLf
    s = s & "        elseif ($c -lt 32 -and $c -ne 9 -and $c -ne 10 -and $c -ne 13) { }" & vbLf
    s = s & "        elseif ($c -eq 0xFFFE -or $c -eq 0xFFFF) { }" & vbLf
    s = s & "        else { [void]$sb.Append($ch) }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $sb.ToString()" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Numbers([string]$s) {" & vbLf
    s = s & "    $list = New-Object System.Collections.Generic.List[double]" & vbLf
    s = s & "    foreach ($m in [regex]::Matches($s, '[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?')) { $list.Add((Num $m.Value)) }" & vbLf
    s = s & "    return , $list.ToArray()" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# 2D affine matrices as [a, b, c, d, e, f] (x' = a x + c y + e, y' = b x + d y + f), XPS / SVG order" & vbLf
    s = s & "function MatMul($m, $n) {" & vbLf
    s = s & "    # result = apply m first, then n" & vbLf
    s = s & "    return @(($m[0] * $n[0] + $m[1] * $n[2]), ($m[0] * $n[1] + $m[1] * $n[3])," & vbLf
    s = s & "        ($m[2] * $n[0] + $m[3] * $n[2]), ($m[2] * $n[1] + $m[3] * $n[3])," & vbLf
    s = s & "        ($m[4] * $n[0] + $m[5] * $n[2] + $n[4]), ($m[4] * $n[1] + $m[5] * $n[3] + $n[5]))" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function MatText($m) {" & vbLf
    s = s & "    return 'matrix(' + (($m | ForEach-Object { N $_ }) -join ' ') + ')'" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function ParseMatrix([string]$s) {" & vbLf
    s = s & "    $v = Numbers $s" & vbLf
    s = s & "    if ($v.Count -ne 6) { return $null }" & vbLf
    s = s & "    return @($v[0], $v[1], $v[2], $v[3], $v[4], $v[5])" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ package" & vbLf
    s = s & "$Zip = [IO.Compression.ZipFile]::OpenRead($Xps)" & vbLf
    s = s & "$Parts = @{}" & vbLf
    s = s & "foreach ($e in $Zip.Entries) {" & vbLf
    s = s & "    $name = '/' + [Uri]::UnescapeDataString($e.FullName.Replace('\', '/'))" & vbLf
    s = s & "    $name = [regex]::Replace($name, '/\[\d+\]\.(last\.)?piece$', '')" & vbLf
    s = s & "    $Parts[$name.ToLowerInvariant()] = $e" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function ReadBytes([string]$part) {" & vbLf
    s = s & "    $e = $Parts[$part.ToLowerInvariant()]" & vbLf
    s = s & "    if ($null -eq $e) { return $null }" & vbLf
    s = s & "    $s = $e.Open()" & vbLf
    s = s & "    $ms = New-Object IO.MemoryStream" & vbLf
    s = s & "    $s.CopyTo($ms)" & vbLf
    s = s & "    $s.Dispose()" & vbLf
    s = s & "    return , $ms.ToArray()" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function ReadXml([string]$part) {" & vbLf
    s = s & "    $b = ReadBytes $part" & vbLf
    s = s & "    if ($null -eq $b) { return $null }" & vbLf
    s = s & "    $x = New-Object System.Xml.XmlDocument" & vbLf
    s = s & "    $x.PreserveWhitespace = $true" & vbLf
    s = s & "    $x.Load((New-Object IO.MemoryStream (, $b)))" & vbLf
    s = s & "    return $x" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Resolve([string]$base, [string]$uri) {" & vbLf
    s = s & "    $u = [Uri]::UnescapeDataString($uri.Trim())" & vbLf
    s = s & "    if ($u.StartsWith('/')) { return $u }" & vbLf
    s = s & "    $dir = $base.Substring(0, $base.LastIndexOf('/') + 1)" & vbLf
    s = s & "    $stack = New-Object System.Collections.ArrayList" & vbLf
    s = s & "    foreach ($seg in ($dir + $u).Split('/')) {" & vbLf
    s = s & "        if ($seg -eq '' -or $seg -eq '.') { continue }" & vbLf
    s = s & "        if ($seg -eq '..') { if ($stack.Count) { $stack.RemoveAt($stack.Count - 1) } ; continue }" & vbLf
    s = s & "        [void]$stack.Add($seg)" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return '/' + ($stack -join '/')" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function FirstPage {" & vbLf
    s = s & "    $rels = ReadXml '/_rels/.rels'" & vbLf
    s = s & "    if ($rels) {" & vbLf
    s = s & "        foreach ($r in $rels.DocumentElement.ChildNodes) {" & vbLf
    s = s & "            if ($r.LocalName -eq 'Relationship' -and $r.GetAttribute('Type') -match 'fixedrepresentation$') {" & vbLf
    s = s & "                $seqPart = Resolve '/' $r.GetAttribute('Target')" & vbLf
    s = s & "                $seq = ReadXml $seqPart" & vbLf
    s = s & "                if (-not $seq) { break }" & vbLf
    s = s & "                foreach ($d in $seq.DocumentElement.ChildNodes) {" & vbLf
    s = s & "                    if ($d.LocalName -ne 'DocumentReference') { continue }" & vbLf
    s = s & "                    $docPart = Resolve $seqPart $d.GetAttribute('Source')" & vbLf
    s = s & "                    $doc = ReadXml $docPart" & vbLf
    s = s & "                    if (-not $doc) { continue }" & vbLf
    s = s & "                    foreach ($p in $doc.DocumentElement.ChildNodes) {" & vbLf
    s = s & "                        if ($p.LocalName -eq 'PageContent') { return (Resolve $docPart $p.GetAttribute('Source')) }" & vbLf
    s = s & "                    }" & vbLf
    s = s & "                }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    $names = @($Parts.Keys | Where-Object { $_ -like '*.fpage' } | Sort-Object)" & vbLf
    s = s & "    if ($names.Count) { return $names[0] }" & vbLf
    s = s & "    throw 'no FixedPage in XPS'" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ output state" & vbLf
    s = s & "$Defs = New-Object System.Text.StringBuilder" & vbLf
    s = s & "$Body = New-Object System.Text.StringBuilder" & vbLf
    s = s & "$script:IdN = 0" & vbLf
    s = s & "$Fonts = @{}          # font part -> info hashtable" & vbLf
    s = s & "$Resources = @{}      # x:Key -> XmlElement" & vbLf
    s = s & "$script:Box = $null   # content bounding box in page units [x0, y0, x1, y1]" & vbLf
    s = s & "$script:PageW = 0.0" & vbLf
    s = s & "$script:PageH = 0.0" & vbLf
    s = s & "" & vbLf
    s = s & "function NewId([string]$prefix) {" & vbLf
    s = s & "    $script:IdN++" & vbLf
    s = s & "    return $prefix + $script:IdN" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function AddBox($m, [double[]]$pts, [double]$pad) {" & vbLf
    s = s & "    # pts: x0, y0, x1, y1, ... in local coordinates; m: matrix to page coordinates" & vbLf
    s = s & "    for ($i = 0; $i + 1 -lt $pts.Count; $i += 2) {" & vbLf
    s = s & "        $x = $m[0] * $pts[$i] + $m[2] * $pts[$i + 1] + $m[4]" & vbLf
    s = s & "        $y = $m[1] * $pts[$i] + $m[3] * $pts[$i + 1] + $m[5]" & vbLf
    s = s & "        if ($null -eq $script:Box) { $script:Box = @(($x - $pad), ($y - $pad), ($x + $pad), ($y + $pad)) }" & vbLf
    s = s & "        else {" & vbLf
    s = s & "            if ($x - $pad -lt $script:Box[0]) { $script:Box[0] = $x - $pad }" & vbLf
    s = s & "            if ($y - $pad -lt $script:Box[1]) { $script:Box[1] = $y - $pad }" & vbLf
    s = s & "            if ($x + $pad -gt $script:Box[2]) { $script:Box[2] = $x + $pad }" & vbLf
    s = s & "            if ($y + $pad -gt $script:Box[3]) { $script:Box[3] = $y + $pad }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ geometry" & vbLf
    s = s & "function FigureBounds([string]$d) {" & vbLf
    s = s & "    # absolute points of an abbreviated-syntax path (end and control points) for bounding boxes" & vbLf
    s = s & "    $pts = New-Object System.Collections.Generic.List[double]" & vbLf
    s = s & "    $toks = [regex]::Matches($d, '[A-Za-z]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?')" & vbLf
    s = s & "    $cmd = 'M'; $x = 0.0; $y = 0.0; $sx = 0.0; $sy = 0.0" & vbLf
    s = s & "    $nums = New-Object System.Collections.Generic.List[double]" & vbLf
    s = s & "    $flush = {" & vbLf
    s = s & "        param($c, $vals)" & vbLf
    s = s & "        $rel = ($c -cmatch '[a-z]')" & vbLf
    s = s & "        $C = $c.ToUpperInvariant()" & vbLf
    s = s & "        $k = 0" & vbLf
    s = s & "        $arity = @{ 'M' = 2; 'L' = 2; 'H' = 1; 'V' = 1; 'C' = 6; 'Q' = 4; 'S' = 4; 'T' = 2; 'A' = 7; 'Z' = 0 }[$C]" & vbLf
    s = s & "        if ($null -eq $arity -or $arity -eq 0) { return }" & vbLf
    s = s & "        while ($k + $arity -le $vals.Count) {" & vbLf
    s = s & "            switch ($C) {" & vbLf
    s = s & "                'H' { $script:px = $(if ($rel) { $script:px + $vals[$k] } else { $vals[$k] }) }" & vbLf
    s = s & "                'V' { $script:py = $(if ($rel) { $script:py + $vals[$k] } else { $vals[$k] }) }" & vbLf
    s = s & "                'A' {" & vbLf
    s = s & "                    $nx = $vals[$k + 5]; $ny = $vals[$k + 6]" & vbLf
    s = s & "                    if ($rel) { $nx += $script:px; $ny += $script:py }" & vbLf
    s = s & "                    $r = [Math]::Max([Math]::Abs($vals[$k]), [Math]::Abs($vals[$k + 1]))" & vbLf
    s = s & "                    $pts.Add($script:px - $r); $pts.Add($script:py - $r); $pts.Add($script:px + $r); $pts.Add($script:py + $r)" & vbLf
    s = s & "                    $script:px = $nx; $script:py = $ny" & vbLf
    s = s & "                }" & vbLf
    HelperScript1 = s
End Function

Private Function HelperScript2() As String
    Dim s As String
    s = s & "                default {" & vbLf
    s = s & "                    for ($j = 0; $j -lt $arity; $j += 2) {" & vbLf
    s = s & "                        $nx = $vals[$k + $j]; $ny = $vals[$k + $j + 1]" & vbLf
    s = s & "                        if ($rel) { $nx += $script:px; $ny += $script:py }" & vbLf
    s = s & "                        $pts.Add($nx); $pts.Add($ny)" & vbLf
    s = s & "                        if ($j + 2 -ge $arity) { $script:px = $nx; $script:py = $ny }" & vbLf
    s = s & "                    }" & vbLf
    s = s & "                }" & vbLf
    s = s & "            }" & vbLf
    s = s & "            $pts.Add($script:px); $pts.Add($script:py)" & vbLf
    s = s & "            $k += $arity" & vbLf
    s = s & "            if ($C -eq 'M') { $C = 'L' }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    $script:px = 0.0; $script:py = 0.0" & vbLf
    s = s & "    foreach ($t in $toks) {" & vbLf
    s = s & "        $v = $t.Value" & vbLf
    s = s & "        if ($v -cmatch '^[A-Za-z]$') {" & vbLf
    s = s & "            if ($v -eq 'F' -or $v -eq 'f') { $cmd = 'F'; $nums.Clear(); continue }" & vbLf
    s = s & "            if ($cmd -ne 'F') { & $flush $cmd $nums }" & vbLf
    s = s & "            $nums.Clear()" & vbLf
    s = s & "            $cmd = $v" & vbLf
    s = s & "        }" & vbLf
    s = s & "        else { $nums.Add((Num $v)) }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    if ($cmd -ne 'F') { & $flush $cmd $nums }" & vbLf
    s = s & "    return , $pts.ToArray()" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function GeometryFromElement($g) {" & vbLf
    s = s & "    # PathGeometry element -> @{ d = svg path; rule = 'nonzero'|'evenodd' }" & vbLf
    s = s & "    $rule = 'evenodd'" & vbLf
    s = s & "    if ($g.GetAttribute('FillRule') -eq 'NonZero') { $rule = 'nonzero' }" & vbLf
    s = s & "    $sb = New-Object System.Text.StringBuilder" & vbLf
    s = s & "    $fig = $g.GetAttribute('Figures')" & vbLf
    s = s & "    if ($fig) { [void]$sb.Append(($fig -replace '^\s*[Ff][01]\s*', '')) }" & vbLf
    s = s & "    foreach ($f in $g.ChildNodes) {" & vbLf
    s = s & "        if ($f.LocalName -ne 'PathFigure') { continue }" & vbLf
    s = s & "        [void]$sb.Append(' M ' + $f.GetAttribute('StartPoint'))" & vbLf
    s = s & "        foreach ($s in $f.ChildNodes) {" & vbLf
    s = s & "            switch ($s.LocalName) {" & vbLf
    s = s & "                'PolyLineSegment' { [void]$sb.Append(' L ' + $s.GetAttribute('Points')) }" & vbLf
    s = s & "                'PolyBezierSegment' { [void]$sb.Append(' C ' + $s.GetAttribute('Points')) }" & vbLf
    s = s & "                'PolyQuadraticBezierSegment' { [void]$sb.Append(' Q ' + $s.GetAttribute('Points')) }" & vbLf
    s = s & "                'ArcSegment' {" & vbLf
    s = s & "                    $sz = $s.GetAttribute('Size'); $rot = $s.GetAttribute('RotationAngle'); if (-not $rot) { $rot = '0' }" & vbLf
    s = s & "                    $large = $(if ($s.GetAttribute('IsLargeArc') -eq 'true') { '1' } else { '0' })" & vbLf
    s = s & "                    $sweep = $(if ($s.GetAttribute('SweepDirection') -eq 'Clockwise') { '1' } else { '0' })" & vbLf
    s = s & "                    [void]$sb.Append(' A ' + $sz + ' ' + $rot + ' ' + $large + ' ' + $sweep + ' ' + $s.GetAttribute('Point'))" & vbLf
    s = s & "                }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        if ($f.GetAttribute('IsClosed') -eq 'true') { [void]$sb.Append(' Z') }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return @{ d = $sb.ToString().Trim(); rule = $rule }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Geometry([string]$attr, $node, [string]$propName) {" & vbLf
    s = s & "    # Data / Clip attribute (abbreviated syntax or StaticResource) or <X.Data> child element" & vbLf
    s = s & "    if ($attr) {" & vbLf
    s = s & "        if ($attr -match '^\{StaticResource\s+(.+)\}$') {" & vbLf
    s = s & "            $r = $Resources[$matches[1].Trim()]" & vbLf
    s = s & "            if ($r) { return (GeometryFromElement $r) }" & vbLf
    s = s & "            return $null" & vbLf
    s = s & "        }" & vbLf
    s = s & "        $rule = 'evenodd'" & vbLf
    s = s & "        if ($attr -match '^\s*[Ff]1') { $rule = 'nonzero' }" & vbLf
    s = s & "        return @{ d = ($attr -replace '^\s*[Ff][01]\s*', ''); rule = $rule }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    foreach ($c in $node.ChildNodes) {" & vbLf
    s = s & "        if ($c.LocalName -eq $propName) {" & vbLf
    s = s & "            foreach ($g in $c.ChildNodes) { if ($g.LocalName -eq 'PathGeometry') { return (GeometryFromElement $g) } }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $null" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ colours and brushes" & vbLf
    s = s & "function Color([string]$s) {" & vbLf
    s = s & "    # -> @{ rgb = '#rrggbb'; a = 0..1 }" & vbLf
    s = s & "    $s = $s.Trim()" & vbLf
    s = s & "    if ($s -match '^#([0-9A-Fa-f]{8})$') {" & vbLf
    s = s & "        $h = $matches[1]" & vbLf
    s = s & "        return @{ rgb = '#' + $h.Substring(2, 6).ToLowerInvariant(); a = [Convert]::ToInt32($h.Substring(0, 2), 16) / 255.0 }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    if ($s -match '^#([0-9A-Fa-f]{6})$') { return @{ rgb = '#' + $matches[1].ToLowerInvariant(); a = 1.0 } }" & vbLf
    s = s & "    if ($s -match '^sc#') {" & vbLf
    s = s & "        $v = Numbers $s.Substring(3)" & vbLf
    s = s & "        if ($v.Count -eq 3) { $v = @(1.0) + $v }" & vbLf
    s = s & "        $c = @()" & vbLf
    s = s & "        for ($i = 1; $i -le 3; $i++) {" & vbLf
    s = s & "            $x = [Math]::Max(0.0, [Math]::Min(1.0, $v[$i]))" & vbLf
    s = s & "            $x = $(if ($x -le 0.0031308) { 12.92 * $x } else { 1.055 * [Math]::Pow($x, 1 / 2.4) - 0.055 })" & vbLf
    s = s & "            $c += [int][Math]::Round($x * 255)" & vbLf
    s = s & "        }" & vbLf
    s = s & "        return @{ rgb = ('#{0:x2}{1:x2}{2:x2}' -f $c[0], $c[1], $c[2]); a = [Math]::Max(0.0, [Math]::Min(1.0, $v[0])) }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return @{ rgb = '#000000'; a = 1.0 }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Stops($brush) {" & vbLf
    s = s & "    $sb = New-Object System.Text.StringBuilder" & vbLf
    s = s & "    foreach ($c in $brush.ChildNodes) {" & vbLf
    s = s & "        if ($c.LocalName -notlike '*.GradientStops') { continue }" & vbLf
    s = s & "        foreach ($s in $c.ChildNodes) {" & vbLf
    s = s & "            if ($s.LocalName -ne 'GradientStop') { continue }" & vbLf
    s = s & "            $col = Color $s.GetAttribute('Color')" & vbLf
    s = s & "            [void]$sb.Append('<stop offset=""' + (N (Num $s.GetAttribute('Offset'))) + '"" stop-color=""' + $col.rgb + '""')" & vbLf
    s = s & "            if ($col.a -lt 1) { [void]$sb.Append(' stop-opacity=""' + (N $col.a) + '""') }" & vbLf
    s = s & "            [void]$sb.Append('/>')" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $sb.ToString()" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function BrushTransform($brush) {" & vbLf
    s = s & "    $t = $brush.GetAttribute('Transform')" & vbLf
    s = s & "    if ($t -match '^\{StaticResource') { $t = '' }" & vbLf
    s = s & "    if ($t) { return (ParseMatrix $t) }" & vbLf
    s = s & "    foreach ($c in $brush.ChildNodes) {" & vbLf
    s = s & "        if ($c.LocalName -like '*.Transform') {" & vbLf
    s = s & "            foreach ($m in $c.ChildNodes) { if ($m.LocalName -eq 'MatrixTransform') { return (ParseMatrix $m.GetAttribute('Matrix')) } }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $null" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function ImageInfo([byte[]]$b) {" & vbLf
    s = s & "    # -> @{ mime; w; h; dpiX; dpiY } (pixel size and resolution)" & vbLf
    s = s & "    $info = @{ mime = 'image/png'; w = 0; h = 0; dpiX = 96.0; dpiY = 96.0 }" & vbLf
    s = s & "    if ($b.Length -gt 24 -and $b[0] -eq 0x89 -and $b[1] -eq 0x50) {" & vbLf
    s = s & "        $info.w = [int](U32 $b 16)" & vbLf
    s = s & "        $info.h = [int](U32 $b 20)" & vbLf
    s = s & "        $i = 8" & vbLf
    s = s & "        while ($i + 12 -le $b.Length) {" & vbLf
    s = s & "            $len = [int](U32 $b $i)" & vbLf
    s = s & "            $typ = [Text.Encoding]::ASCII.GetString($b, $i + 4, 4)" & vbLf
    s = s & "            if ($typ -eq 'pHYs' -and $b[$i + 16] -eq 1) {" & vbLf
    s = s & "                $px = U32 $b ($i + 8)" & vbLf
    s = s & "                $py = U32 $b ($i + 12)" & vbLf
    s = s & "                if ($px -gt 0) { $info.dpiX = $px * 0.0254 }" & vbLf
    s = s & "                if ($py -gt 0) { $info.dpiY = $py * 0.0254 }" & vbLf
    s = s & "            }" & vbLf
    s = s & "            if ($typ -eq 'IDAT' -or $typ -eq 'IEND' -or $len -lt 0) { break }" & vbLf
    s = s & "            $i += 12 + $len" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    elseif ($b.Length -gt 4 -and $b[0] -eq 0xFF -and $b[1] -eq 0xD8) {" & vbLf
    s = s & "        $info.mime = 'image/jpeg'" & vbLf
    s = s & "        $i = 2" & vbLf
    s = s & "        while ($i + 9 -lt $b.Length) {" & vbLf
    s = s & "            if ($b[$i] -ne 0xFF) { $i++; continue }" & vbLf
    s = s & "            $mk = $b[$i + 1]" & vbLf
    s = s & "            $len = ([int]$b[$i + 2] -shl 8) -bor $b[$i + 3]" & vbLf
    s = s & "            if ($mk -eq 0xE0 -and $b[$i + 4] -eq 0x4A -and $b[$i + 11] -eq 1) {" & vbLf
    s = s & "                $dx = ([int]$b[$i + 12] -shl 8) -bor $b[$i + 13]; $dy = ([int]$b[$i + 14] -shl 8) -bor $b[$i + 15]" & vbLf
    s = s & "                if ($dx -gt 0) { $info.dpiX = $dx; $info.dpiY = $dy }" & vbLf
    s = s & "            }" & vbLf
    s = s & "            if ($mk -ge 0xC0 -and $mk -le 0xCF -and $mk -ne 0xC4 -and $mk -ne 0xC8 -and $mk -ne 0xCC) {" & vbLf
    s = s & "                $info.h = ([int]$b[$i + 5] -shl 8) -bor $b[$i + 6]" & vbLf
    s = s & "                $info.w = ([int]$b[$i + 7] -shl 8) -bor $b[$i + 8]" & vbLf
    s = s & "                break" & vbLf
    s = s & "            }" & vbLf
    s = s & "            $i += 2 + $len" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    elseif ($b.Length -gt 4 -and (($b[0] -eq 0x49 -and $b[1] -eq 0x49) -or ($b[0] -eq 0x4D -and $b[1] -eq 0x4D))) {" & vbLf
    s = s & "        $info.mime = 'image/tiff'" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $info" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Rect4([string]$s) {" & vbLf
    s = s & "    $v = Numbers $s" & vbLf
    s = s & "    if ($v.Count -lt 4) { return @(0.0, 0.0, 1.0, 1.0) }" & vbLf
    s = s & "    return @($v[0], $v[1], $v[2], $v[3])" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Brush($brush, [string]$base) {" & vbLf
    s = s & "    # -> @{ paint = 'none' | '#rrggbb' | 'url(#id)'; a = opacity factor }" & vbLf
    s = s & "    if ($null -eq $brush) { return @{ paint = 'none'; a = 1.0 } }" & vbLf
    s = s & "    $op = 1.0" & vbLf
    s = s & "    if ($brush.GetAttribute('Opacity')) { $op = Num $brush.GetAttribute('Opacity') }" & vbLf
    s = s & "    switch ($brush.LocalName) {" & vbLf
    s = s & "        'SolidColorBrush' {" & vbLf
    s = s & "            $c = Color $brush.GetAttribute('Color')" & vbLf
    s = s & "            return @{ paint = $c.rgb; a = $c.a * $op }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        'LinearGradientBrush' {" & vbLf
    s = s & "            $id = NewId 'g'" & vbLf
    s = s & "            $s = Numbers $brush.GetAttribute('StartPoint'); $e = Numbers $brush.GetAttribute('EndPoint')" & vbLf
    s = s & "            $units = $(if ($brush.GetAttribute('MappingMode') -eq 'RelativeToBoundingBox') { 'objectBoundingBox' } else { 'userSpaceOnUse' })" & vbLf
    s = s & "            $spread = $brush.GetAttribute('SpreadMethod').ToLowerInvariant(); if (-not $spread) { $spread = 'pad' }" & vbLf
    s = s & "            [void]$Defs.Append('<linearGradient id=""' + $id + '"" gradientUnits=""' + $units + '"" spreadMethod=""' + $spread + '"" x1=""' + (N $s[0]) + '"" y1=""' + (N $s[1]) + '"" x2=""' + (N $e[0]) + '"" y2=""' + (N $e[1]) + '""')" & vbLf
    s = s & "            $m = BrushTransform $brush" & vbLf
    s = s & "            if ($m) { [void]$Defs.Append(' gradientTransform=""' + (MatText $m) + '""') }" & vbLf
    s = s & "            [void]$Defs.Append('>' + (Stops $brush) + '</linearGradient>')" & vbLf
    s = s & "            return @{ paint = 'url(#' + $id + ')'; a = $op }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        'RadialGradientBrush' {" & vbLf
    HelperScript2 = s
End Function

Private Function HelperScript3() As String
    Dim s As String
    s = s & "            $id = NewId 'g'" & vbLf
    s = s & "            $c = Numbers $brush.GetAttribute('Center'); $o = Numbers $brush.GetAttribute('GradientOrigin')" & vbLf
    s = s & "            $rx = Num $brush.GetAttribute('RadiusX'); $ry = Num $brush.GetAttribute('RadiusY')" & vbLf
    s = s & "            if ($rx -le 0) { $rx = 0.001 }" & vbLf
    s = s & "            if ($ry -le 0) { $ry = 0.001 }" & vbLf
    s = s & "            $units = $(if ($brush.GetAttribute('MappingMode') -eq 'RelativeToBoundingBox') { 'objectBoundingBox' } else { 'userSpaceOnUse' })" & vbLf
    s = s & "            $spread = $brush.GetAttribute('SpreadMethod').ToLowerInvariant(); if (-not $spread) { $spread = 'pad' }" & vbLf
    s = s & "            $k = $ry / $rx" & vbLf
    s = s & "            $scale = @(1.0, 0.0, 0.0, $k, 0.0, ($c[1] - $k * $c[1]))" & vbLf
    s = s & "            $m = BrushTransform $brush" & vbLf
    s = s & "            if ($m) { $scale = MatMul $scale $m }" & vbLf
    s = s & "            $fy = $c[1] + ($o[1] - $c[1]) / $k" & vbLf
    s = s & "            [void]$Defs.Append('<radialGradient id=""' + $id + '"" gradientUnits=""' + $units + '"" spreadMethod=""' + $spread + '"" cx=""' + (N $c[0]) + '"" cy=""' + (N $c[1]) + '"" r=""' + (N $rx) + '"" fx=""' + (N $o[0]) + '"" fy=""' + (N $fy) + '"" gradientTransform=""' + (MatText $scale) + '"">' + (Stops $brush) + '</radialGradient>')" & vbLf
    s = s & "            return @{ paint = 'url(#' + $id + ')'; a = $op }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        'ImageBrush' {" & vbLf
    s = s & "            $src = $brush.GetAttribute('ImageSource')" & vbLf
    s = s & "            if ($src -match '^\{') { return @{ paint = 'none'; a = 1.0 } }" & vbLf
    s = s & "            $part = Resolve $base $src" & vbLf
    s = s & "            $bytes = ReadBytes $part" & vbLf
    s = s & "            if ($null -eq $bytes) { return @{ paint = 'none'; a = 1.0 } }" & vbLf
    s = s & "            $info = ImageInfo $bytes" & vbLf
    s = s & "            $vb = Rect4 $brush.GetAttribute('Viewbox'); $vp = Rect4 $brush.GetAttribute('Viewport')" & vbLf
    s = s & "            $nw = $(if ($info.w -gt 0) { $info.w * 96.0 / $info.dpiX } else { $vb[2] })" & vbLf
    s = s & "            $nh = $(if ($info.h -gt 0) { $info.h * 96.0 / $info.dpiY } else { $vb[3] })" & vbLf
    s = s & "            if ($brush.GetAttribute('ViewboxUnits') -eq 'RelativeToBoundingBox') {" & vbLf
    s = s & "                $vb = @(($vb[0] * $nw), ($vb[1] * $nh), ($vb[2] * $nw), ($vb[3] * $nh))" & vbLf
    s = s & "            }" & vbLf
    s = s & "            if ($vb[2] -le 0) { $vb[2] = 1 }" & vbLf
    s = s & "            if ($vb[3] -le 0) { $vb[3] = 1 }" & vbLf
    s = s & "            $sx = $vp[2] / $vb[2]; $sy = $vp[3] / $vb[3]" & vbLf
    s = s & "            $id = NewId 'p'" & vbLf
    s = s & "            [void]$Defs.Append('<pattern id=""' + $id + '"" patternUnits=""userSpaceOnUse"" x=""' + (N $vp[0]) + '"" y=""' + (N $vp[1]) + '"" width=""' + (N $vp[2]) + '"" height=""' + (N $vp[3]) + '""')" & vbLf
    s = s & "            $m = BrushTransform $brush" & vbLf
    s = s & "            if ($m) { [void]$Defs.Append(' patternTransform=""' + (MatText $m) + '""') }" & vbLf
    s = s & "            [void]$Defs.Append('><image x=""' + (N (-$vb[0] * $sx)) + '"" y=""' + (N (-$vb[1] * $sy)) + '"" width=""' + (N ($nw * $sx)) + '"" height=""' + (N ($nh * $sy)) + '"" preserveAspectRatio=""none"" xlink:href=""data:' + $info.mime + ';base64,' + [Convert]::ToBase64String($bytes) + '""/></pattern>')" & vbLf
    s = s & "            return @{ paint = 'url(#' + $id + ')'; a = $op }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return @{ paint = 'none'; a = 1.0 }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function BrushOf($node, [string]$attr, [string]$prop, [string]$base) {" & vbLf
    s = s & "    $v = $node.GetAttribute($attr)" & vbLf
    s = s & "    if ($v) {" & vbLf
    s = s & "        if ($v -match '^\{StaticResource\s+(.+)\}$') { return (Brush $Resources[$matches[1].Trim()] $base) }" & vbLf
    s = s & "        $c = Color $v" & vbLf
    s = s & "        return @{ paint = $c.rgb; a = $c.a }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    foreach ($c in $node.ChildNodes) {" & vbLf
    s = s & "        if ($c.LocalName -eq $prop) {" & vbLf
    s = s & "            foreach ($b in $c.ChildNodes) { if ($b.NodeType -eq 'Element') { return (Brush $b $base) } }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $null" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ fonts" & vbLf
    s = s & "function U16([byte[]]$b, [int]$o) { return ([int]$b[$o] -shl 8) -bor $b[$o + 1] }" & vbLf
    s = s & "function U32([byte[]]$b, [int]$o) { return ([int64]$b[$o] -shl 24) -bor ([int64]$b[$o + 1] -shl 16) -bor ([int64]$b[$o + 2] -shl 8) -bor [int64]$b[$o + 3] }" & vbLf
    s = s & "function S16([byte[]]$b, [int]$o) { $v = U16 $b $o; if ($v -ge 32768) { $v -= 65536 }; return $v }" & vbLf
    s = s & "" & vbLf
    s = s & "function FontTables([byte[]]$b) {" & vbLf
    s = s & "    $t = @{}" & vbLf
    s = s & "    if ($b.Length -lt 12) { return $t }" & vbLf
    s = s & "    $off = 0" & vbLf
    s = s & "    if ([Text.Encoding]::ASCII.GetString($b, 0, 4) -eq 'ttcf') { $off = [int](U32 $b 12) }" & vbLf
    s = s & "    $n = U16 $b ($off + 4)" & vbLf
    s = s & "    for ($i = 0; $i -lt $n; $i++) {" & vbLf
    s = s & "        $r = $off + 12 + 16 * $i" & vbLf
    s = s & "        if ($r + 16 -gt $b.Length) { break }" & vbLf
    s = s & "        $t[[Text.Encoding]::ASCII.GetString($b, $r, 4)] = @([int](U32 $b ($r + 8)), [int](U32 $b ($r + 12)))" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $t" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function IsFont([byte[]]$b) {" & vbLf
    s = s & "    if ($b.Length -lt 4) { return $false }" & vbLf
    s = s & "    $tag = [Text.Encoding]::ASCII.GetString($b, 0, 4)" & vbLf
    s = s & "    return ($tag -eq 'OTTO' -or $tag -eq 'true' -or $tag -eq 'ttcf' -or ($b[0] -eq 0 -and $b[1] -eq 1 -and $b[2] -eq 0 -and $b[3] -eq 0))" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Deobfuscate([byte[]]$b, [string]$part) {" & vbLf
    s = s & "    if (IsFont $b) { return $b }" & vbLf
    s = s & "    $m = [regex]::Match($part, '([0-9A-Fa-f]{8})-?([0-9A-Fa-f]{4})-?([0-9A-Fa-f]{4})-?([0-9A-Fa-f]{4})-?([0-9A-Fa-f]{12})')" & vbLf
    s = s & "    if (-not $m.Success) { return $b }" & vbLf
    s = s & "    $hex = ($m.Groups[1].Value + $m.Groups[2].Value + $m.Groups[3].Value + $m.Groups[4].Value + $m.Groups[5].Value)" & vbLf
    s = s & "    $fwd = New-Object byte[] 16" & vbLf
    s = s & "    for ($i = 0; $i -lt 16; $i++) { $fwd[$i] = [Convert]::ToByte($hex.Substring(2 * $i, 2), 16) }" & vbLf
    s = s & "    $rev = New-Object byte[] 16" & vbLf
    s = s & "    for ($i = 0; $i -lt 16; $i++) { $rev[$i] = $fwd[15 - $i] }" & vbLf
    s = s & "    foreach ($key in @($rev, $fwd)) {" & vbLf
    s = s & "        $c = [byte[]]$b.Clone()" & vbLf
    s = s & "        for ($i = 0; $i -lt 32 -and $i -lt $c.Length; $i++) { $c[$i] = $c[$i] -bxor $key[$i % 16] }" & vbLf
    s = s & "        if (IsFont $c) { return $c }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $b" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function FontInfo([string]$part) {" & vbLf
    s = s & "    if ($Fonts.ContainsKey($part)) { return $Fonts[$part] }" & vbLf
    s = s & "    $info = @{ family = 'sans-serif'; local = ''; weight = 400; italic = $false; upm = 1000; adv = $null; cmap = $null; rev = $null; id = ''; data = $null }" & vbLf
    s = s & "    $Fonts[$part] = $info" & vbLf
    s = s & "    $b = ReadBytes $part" & vbLf
    s = s & "    if ($null -eq $b) { return $info }" & vbLf
    s = s & "    $b = Deobfuscate $b $part" & vbLf
    s = s & "    if (-not (IsFont $b)) { return $info }" & vbLf
    s = s & "    $info.data = $b" & vbLf
    s = s & "    $t = FontTables $b" & vbLf
    s = s & "    try {" & vbLf
    s = s & "        if ($t.ContainsKey('name')) {" & vbLf
    s = s & "            $o = $t['name'][0]" & vbLf
    s = s & "            $count = U16 $b ($o + 2); $str = $o + (U16 $b ($o + 4))" & vbLf
    s = s & "            $best = @{}" & vbLf
    s = s & "            for ($i = 0; $i -lt $count; $i++) {" & vbLf
    s = s & "                $r = $o + 6 + 12 * $i" & vbLf
    s = s & "                $plat = U16 $b $r; $lang = U16 $b ($r + 4); $nid = U16 $b ($r + 6)" & vbLf
    s = s & "                $len = U16 $b ($r + 8); $off = U16 $b ($r + 10)" & vbLf
    s = s & "                if ($plat -ne 3 -or ($nid -ne 1 -and $nid -ne 16)) { continue }" & vbLf
    s = s & "                $s = [Text.Encoding]::BigEndianUnicode.GetString($b, $str + $off, $len)" & vbLf
    s = s & "                $key = $(if ($lang -eq 0x409) { 'en' } elseif ($lang -eq 0x804 -or $lang -eq 0x404) { 'zh' } else { 'x' }) + $nid" & vbLf
    s = s & "                $best[$key] = $s" & vbLf
    s = s & "            }" & vbLf
    s = s & "            foreach ($k in @('en16', 'en1', 'x16', 'x1')) { if ($best[$k]) { $info.family = $best[$k]; break } }" & vbLf
    s = s & "            foreach ($k in @('zh16', 'zh1')) { if ($best[$k] -and $best[$k] -ne $info.family) { $info.local = $best[$k]; break } }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        if ($t.ContainsKey('OS/2')) {" & vbLf
    s = s & "            $o = $t['OS/2'][0]" & vbLf
    s = s & "            $info.weight = U16 $b ($o + 4)" & vbLf
    s = s & "            $info.italic = ((U16 $b ($o + 62)) -band 1) -eq 1" & vbLf
    s = s & "        }" & vbLf
    s = s & "        if ($t.ContainsKey('head')) { $info.upm = U16 $b ($t['head'][0] + 18) }" & vbLf
    s = s & "        if ($t.ContainsKey('hhea') -and $t.ContainsKey('hmtx')) {" & vbLf
    s = s & "            $nh = U16 $b ($t['hhea'][0] + 34)" & vbLf
    s = s & "            $o = $t['hmtx'][0]" & vbLf
    s = s & "            $adv = New-Object int[] ([Math]::Max($nh, 1))" & vbLf
    s = s & "            for ($i = 0; $i -lt $nh; $i++) { $adv[$i] = U16 $b ($o + 4 * $i) }" & vbLf
    s = s & "            $info.adv = $adv" & vbLf
    s = s & "        }" & vbLf
    s = s & "        if ($t.ContainsKey('cmap')) {" & vbLf
    s = s & "            $o = $t['cmap'][0]" & vbLf
    s = s & "            $n = U16 $b ($o + 2)" & vbLf
    s = s & "            $sub = -1" & vbLf
    s = s & "            for ($i = 0; $i -lt $n; $i++) {" & vbLf
    s = s & "                $r = $o + 4 + 8 * $i" & vbLf
    s = s & "                $plat = U16 $b $r; $eid = U16 $b ($r + 2); $so = $o + [int](U32 $b ($r + 4))" & vbLf
    s = s & "                $fmt = U16 $b $so" & vbLf
    s = s & "                if ($plat -eq 3 -and ($eid -eq 1 -or $eid -eq 10 -or $eid -eq 0) -and ($fmt -eq 4 -or $fmt -eq 12)) { $sub = $so; if ($fmt -eq 12) { break } }" & vbLf
    s = s & "            }" & vbLf
    s = s & "            if ($sub -ge 0) {" & vbLf
    s = s & "                $map = @{}" & vbLf
    s = s & "                $fmt = U16 $b $sub" & vbLf
    s = s & "                if ($fmt -eq 4) {" & vbLf
    s = s & "                    $segX2 = U16 $b ($sub + 6)" & vbLf
    s = s & "                    $ends = $sub + 14; $starts = $ends + $segX2 + 2; $deltas = $starts + $segX2; $ranges = $deltas + $segX2" & vbLf
    s = s & "                    for ($s = 0; $s -lt $segX2 / 2; $s++) {" & vbLf
    s = s & "                        $e = U16 $b ($ends + 2 * $s); $st = U16 $b ($starts + 2 * $s)" & vbLf
    s = s & "                        $dl = U16 $b ($deltas + 2 * $s); $ro = U16 $b ($ranges + 2 * $s)" & vbLf
    s = s & "                        if ($st -eq 0xFFFF) { continue }" & vbLf
    s = s & "                        for ($c = $st; $c -le $e; $c++) {" & vbLf
    s = s & "                            if ($ro -eq 0) { $g = ($c + $dl) % 65536 }" & vbLf
    s = s & "                            else {" & vbLf
    s = s & "                                $ga = $ranges + 2 * $s + $ro + 2 * ($c - $st)" & vbLf
    s = s & "                                $g = U16 $b $ga" & vbLf
    s = s & "                                if ($g -ne 0) { $g = ($g + $dl) % 65536 }" & vbLf
    s = s & "                            }" & vbLf
    s = s & "                            if ($g -ne 0) { $map[$c] = $g }" & vbLf
    s = s & "                        }" & vbLf
    s = s & "                    }" & vbLf
    s = s & "                }" & vbLf
    s = s & "                else {" & vbLf
    s = s & "                    $ng = [int](U32 $b ($sub + 12))" & vbLf
    s = s & "                    for ($s = 0; $s -lt $ng; $s++) {" & vbLf
    s = s & "                        $r = $sub + 16 + 12 * $s" & vbLf
    s = s & "                        $st = [int](U32 $b $r); $e = [int](U32 $b ($r + 4)); $g0 = [int](U32 $b ($r + 8))" & vbLf
    s = s & "                        if ($e - $st -gt 70000) { continue }" & vbLf
    s = s & "                        for ($c = $st; $c -le $e; $c++) { $map[$c] = $g0 + ($c - $st) }" & vbLf
    s = s & "                    }" & vbLf
    s = s & "                }" & vbLf
    s = s & "                $info.cmap = $map" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    catch { }" & vbLf
    s = s & "    $info.id = NewId 'f'" & vbLf
    s = s & "    if (-not $NoEmbedFonts -and [Text.Encoding]::ASCII.GetString($b, 0, 4) -ne 'ttcf') {" & vbLf
    s = s & "        [void]$Defs.Append('<style>@font-face{font-family:''' + $info.id + ''';src:url(data:font/ttf;base64,' + [Convert]::ToBase64String($b) + ');}</style>')" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $info" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function GlyphAdvance($info, [int]$gid) {" & vbLf
    s = s & "    if ($null -eq $info.adv -or $info.adv.Count -eq 0) { return 0.0 }" & vbLf
    s = s & "    if ($gid -ge $info.adv.Count) { $gid = $info.adv.Count - 1 }" & vbLf
    s = s & "    if ($gid -lt 0) { return 0.0 }" & vbLf
    s = s & "    return $info.adv[$gid] / [double]$info.upm" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function GidOf($info, [int]$cp) {" & vbLf
    s = s & "    if ($null -eq $info.cmap) { return -1 }" & vbLf
    HelperScript3 = s
End Function

Private Function HelperScript4() As String
    Dim s As String
    s = s & "    $g = $info.cmap[$cp]" & vbLf
    s = s & "    if ($null -eq $g) { return -1 }" & vbLf
    s = s & "    return [int]$g" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function CharFromGid($info, [int]$gid) {" & vbLf
    s = s & "    if ($null -eq $info.cmap) { return $null }" & vbLf
    s = s & "    if ($null -eq $info.rev) {" & vbLf
    s = s & "        $info.rev = @{}" & vbLf
    s = s & "        foreach ($k in $info.cmap.Keys) { if (-not $info.rev.ContainsKey($info.cmap[$k])) { $info.rev[$info.cmap[$k]] = $k } }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    $cp = $info.rev[$gid]" & vbLf
    s = s & "    if ($null -eq $cp) { return $null }" & vbLf
    s = s & "    return [char]::ConvertFromUtf32([int]$cp)" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ elements" & vbLf
    s = s & "function TransformOf($node, [string]$prop) {" & vbLf
    s = s & "    $t = $node.GetAttribute('RenderTransform')" & vbLf
    s = s & "    if ($t -match '^\{StaticResource\s+(.+)\}$') {" & vbLf
    s = s & "        $r = $Resources[$matches[1].Trim()]" & vbLf
    s = s & "        if ($r) { return (ParseMatrix $r.GetAttribute('Matrix')) }" & vbLf
    s = s & "        return $null" & vbLf
    s = s & "    }" & vbLf
    s = s & "    if ($t) { return (ParseMatrix $t) }" & vbLf
    s = s & "    foreach ($c in $node.ChildNodes) {" & vbLf
    s = s & "        if ($c.LocalName -eq $prop) {" & vbLf
    s = s & "            foreach ($m in $c.ChildNodes) { if ($m.LocalName -eq 'MatrixTransform') { return (ParseMatrix $m.GetAttribute('Matrix')) } }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $null" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function CommonAttrs($node, [string]$kind, $ctm, [ref]$mOut, [string]$base) {" & vbLf
    s = s & "    # transform, clip and opacity of an element -> SVG attribute string; mOut receives the new CTM" & vbLf
    s = s & "    $s = ''" & vbLf
    s = s & "    $m = TransformOf $node ($kind + '.RenderTransform')" & vbLf
    s = s & "    $cur = $ctm" & vbLf
    s = s & "    if ($m) { $s += ' transform=""' + (MatText $m) + '""'; $cur = MatMul $m $ctm }" & vbLf
    s = s & "    $mOut.Value = $cur" & vbLf
    s = s & "    $clip = Geometry $node.GetAttribute('Clip') $node ($kind + '.Clip')" & vbLf
    s = s & "    if ($clip -and $clip.d) {" & vbLf
    s = s & "        $id = NewId 'c'" & vbLf
    s = s & "        [void]$Defs.Append('<clipPath id=""' + $id + '""><path d=""' + $clip.d + '"" clip-rule=""' + $clip.rule + '""/></clipPath>')" & vbLf
    s = s & "        $s += ' clip-path=""url(#' + $id + ')""'" & vbLf
    s = s & "    }" & vbLf
    s = s & "    $op = 1.0" & vbLf
    s = s & "    if ($node.GetAttribute('Opacity')) { $op = Num $node.GetAttribute('Opacity') }" & vbLf
    s = s & "    $mask = BrushOf $node 'OpacityMask' ($kind + '.OpacityMask') $base" & vbLf
    s = s & "    if ($mask -and $mask.paint -match '^#') { $op *= $mask.a }" & vbLf
    s = s & "    if ($op -lt 0.999) { $s += ' opacity=""' + (N $op) + '""' }" & vbLf
    s = s & "    return $s" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function LoadResources($node, [string]$base) {" & vbLf
    s = s & "    foreach ($c in $node.ChildNodes) {" & vbLf
    s = s & "        if ($c.LocalName -notlike '*.Resources') { continue }" & vbLf
    s = s & "        foreach ($d in $c.ChildNodes) {" & vbLf
    s = s & "            if ($d.LocalName -ne 'ResourceDictionary') { continue }" & vbLf
    s = s & "            $src = $d.GetAttribute('Source')" & vbLf
    s = s & "            if ($src) {" & vbLf
    s = s & "                $part = Resolve $base $src" & vbLf
    s = s & "                $x = ReadXml $part" & vbLf
    s = s & "                if ($x) { foreach ($r in $x.DocumentElement.ChildNodes) { if ($r.NodeType -eq 'Element') { $k = $r.GetAttribute('Key', $KeyNs); if ($k) { $Resources[$k] = $r } } } }" & vbLf
    s = s & "            }" & vbLf
    s = s & "            foreach ($r in $d.ChildNodes) {" & vbLf
    s = s & "                if ($r.NodeType -ne 'Element') { continue }" & vbLf
    s = s & "                $k = $r.GetAttribute('Key', $KeyNs)" & vbLf
    s = s & "                if ($k) { $Resources[$k] = $r }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function IsPageBackground($pts, $ctm) {" & vbLf
    s = s & "    if ($null -eq $pts -or $pts.Count -lt 4) { return $false }" & vbLf
    s = s & "    $x0 = [double]::MaxValue; $y0 = [double]::MaxValue; $x1 = -[double]::MaxValue; $y1 = -[double]::MaxValue" & vbLf
    s = s & "    for ($i = 0; $i + 1 -lt $pts.Count; $i += 2) {" & vbLf
    s = s & "        $x = $ctm[0] * $pts[$i] + $ctm[2] * $pts[$i + 1] + $ctm[4]" & vbLf
    s = s & "        $y = $ctm[1] * $pts[$i] + $ctm[3] * $pts[$i + 1] + $ctm[5]" & vbLf
    s = s & "        $x0 = [Math]::Min($x0, $x); $y0 = [Math]::Min($y0, $y); $x1 = [Math]::Max($x1, $x); $y1 = [Math]::Max($y1, $y)" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return ($x0 -le 0.5 -and $y0 -le 0.5 -and $x1 -ge $script:PageW - 0.5 -and $y1 -ge $script:PageH - 0.5)" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function DoPath($node, $ctm, [string]$base, [System.Text.StringBuilder]$sb) {" & vbLf
    s = s & "    $g = Geometry $node.GetAttribute('Data') $node 'Path.Data'" & vbLf
    s = s & "    if ($null -eq $g -or -not $g.d) { return }" & vbLf
    s = s & "    $pts = FigureBounds $g.d" & vbLf
    s = s & "    $m = $null" & vbLf
    s = s & "    $attrs = CommonAttrs $node 'Path' $ctm ([ref]$m) $base" & vbLf
    s = s & "    if (IsPageBackground $pts $m) { return }" & vbLf
    s = s & "    $fill = BrushOf $node 'Fill' 'Path.Fill' $base" & vbLf
    s = s & "    $stroke = BrushOf $node 'Stroke' 'Path.Stroke' $base" & vbLf
    s = s & "    [void]$sb.Append('<path d=""' + $g.d + '""' + $attrs)" & vbLf
    s = s & "    if ($null -eq $fill -or $fill.paint -eq 'none') { [void]$sb.Append(' fill=""none""') }" & vbLf
    s = s & "    else {" & vbLf
    s = s & "        [void]$sb.Append(' fill=""' + $fill.paint + '""')" & vbLf
    s = s & "        if ($fill.a -lt 0.999) { [void]$sb.Append(' fill-opacity=""' + (N $fill.a) + '""') }" & vbLf
    s = s & "        if ($g.rule -eq 'evenodd') { [void]$sb.Append(' fill-rule=""evenodd""') }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    $pad = 0.0" & vbLf
    s = s & "    if ($stroke -and $stroke.paint -ne 'none') {" & vbLf
    s = s & "        $w = 1.0" & vbLf
    s = s & "        if ($node.GetAttribute('StrokeThickness')) { $w = Num $node.GetAttribute('StrokeThickness') }" & vbLf
    s = s & "        $pad = $w / 2" & vbLf
    s = s & "        [void]$sb.Append(' stroke=""' + $stroke.paint + '"" stroke-width=""' + (N $w) + '""')" & vbLf
    s = s & "        if ($stroke.a -lt 0.999) { [void]$sb.Append(' stroke-opacity=""' + (N $stroke.a) + '""') }" & vbLf
    s = s & "        $join = $node.GetAttribute('StrokeLineJoin')" & vbLf
    s = s & "        if ($join -eq 'Round') { [void]$sb.Append(' stroke-linejoin=""round""') }" & vbLf
    s = s & "        elseif ($join -eq 'Bevel') { [void]$sb.Append(' stroke-linejoin=""bevel""') }" & vbLf
    s = s & "        else {" & vbLf
    s = s & "            $ml = 10.0" & vbLf
    s = s & "            if ($node.GetAttribute('StrokeMiterLimit')) { $ml = Num $node.GetAttribute('StrokeMiterLimit') }" & vbLf
    s = s & "            [void]$sb.Append(' stroke-miterlimit=""' + (N ([Math]::Max($ml, 1.0))) + '""')" & vbLf
    s = s & "        }" & vbLf
    s = s & "        $cap = $node.GetAttribute('StrokeStartLineCap')" & vbLf
    s = s & "        if ($cap -eq 'Round') { [void]$sb.Append(' stroke-linecap=""round""') }" & vbLf
    s = s & "        elseif ($cap -eq 'Square') { [void]$sb.Append(' stroke-linecap=""square""') }" & vbLf
    s = s & "        $dash = $node.GetAttribute('StrokeDashArray')" & vbLf
    s = s & "        if ($dash) {" & vbLf
    s = s & "            $dv = Numbers $dash" & vbLf
    s = s & "            if ($dv.Count -gt 0) {" & vbLf
    s = s & "                [void]$sb.Append(' stroke-dasharray=""' + (($dv | ForEach-Object { N ([Math]::Max($_ * $w, 0.01)) }) -join ' ') + '""')" & vbLf
    s = s & "                if ($node.GetAttribute('StrokeDashOffset')) { [void]$sb.Append(' stroke-dashoffset=""' + (N ((Num $node.GetAttribute('StrokeDashOffset')) * $w)) + '""') }" & vbLf
    s = s & "                if ($node.GetAttribute('StrokeDashCap') -eq 'Round') { [void]$sb.Append(' stroke-linecap=""round""') }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    [void]$sb.Append('/>')" & vbLf
    s = s & "    if (($fill -and $fill.paint -ne 'none') -or ($stroke -and $stroke.paint -ne 'none')) { AddBox $m $pts $pad }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function DoGlyphs($node, $ctm, [string]$base, [System.Text.StringBuilder]$sb) {" & vbLf
    s = s & "    $fontPart = Resolve $base $node.GetAttribute('FontUri')" & vbLf
    s = s & "    $info = FontInfo $fontPart" & vbLf
    s = s & "    $em = Num $node.GetAttribute('FontRenderingEmSize')" & vbLf
    s = s & "    $ox = Num $node.GetAttribute('OriginX'); $oy = Num $node.GetAttribute('OriginY')" & vbLf
    s = s & "    $text = $node.GetAttribute('UnicodeString')" & vbLf
    s = s & "    if ($text.StartsWith('{}')) { $text = $text.Substring(2) }" & vbLf
    s = s & "    $idx = $node.GetAttribute('Indices')" & vbLf
    s = s & "    $entries = @()" & vbLf
    s = s & "    if ($idx) { $entries = $idx.Split(';') }" & vbLf
    s = s & "    $hasClusters = $idx -match '\('" & vbLf
    s = s & "    $rtl = $false" & vbLf
    s = s & "    if ($node.GetAttribute('BidiLevel')) { $rtl = ((Num $node.GetAttribute('BidiLevel')) % 2) -eq 1 }" & vbLf
    s = s & "    # code points of the text" & vbLf
    s = s & "    $cps = New-Object System.Collections.Generic.List[int]" & vbLf
    s = s & "    for ($i = 0; $i -lt $text.Length; $i++) {" & vbLf
    s = s & "        if ([char]::IsHighSurrogate($text[$i]) -and $i + 1 -lt $text.Length) { $cps.Add([char]::ConvertToUtf32($text[$i], $text[$i + 1])); $i++ }" & vbLf
    s = s & "        else { $cps.Add([int]$text[$i]) }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    if ($cps.Count -eq 0 -and $entries.Count -gt 0) {" & vbLf
    s = s & "        # glyph indices only: recover the characters from the font's cmap" & vbLf
    s = s & "        $chars = ''" & vbLf
    s = s & "        foreach ($e in $entries) {" & vbLf
    s = s & "            $gi = ($e.Split(',')[0] -replace '\(.*\)', '').Trim()" & vbLf
    s = s & "            if ($gi -match '^\d+$') { $c = CharFromGid $info ([int]$gi); if ($c) { $chars += $c } }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        if (-not $chars) { return }" & vbLf
    s = s & "        $text = $chars" & vbLf
    s = s & "        for ($i = 0; $i -lt $text.Length; $i++) {" & vbLf
    s = s & "            if ([char]::IsHighSurrogate($text[$i]) -and $i + 1 -lt $text.Length) { $cps.Add([char]::ConvertToUtf32($text[$i], $text[$i + 1])); $i++ }" & vbLf
    s = s & "            else { $cps.Add([int]$text[$i]) }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    if ($cps.Count -eq 0) { return }" & vbLf
    s = s & "    # x position of every character from the advance widths" & vbLf
    s = s & "    $xs = New-Object System.Collections.Generic.List[string]" & vbLf
    s = s & "    $x = $ox" & vbLf
    s = s & "    $dir = $(if ($rtl) { -1.0 } else { 1.0 })" & vbLf
    s = s & "    $total = 0.0" & vbLf
    s = s & "    for ($i = 0; $i -lt $cps.Count; $i++) {" & vbLf
    s = s & "        $gid = -1; $adv = $null; $uoff = 0.0" & vbLf
    s = s & "        if (-not $hasClusters -and $i -lt $entries.Count) {" & vbLf
    s = s & "            $f = $entries[$i].Split(',')" & vbLf
    s = s & "            if ($f[0].Trim() -match '^\d+$') { $gid = [int]$f[0].Trim() }" & vbLf
    s = s & "            if ($f.Count -gt 1 -and $f[1].Trim()) { $adv = (Num $f[1]) / 100.0 }" & vbLf
    s = s & "            if ($f.Count -gt 2 -and $f[2].Trim()) { $uoff = (Num $f[2]) / 100.0 * $em }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        if ($gid -lt 0) { $gid = GidOf $info $cps[$i] }" & vbLf
    s = s & "        if ($null -eq $adv) { $adv = GlyphAdvance $info $gid }" & vbLf
    s = s & "        $xs.Add((N ($x + $dir * $uoff)))" & vbLf
    s = s & "        $x += $dir * $adv * $em" & vbLf
    s = s & "        $total += $adv * $em" & vbLf
    s = s & "    }" & vbLf
    s = s & "    $fill = BrushOf $node 'Fill' 'Glyphs.Fill' $base" & vbLf
    s = s & "    if ($null -eq $fill) { $fill = @{ paint = '#000000'; a = 1.0 } }" & vbLf
    s = s & "    $m = $null" & vbLf
    s = s & "    $attrs = CommonAttrs $node 'Glyphs' $ctm ([ref]$m) $base" & vbLf
    s = s & "    $fam = ""'"" + $info.family.Replace(""'"", '') + ""'""" & vbLf
    s = s & "    if ($info.local) { $fam += "", '"" + $info.local.Replace(""'"", '') + ""'"" }" & vbLf
    s = s & "    if ($info.data -and -not $NoEmbedFonts) { $fam = $info.id + ', ' + $fam }" & vbLf
    s = s & "    $weight = $info.weight" & vbLf
    s = s & "    $sim = $node.GetAttribute('StyleSimulations')" & vbLf
    s = s & "    if ($sim -match 'Bold') { $weight = 700 }" & vbLf
    s = s & "    [void]$sb.Append('<text xml:space=""preserve""' + $attrs + ' x=""' + ($xs -join ' ') + '"" y=""' + (N $oy) + '"" font-family=""' + (Esc $fam) + '"" font-size=""' + (N $em) + '""')" & vbLf
    s = s & "    if ($weight -ge 600) { [void]$sb.Append(' font-weight=""' + $weight + '""') }" & vbLf
    s = s & "    if ($info.italic -or $sim -match 'Italic') { [void]$sb.Append(' font-style=""italic""') }" & vbLf
    s = s & "    if ($fill.paint -ne '#000000') { [void]$sb.Append(' fill=""' + $fill.paint + '""') }" & vbLf
    HelperScript4 = s
End Function

Private Function HelperScript5() As String
    Dim s As String
    s = s & "    if ($fill.a -lt 0.999) { [void]$sb.Append(' fill-opacity=""' + (N $fill.a) + '""') }" & vbLf
    s = s & "    [void]$sb.Append('>' + (Esc $text) + '</text>')" & vbLf
    s = s & "    $x0 = [Math]::Min($ox, $ox + $dir * $total); $x1 = [Math]::Max($ox, $ox + $dir * $total)" & vbLf
    s = s & "    AddBox $m @($x0, ($oy - 0.95 * $em), $x1, ($oy + 0.3 * $em)) 0.0" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function DoChildren($node, $ctm, [string]$base, [System.Text.StringBuilder]$sb) {" & vbLf
    s = s & "    foreach ($c in $node.ChildNodes) {" & vbLf
    s = s & "        if ($c.NodeType -ne 'Element') { continue }" & vbLf
    s = s & "        switch ($c.LocalName) {" & vbLf
    s = s & "            'Path' { DoPath $c $ctm $base $sb }" & vbLf
    s = s & "            'Glyphs' { DoGlyphs $c $ctm $base $sb }" & vbLf
    s = s & "            'Canvas' {" & vbLf
    s = s & "                LoadResources $c $base" & vbLf
    s = s & "                $m = $null" & vbLf
    s = s & "                $attrs = CommonAttrs $c 'Canvas' $ctm ([ref]$m) $base" & vbLf
    s = s & "                $inner = New-Object System.Text.StringBuilder" & vbLf
    s = s & "                DoChildren $c $m $base $inner" & vbLf
    s = s & "                if ($inner.Length -gt 0) { [void]$sb.Append('<g' + $attrs + '>' + $inner.ToString() + '</g>') }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ main" & vbLf
    s = s & "try {" & vbLf
    s = s & "    $page = FirstPage" & vbLf
    s = s & "    $doc = ReadXml $page" & vbLf
    s = s & "    if (-not $doc) { throw ""cannot read $page"" }" & vbLf
    s = s & "    $root = $doc.DocumentElement" & vbLf
    s = s & "    $script:PageW = Num $root.GetAttribute('Width')" & vbLf
    s = s & "    $script:PageH = Num $root.GetAttribute('Height')" & vbLf
    s = s & "    LoadResources $root $page" & vbLf
    s = s & "    DoChildren $root @(1.0, 0.0, 0.0, 1.0, 0.0, 0.0) $page $Body" & vbLf
    s = s & "    if ($Body.Length -eq 0) { throw 'empty page' }" & vbLf
    s = s & "" & vbLf
    s = s & "    # crop: the drawing's frame (from PowerPoint) united with what was actually drawn" & vbLf
    s = s & "    $crop = $null" & vbLf
    s = s & "    if ($Frame -and $Slide) {" & vbLf
    s = s & "        $f = Numbers $Frame; $s = Numbers $Slide" & vbLf
    s = s & "        if ($f.Count -eq 4 -and $s.Count -eq 2 -and $s[0] -gt 0) {" & vbLf
    s = s & "            $k = $script:PageW / $s[0]" & vbLf
    s = s & "            $crop = @(($f[0] * $k), ($f[1] * $k), (($f[0] + $f[2]) * $k), (($f[1] + $f[3]) * $k))" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    if ($script:Box) {" & vbLf
    s = s & "        $b = @([Math]::Max($script:Box[0], 0.0), [Math]::Max($script:Box[1], 0.0)," & vbLf
    s = s & "            [Math]::Min($script:Box[2], $script:PageW), [Math]::Min($script:Box[3], $script:PageH))" & vbLf
    s = s & "        if ($null -eq $crop) { $crop = $b }" & vbLf
    s = s & "        elseif ($b[2] -gt $b[0] -and $b[3] -gt $b[1]) {" & vbLf
    s = s & "            $crop = @([Math]::Min($crop[0], $b[0]), [Math]::Min($crop[1], $b[1]), [Math]::Max($crop[2], $b[2]), [Math]::Max($crop[3], $b[3]))" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    if ($null -eq $crop) { $crop = @(0.0, 0.0, $script:PageW, $script:PageH) }" & vbLf
    s = s & "    $cw = [Math]::Max($crop[2] - $crop[0], 0.1); $ch = [Math]::Max($crop[3] - $crop[1], 0.1)" & vbLf
    s = s & "" & vbLf
    s = s & "    $svg = New-Object System.Text.StringBuilder" & vbLf
    s = s & "    [void]$svg.Append('<?xml version=""1.0"" encoding=""UTF-8""?>' + ""`n"")" & vbLf
    s = s & "    [void]$svg.Append('<svg xmlns=""http://www.w3.org/2000/svg"" xmlns:xlink=""http://www.w3.org/1999/xlink"" version=""1.1"" width=""' + (N ($cw * 0.75)) + 'pt"" height=""' + (N ($ch * 0.75)) + 'pt"" viewBox=""' + (N $crop[0]) + ' ' + (N $crop[1]) + ' ' + (N $cw) + ' ' + (N $ch) + '"">' + ""`n"")" & vbLf
    s = s & "    [void]$svg.Append('<!-- Created by VecStamp (https://github.com/vluckyzhang/VecStamp) from PowerPoint XPS output -->' + ""`n"")" & vbLf
    s = s & "    if ($Defs.Length -gt 0) { [void]$svg.Append('<defs>' + $Defs.ToString() + '</defs>' + ""`n"") }" & vbLf
    s = s & "    [void]$svg.Append($Body.ToString() + ""`n</svg>`n"")" & vbLf
    s = s & "    $enc = New-Object System.Text.UTF8Encoding($false)" & vbLf
    s = s & "    [IO.File]::WriteAllText($Out, $svg.ToString(), $enc)" & vbLf
    s = s & "    $Zip.Dispose()" & vbLf
    s = s & "    exit 0" & vbLf
    s = s & "}" & vbLf
    s = s & "catch {" & vbLf
    s = s & "    try { $Zip.Dispose() } catch { }" & vbLf
    s = s & "    [Console]::Error.WriteLine('xps2svg: ' + $_.Exception.Message)" & vbLf
    s = s & "    exit 1" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    HelperScript5 = s
End Function

Private Function HelperScript() As String
    HelperScript = HelperScript1() & HelperScript2() & HelperScript3() & HelperScript4() & HelperScript5()
End Function
' <<< GENERATED

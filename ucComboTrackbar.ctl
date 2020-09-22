VERSION 5.00
Begin VB.UserControl ucComboTrackbar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1245
   FillStyle       =   0  'Solid
   ScaleHeight     =   53
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   83
End
Attribute VB_Name = "ucComboTrackbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' User control:  ucComboTrackbar.ctl
' Author:        Carles P.V. - 2005 (*)
' Dependencies:  None
' Last revision: 03.11.2005
' Version:       1.2.5
'----------------------------------------------------------------------------------------
'
' (*) 1. Self-Subclassing UserControl template (IDE safe) by Paul Caton:
'
'        Self-subclassing Controls/Forms - NO dependencies
'        http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'
'     2. pvCheckEnvironment() and pvIsLuna() routines by Paul Caton
'
'----------------------------------------------------------------------------------------
'
' History:
'
'   * 1.0.0: - First release.
'   * 1.1.0: - Added EditDelayedUpdate and EditDelay props. (disabled by default).
'            - Automatic decimal symbol conversion to match regional settings.
'              (see m_oEdit_KeyPress routine)
'   * 1.2.0: - Forgotten 'Value' (RO prop.) added.
'              This way you always get real current value instead of Edit contents.
'   * 1.2.1: - Explicitly destroying m_oEdit and m_oFont objects on _Terminate().
'   * 1.2.2: - Minor fix: themed button painting (hot state not restored sometimes).
'            - Fixed _Show() event. Removed Ambient.UserMode condition.
'   * 1.2.3: - Style metrics are now calculated on _Resize(). Because of different
'              event order (_ReadProperties() and _Resize() in design/runtime modes)
'              this could cause not valid metrics for a given style.
'   * 1.2.4: - No need to get channel and thumb rectangles via TBM_GETCHANNELRECT and
'              TBM_GETTHUMBRECT. Custom draw structure already passes those rectangles.
'   * 1.2.5: - I've not been able to find out why, but it seems that problem is fixed.
'              Stopped subclassing of parent window after Trackbar window.
'----------------------------------------------------------------------------------------
'
' Known issues:
'
'   * Trackbar maximum integer-range lenght is limited to 32,768 'steps'.
'     So be careful which values you set as RangeXXX ones.
'     Anyway, it's supposed that trackbar is not used to deal with large ranges.
'========================================================================================





Option Explicit

Private Const VERSION_INFO As String = "1.2.5"

'========================================================================================
' Subclasser declarations
'========================================================================================

Private Enum eMsgWhen
    [MSG_AFTER] = 1                                                                     'Message calls back after the original (previous) WndProc
    [MSG_BEFORE] = 2                                                                    'Message calls back before the original (previous) WndProc
    [MSG_BEFORE_AND_AFTER] = MSG_AFTER Or MSG_BEFORE                                    'Message calls back before and after the original (previous) WndProc
End Enum

Private Type tSubData                                                                   'Subclass data type
    hWnd                             As Long                                            'Handle of the window being subclassed
    nAddrSub                         As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig                        As Long                                            'The address of the pre-existing WndProc
    nMsgCntA                         As Long                                            'Msg after table entry count
    nMsgCntB                         As Long                                            'Msg before table entry count
    aMsgTblA()                       As Long                                            'Msg after table array
    aMsgTblB()                       As Long                                            'Msg Before table array
End Type

Private sc_aSubData()                As tSubData                                        'Subclass data array
Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long



'========================================================================================
' UserControl API declarations
'========================================================================================

Private Const SM_CXVSCROLL As Long = 2

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private Type RECT
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long

Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BDR_RAISEDINNER As Long = &H4
Private Const BDR_RAISED      As Long = &H5
Private Const BDR_SUNKEN      As Long = &HA
Private Const BF_RECT         As Long = &HF
Private Const BF_FLAT         As Long = &H4000
Private Const BF_MONO         As Long = &H8000

Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Private Const DFC_SCROLL          As Long = 3
Private Const DFCS_SCROLLCOMBOBOX As Long = &H5
Private Const DFCS_INACTIVE       As Long = &H100
Private Const DFCS_PUSHED         As Long = &H200
Private Const DFCS_FLAT           As Long = &H4000
Private Const DFCS_MONO           As Long = &H8000

Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long

Private Const HWND_TOPMOST   As Long = -1
Private Const SWP_NOZORDER   As Long = &H4
Private Const SWP_NOREDRAW   As Long = &H8
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_SHOWWINDOW As Long = &H40

Private Const DT_LEFT       As Long = &H0
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_VCENTER    As Long = &H4

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long

Private Const COLOR_BTNFACE As Long = 15

Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Const TRANSPARENT As Long = 1

Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
    
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long

Private Const WM_TIMER As Long = &H113

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

'//

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0
Private Const OBJ_BITMAP     As Long = 7

Private Declare Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long

'//

Private Const WS_BORDER         As Long = &H800000
Private Const WS_CLIPSIBLINGS   As Long = &H4000000
Private Const WS_VISIBLE        As Long = &H10000000
Private Const WS_CHILD          As Long = &H40000000
Private Const WS_EX_TOOLWINDOW  As Long = &H80&

Private Const WM_SETFOCUS       As Long = &H7
Private Const WM_KILLFOCUS      As Long = &H8
Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_MOUSEACTIVATE  As Long = &H21
Private Const WM_GETMINMAXINFO  As Long = &H24
Private Const WM_NOTIFY         As Long = &H4E
Private Const WM_SYSCOMMAND     As Long = &H112
Private Const WM_HSCROLL        As Long = &H114
Private Const WM_CTLCOLOREDIT   As Long = &H133
Private Const WM_CTLCOLORSTATIC As Long = &H138
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const WM_LBUTTONUP      As Long = &H202
Private Const WM_RBUTTONDOWN    As Long = &H204
Private Const WM_RBUTTONUP      As Long = &H205
Private Const WM_MOUSEWHEEL     As Long = &H20A
Private Const WM_THEMECHANGED   As Long = &H31A

Private Const MK_LBUTTON        As Long = &H1

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

'//

'- Trackbar class name
Private Const WC_TRACKBAR32      As String = "msctls_trackbar32"
 
'- Trackbar styles
Private Const TBS_BOTH           As Long = &H8
Private Const TBS_NOTICKS        As Long = &H10
Private Const TBS_FIXEDLENGTH    As Long = &H40

'- Trackbar messages
Private Const WM_USER            As Long = &H400
Private Const TBM_GETPOS         As Long = WM_USER
Private Const TBM_SETPOS         As Long = WM_USER + 5
Private Const TBM_SETRANGE       As Long = WM_USER + 6
Private Const TBM_GETTHUMBRECT   As Long = WM_USER + 25
Private Const TBM_SETTHUMBLENGTH As Long = WM_USER + 27

'- Trackbar notifications
Private Const NM_FIRST           As Long = 0
Private Const NM_CUSTOMDRAW      As Long = NM_FIRST - 12

'- Trackbar 'custom draw' specifications
Private Const TBCD_TICS          As Long = &H1
Private Const TBCD_THUMB         As Long = &H2
Private Const TBCD_CHANNEL       As Long = &H3

' Notification structure
Private Type NMHDR
    hwndFrom As Long
    idfrom   As Long
    code     As Long
End Type

' 'Custom draw' structure
Private Type NMCUSTOMDRAW
    hdr         As NMHDR
    dwDrawStage As Long
    hDC         As Long
    rc          As RECT
    dwItemSpec  As Long
    uItemState  As Long
    lItemlParam As Long
End Type

'- Custom draw paint stages (only used ones)
Private Const CDDS_PREPAINT       As Long = &H1
Private Const CDDS_ITEM           As Long = &H10000
Private Const CDDS_ITEMPREPAINT   As Long = CDDS_ITEM Or CDDS_PREPAINT

'- Custom draw item states (only used ones)
Private Const CDIS_SELECTED       As Long = &H1

'- Custom draw return values (only used ones)
Private Const CDRF_SKIPDEFAULT    As Long = &H4
Private Const CDRF_NOTIFYITEMDRAW As Long = &H20

'//

'- ComboBox class string
Private Const CB_THEME As String = "ComboBox"

'- ComboBox parts
Private Const CP_DROPDOWNBUTTON As Long = 1
Private Const CP_BORDER         As Long = 2
 
'- ComboBox states
Private Const CBXS_NORMAL   As Long = 1
Private Const CBXS_HOT      As Long = 2
Private Const CBXS_PRESSED  As Long = 3
Private Const CBXS_DISABLED As Long = 4

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function GetCurrentThemeName Lib "uxtheme" (ByVal pszThemeFileName As Long, ByVal cchMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function GetThemeDocumentationProperty Lib "uxtheme" (ByVal pszThemeName As Long, ByVal pszPropertyName As Long, ByVal pszValueBuff As Long, ByVal cchMaxValChars As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long



'========================================================================================
' UserControl enums., variables and constants
'========================================================================================

'-- Public enums.:

Public Enum ctBackStyleCts
    [bsSolidColor] = 0
    [bsImage] = 1
End Enum

Public Enum ctStyleCts
    [sClassic] = 0
    [sFlat] = 1
    [sFlatMono] = 2
    [sThemed] = 3
End Enum

Public Enum ctRangePrecisionCts
    [rpInteger] = 0
    [rpTenth] = 1
    [rpHundredth] = 2
    [rpThousandth] = 3
End Enum

Public Enum ctTrackbarPositionCts
    [tpWide] = 0
    [tpCentered] = 1
End Enum

'-- Private constants:

Private Const ERR_OVERFLOW          As Long = 6
Private Const INTEGER_OVERFLOW      As Long = 32768
Private Const TRACKBAR_HEIGHT_MIN   As Long = 16
Private Const TRACKBAR_WIDTH_MIN    As Long = 32
Private Const THUMB_OFFSET          As Long = 3
Private Const FONT_EXTENT           As Long = 4
Private Const EDGE_THICK            As Long = 2
Private Const EDGE_THIN             As Long = 1
Private Const EDGE_NULL             As Long = 0
Private Const TIMERID_HOT           As Long = 1
Private Const TIMERDT_HOT           As Long = 25
Private Const TIMERID_EDIT          As Long = 2
Private Const TIMERMINDT_EDIT       As Long = 250

'-- Private variables:

Private WithEvents m_oEdit          As TextBox
Attribute m_oEdit.VB_VarHelpID = -1
Private WithEvents m_oFont          As StdFont
Attribute m_oFont.VB_VarHelpID = -1

Private m_hWndTrackbar              As Long
Private m_hWndParent                As Long

Private m_uRctControl               As RECT
Private m_uRctEdit                  As RECT
Private m_uRctButton                As RECT
Private m_lEditEdge                 As Long
Private m_lEditExtent               As Long
Private m_lButtonEdge               As Long
Private m_lButtonExtent             As Long

Private m_bHasFocus                 As Boolean
Private m_bButtonPressed            As Boolean
Private m_bButtonHot                As Boolean

Private m_hBackBrush                As Long
Private m_hPatternBrush             As Long

Private m_lValue                    As Long
Private m_lCancelValue              As Long
Private m_lMax                      As Long
Private m_lMin                      As Long
Private m_lPrecisionFactor          As Long
Private m_sPrecisionFormat(3)       As String

Private m_bIsXP                     As Boolean
Private m_bIsLuna                   As Boolean

'-- Property variables:

Private m_oleBackColor              As OLE_COLOR
Private m_oBackImage                As StdPicture
Private m_eBackStyle                As ctBackStyleCts
Private m_bEditDelayedUpdate        As Boolean
Private m_lEditDelay                As Long
Private m_snRangeMax                As Single
Private m_snRangeMin                As Single
Private m_eRangePrecision           As ctRangePrecisionCts
Private m_eStyle                    As ctStyleCts
Private m_lTrackbarHeight           As Long
Private m_eTrackbarPosition         As ctTrackbarPositionCts
Private m_lTrackbarWidth            As Long

'-- Default property values:

Private Const BACKCOLOR_DEF         As Long = vbWindowBackground
Private Const BACKSTYLE_DEF         As Long = [bsSolidColor]
Private Const EDITDELAYEDUPDATE_DEF As Boolean = False
Private Const EDITDELAY_DEF         As Long = 1000
Private Const ENABLED_DEF           As Boolean = True
Private Const FORECOLOR_DEF         As Long = vbWindowText
Private Const LOCKED_DEF            As Boolean = False
Private Const RANGEMAX_DEF          As Single = 100
Private Const RANGEMIN_DEF          As Single = 0
Private Const RANGEPRECISION_DEF    As Long = [rpInteger]
Private Const STYLE_DEF             As Long = [sClassic]
Private Const TRACKBARHEIGHT_DEF    As Long = 22
Private Const TRACKBARPOSITION_DEF  As Long = [tpWide]
Private Const TRACKBARWIDTH_DEF     As Long = 100

'-- Events:

Public Event Change()
Public Event Scroll()
Public Event TrackbarShow()
Public Event TrackbarHide()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event ThemeChanged() ' XP only




'========================================================================================
' UserControl subclass procedure
'========================================================================================

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, _
                          ByRef bHandled As Boolean, _
                          ByRef lReturn As Long, _
                          ByRef lhWnd As Long, _
                          ByRef uMsg As Long, _
                          ByRef wParam As Long, _
                          ByRef lParam As Long _
                          )
Attribute zSubclass_Proc.VB_MemberFlags = "40"

    Select Case lhWnd
        
        Case m_hWndParent
            
            If (m_hWndTrackbar <> 0) Then
            
                Select Case uMsg
                
                    Case WM_GETMINMAXINFO, WM_SYSCOMMAND, WM_RBUTTONDOWN, WM_LBUTTONDOWN
                        Call pvDestroyTrackbar
                        
                    Case WM_MOUSEACTIVATE
                        If (pvOnMouseActivate()) Then
                            Call pvDestroyTrackbar
                        End If
                End Select
            End If
            
        Case m_hWndTrackbar
        
            Select Case uMsg
                
                Case WM_RBUTTONUP, WM_LBUTTONUP
                    Call pvDestroyTrackbar
            End Select
        
        Case UserControl.hWnd

                Select Case uMsg
    
                    Case WM_NOTIFY
                        Call pvOnNotify(lParam, bHandled, lReturn)
                        
                    Case WM_HSCROLL
                        Call pvOnHScroll
                        
                    Case WM_MOUSEWHEEL
                        Call pvOnMouseWheel(wParam)
                    
                    Case WM_LBUTTONDOWN
                        Call pvOnMouseDown(wParam, lParam)
                    
                    Case WM_LBUTTONUP
                        Call pvOnMouseUp
                     
                    Case WM_MOUSEMOVE
                        Call pvOnMouseMove(wParam, lParam)
                                               
                    Case WM_CTLCOLOREDIT
                        Call pvOnCtlColorEdit(wParam, lParam, lReturn)
                    
                    Case WM_CTLCOLORSTATIC
                        Call pvOnCtlColorStatic(wParam, lParam, lReturn)
                        
                    Case WM_TIMER
                        Call pvOnTimer(wParam)
                    
                    Case WM_THEMECHANGED
                        Call pvOnThemeChanged
                        
                    Case WM_SYSCOLORCHANGE
                        Call pvOnSysColorChange
               End Select
               
        Case m_oEdit.hWnd

            Select Case uMsg
                
                Case WM_SETFOCUS
                    Call pvOnSetFocus
                    
                Case WM_KILLFOCUS
                    Call pvOnKillFocus
            End Select
    End Select
End Sub



'========================================================================================
' UserControl initialization/termination
'========================================================================================

Private Sub UserControl_Initialize()
    
    '-- Precision formats
    m_sPrecisionFormat([rpInteger]) = "0"
    m_sPrecisionFormat([rpTenth]) = "0.0"
    m_sPrecisionFormat([rpHundredth]) = "0.00"
    m_sPrecisionFormat([rpThousandth]) = "0.000"
End Sub

Private Sub UserControl_Terminate()
    
    On Error GoTo Catch
    
    '-- Stop subclassing
    Call Subclass_StopAll
    
Catch:
    On Error GoTo 0
    
    '-- Clean up
    Set m_oEdit = Nothing
    Set m_oFont = Nothing
    Call DeleteObject(m_hBackBrush)
    Call DeleteObject(m_hPatternBrush)
End Sub



'========================================================================================
' UserControl misc.
'========================================================================================

Private Sub UserControl_DblClick()
    
    '-- Preserve second click
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
End Sub

Private Sub UserControl_Resize()
    
    '-- Resize (combo)
    Call pvResizeControl
End Sub

Private Sub UserControl_Show()

    '-- Resize (combo)
    Call pvResizeControl
End Sub



'========================================================================================
' Inherent Edit control
'========================================================================================

Private Sub m_oEdit_Change()
  
  Dim lPrevValue As Long
      
    lPrevValue = m_lValue
    
    On Error Resume Next
    
    m_lValue = m_oEdit.Text * m_lPrecisionFactor
    If (m_lValue > m_lMax) Then m_lValue = m_lMax
    If (m_lValue < m_lMin) Then m_lValue = m_lMin
    
    On Error GoTo 0
    
    Call pvUpdateTrackbar
    
    If (m_bEditDelayedUpdate) Then
        Call pvKillTimer(TIMERID_EDIT)
        Call pvSetTimer(TIMERID_EDIT, m_lEditDelay)
    End If
    
    If (m_lValue <> lPrevValue) Then
        RaiseEvent Change
    End If
End Sub

Private Sub m_oEdit_KeyDown(KeyCode As Integer, Shift As Integer)
        
    RaiseEvent KeyDown(KeyCode, Shift)
    
    Select Case KeyCode
        
        Case vbKeyReturn
            Call pvDestroyTrackbar
            KeyCode = 0
        
        Case vbKeyEscape
            Call pvDestroyTrackbar(Cancel:=True)
            Call pvUpdateEdit
            KeyCode = 0
        
        Case vbKeySpace
            Call pvSelectEditContents
            KeyCode = 0
            
        Case vbKeyDelete
            Call InvalidateRect(m_oEdit.hWnd, ByVal 0, 0)
            
        Case vbKeyDown
            If (Shift = vbAltMask) Then
                If (m_hWndTrackbar <> 0) Then
                    Call pvDestroyTrackbar
                  Else
                    Call pvCreateTrackbar
                End If
              Else
                Call pvValueDec
                Call pvUpdateEdit
                Call pvUpdateTrackbar
                Call pvSelectEditContents
            End If
            KeyCode = 0
        
        Case vbKeyUp
            Call pvValueInc
            Call pvUpdateEdit
            Call pvUpdateTrackbar
            Call pvSelectEditContents
            KeyCode = 0
    End Select
End Sub

Private Sub m_oEdit_KeyPress(KeyAscii As Integer)
        
    RaiseEvent KeyPress(KeyAscii)
    
    '-- Special keys
    Select Case KeyAscii
        
        Case vbKeyReturn, vbKeyEscape, vbKeySpace
            KeyAscii = 0
            
        Case Asc("."), Asc(",")
            KeyAscii = Asc(Mid$(CStr(1.1), 2, 1))
            Call InvalidateRect(m_oEdit.hWnd, ByVal 0, 0)
        
        Case Else
            If (m_oEdit.Locked = False) Then
                Call InvalidateRect(m_oEdit.hWnd, ByVal 0, 0)
            End If
    End Select
End Sub

Private Sub m_oEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub



'========================================================================================
' Methods
'========================================================================================

Public Sub ShowTrackbar( _
           Optional ByVal CaptureThumb As Boolean = False _
           )
           
  Dim uRectWnd   As RECT
  Dim uRectThumb As RECT
    
    If (UserControl.Enabled) Then
        If (m_hWndTrackbar = 0) Then
            Call pvCreateTrackbar
            Call SetFocus(m_oEdit.hWnd)
            If (CaptureThumb) Then
                Call GetWindowRect(m_hWndTrackbar, uRectWnd)
                Call SendMessageAny(m_hWndTrackbar, TBM_GETTHUMBRECT, 0, uRectThumb)
                With uRectThumb
                    Call SetCursorPos(uRectWnd.x1 + .x1 + (.x2 - .x1) \ 2, uRectWnd.y1 + (.y2 - .y1) \ 2)
                    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                End With
            End If
        End If
    End If
End Sub

Public Sub HideTrackbar()
    If (m_hWndTrackbar <> 0) Then
        Call pvDestroyTrackbar
    End If
End Sub

Public Sub About()
Attribute About.VB_UserMemId = -552
    Call VBA.MsgBox("ucComboTrackbar " & VERSION_INFO & " - Carles P.V. 2005", , "About")
End Sub



'========================================================================================
' Private
'========================================================================================

'----------------------------------------------------------------------------------------
' Checking OS and Luna theme
'----------------------------------------------------------------------------------------

Private Sub pvCheckEnvironment()

  Dim uOSV As OSVERSIONINFO
    
    m_bIsXP = False
    m_bIsLuna = False
    
    With uOSV
        
        .dwOSVersionInfoSize = Len(uOSV)
        Call GetVersionEx(uOSV)
        
        If (.dwPlatformId = 2) Then
            If (.dwMajorVersion = 5) Then     ' NT based
                If (.dwMinorVersion > 0) Then ' XP
                    m_bIsXP = True
                    m_bIsLuna = pvIsLuna()
                End If
            End If
        End If
    End With
End Sub

Private Function pvIsLuna() As Boolean

  Dim hLib   As Long
  Dim lPos   As Long
  Dim sTheme As String
  Dim sName  As String

    '-- Be sure that the theme dll is present
    hLib = LoadLibrary("uxtheme.dll")
    
    If (hLib <> 0) Then
        '-- Get the theme file name
        sTheme = String$(255, 0)
        Call GetCurrentThemeName(StrPtr(sTheme), Len(sTheme), 0, 0, 0, 0)
        lPos = InStr(1, sTheme, Chr$(0))
        
        If (lPos > 0) Then
            '-- Get the canonical theme name
            sTheme = Left$(sTheme, lPos - 1)
            sName = String$(255, 0)
            Call GetThemeDocumentationProperty(StrPtr(sTheme), StrPtr("ThemeName"), StrPtr(sName), Len(sName))
            lPos = InStr(1, sName, Chr$(0))
            
            If (lPos > 0) Then
                '-- Is it Luna?
                sName = Left$(sName, lPos - 1)
                pvIsLuna = (StrComp(sName, "Luna", vbTextCompare) = 0)
            End If
        End If
        
        Call FreeLibrary(hLib)
    End If
End Function

'----------------------------------------------------------------------------------------
' Objects initialization
'----------------------------------------------------------------------------------------
    
Private Sub pvInitializeObjects()

    '-- 'Clean' UserControl
    Set m_oEdit = UserControl.Controls.Add("VB.TextBox", "m_oEdit")
    Let m_oEdit.BorderStyle = 0
    Let m_oEdit.Alignment = vbRightJustify
    Let m_oEdit.TabStop = False
    
    '-- New *event'ed* font object
    Set m_oFont = New StdFont
End Sub

'----------------------------------------------------------------------------------------
' Trackbar
'----------------------------------------------------------------------------------------

Private Sub pvCreateTrackbar()
    
  Dim lStyleWnd As Long
  Dim lStyleExt As Long
  Dim uRectWnd  As RECT
  Dim lyOff     As Long
        
    '-- Define window style and window extended style
    lStyleWnd = WS_CHILD Or WS_CLIPSIBLINGS Or WS_BORDER Or TBS_NOTICKS Or TBS_BOTH Or TBS_FIXEDLENGTH
    lStyleExt = WS_EX_TOOLWINDOW
    
    '-- Create trackbar window
    m_hWndTrackbar = CreateWindowEx(lStyleExt, WC_TRACKBAR32, vbNullString, lStyleWnd, 0, 0, 0, 0, hWnd, 0, App.hInstance, ByVal 0&)
        
    '-- Success?
    If (m_hWndTrackbar <> 0) Then
        
        '-- Store cancel value
        m_lCancelValue = m_lValue
        
        '-- Store parent
        m_hWndParent = UserControl.Parent.hWnd
        
        '-- Subclass it
        Call Subclass_Start(m_hWndParent)
        Call Subclass_AddMsg(m_hWndParent, WM_GETMINMAXINFO)
        Call Subclass_AddMsg(m_hWndParent, WM_MOUSEACTIVATE)
        Call Subclass_AddMsg(m_hWndParent, WM_RBUTTONDOWN)
        Call Subclass_AddMsg(m_hWndParent, WM_LBUTTONDOWN)
        Call Subclass_AddMsg(m_hWndParent, WM_SYSCOMMAND, [MSG_BEFORE])
        
        '-- Resize thumb
        Call SendMessageLong(m_hWndTrackbar, TBM_SETTHUMBLENGTH, m_lTrackbarHeight - 2 * THUMB_OFFSET, 0)
        
        '-- Set trackbar range and current value
        Call SendMessageLong(m_hWndTrackbar, TBM_SETRANGE, 0, m_lMin + (m_lMax + -(m_lMin < 0)) * &H10000)
        Call SendMessageLong(m_hWndTrackbar, TBM_SETPOS, 0, m_lValue)

        '-- Ensure window over parent
        Call SetParent(m_hWndTrackbar, GetDesktopWindow())
        
        '-- Subclass it
        Call Subclass_Start(m_hWndTrackbar)
        Call Subclass_AddMsg(m_hWndTrackbar, WM_RBUTTONUP)
        Call Subclass_AddMsg(m_hWndTrackbar, WM_LBUTTONUP)
        
        '-- Show it
        With uRectWnd
            Call GetWindowRect(UserControl.hWnd, uRectWnd)
            If (.y2 + m_lTrackbarHeight > Screen.Height \ Screen.TwipsPerPixelY) Then
                lyOff = (.y2 - .y1) + m_lTrackbarHeight
            End If
            Select Case m_eTrackbarPosition
                Case [tpWide]
                    Call SetWindowPos(m_hWndTrackbar, HWND_TOPMOST, .x1, .y2 - lyOff, .x2 - .x1, m_lTrackbarHeight, SWP_SHOWWINDOW Or SWP_NOACTIVATE)
                Case [tpCentered]
                    Call SetWindowPos(m_hWndTrackbar, HWND_TOPMOST, .x1 + m_uRctButton.x1 + (m_uRctButton.x2 - m_uRctButton.x1) \ 2 - m_lTrackbarWidth \ 2, .y2 - lyOff, m_lTrackbarWidth, m_lTrackbarHeight, SWP_SHOWWINDOW Or SWP_NOACTIVATE)
            End Select
        End With
        
        '-- Repaint control
        Call pvPaintCombo
        
        RaiseEvent TrackbarShow
    End If
End Sub

Private Sub pvDestroyTrackbar(Optional ByVal Cancel As Boolean = False)
     
    '-- Trackbar visible?
    If (m_hWndTrackbar <> 0) Then
        
        '-- Stop subclassing Trackbar window and destroy it
        Call Subclass_Stop(m_hWndTrackbar)
        Call DestroyWindow(m_hWndTrackbar)
        m_hWndTrackbar = 0
    
        '-- Stop subclassing parent window
        Call Subclass_Stop(m_hWndParent)
        
        '-- Cancel?
        If (Cancel) Then
            m_lValue = m_lCancelValue
        End If
        
        Call VBA.DoEvents ' WM_SYSCOMMAND
        
        RaiseEvent TrackbarHide
    End If
    
    '-- Repaint control
    m_bButtonPressed = False
    Call pvPaintCombo
End Sub

Private Function pvGetTrackbarValue( _
                 ) As Long

    If (m_hWndTrackbar <> 0) Then
        pvGetTrackbarValue = SendMessageLong(m_hWndTrackbar, TBM_GETPOS, 0, 0)
    End If
End Function

Private Sub pvSetTrackbarValue( _
            ByVal lValue As Long _
            )

    If (m_hWndTrackbar <> 0) Then
        Call SendMessageLong(m_hWndTrackbar, TBM_SETPOS, 1, lValue)
    End If
End Sub

Private Sub pvUpdateTrackbar()

    If (m_hWndTrackbar <> 0) Then
        If (pvGetTrackbarValue() <> m_lValue) Then
            Call pvSetTrackbarValue(m_lValue)
        End If
    End If
End Sub

Private Function pvValidateRange( _
                 ByVal snMin As Single, _
                 ByVal snMax As Single, _
                 ByVal ePrecision As ctRangePrecisionCts _
                 ) As Boolean
                 
  Dim lFactor   As Long
  Dim lMin      As Long
  Dim lMax      As Long
  Dim lAbsRange As Long
    
    On Error GoTo Catch
    
    lFactor = (10 ^ ePrecision)
    lMin = snMin * lFactor
    lMax = snMax * lFactor
    lAbsRange = lMin + (lMax + -(lMin < 0))
    
    pvValidateRange = (lAbsRange < INTEGER_OVERFLOW)
     
Catch:
    On Error GoTo 0
End Function

'----------------------------------------------------------------------------------------
' Messages response
'----------------------------------------------------------------------------------------

Private Sub pvOnNotify( _
            ByVal lParam As Long, _
            bHandled As Boolean, _
            lReturn As Long _
            )

  Dim uNMH As NMHDR
  
    If (m_eStyle <> [sThemed]) Then
        Call CopyMemory(uNMH, ByVal lParam, Len(uNMH))
        If (uNMH.hwndFrom = m_hWndTrackbar) Then
            If (uNMH.code = NM_CUSTOMDRAW) Then
                lReturn = pvPaintTrackbar(lParam)
                bHandled = True
            End If
        End If
    End If
End Sub

Private Sub pvOnSetFocus()

    m_bHasFocus = True
End Sub

Private Sub pvOnKillFocus()

    m_bHasFocus = False
    Call pvDestroyTrackbar
End Sub

Private Function pvOnMouseActivate( _
                 ) As Boolean
    
  Dim uPt As POINTAPI
  
    Call GetCursorPos(uPt)
    Call ScreenToClient(UserControl.hWnd, uPt)
    
    pvOnMouseActivate = Not (PtInRect(m_uRctButton, uPt.x, uPt.y) <> 0)
End Function

Private Sub pvOnHScroll()
    
    RaiseEvent Scroll
                        
    m_lValue = pvGetTrackbarValue()
    Call pvUpdateEdit
End Sub

Private Sub pvOnMouseWheel( _
            ByVal wParam As Long _
            )

    If (wParam < 0) Then
        Call pvValueDec
        Call pvUpdateEdit
        Call pvUpdateTrackbar
      Else
        Call pvValueInc
        Call pvUpdateEdit
        Call pvUpdateTrackbar
    End If
End Sub

Private Sub pvOnMouseDown( _
            ByVal wParam As Long, _
            ByVal lParam As Long _
            )
  
  Dim x As Long
  Dim y As Long
    
    If (wParam = MK_LBUTTON) Then
        If (m_hWndTrackbar <> 0) Then
            Call pvDestroyTrackbar
          Else
            Call pvMakePoints(lParam, x, y)
            m_bButtonPressed = (PtInRect(m_uRctButton, x, y) <> 0)
            If (m_bButtonPressed) Then
                Call pvCreateTrackbar
            End If
        End If
    End If
End Sub

Private Sub pvOnMouseMove( _
            ByVal wParam As Long, _
            ByVal lParam As Long _
            )

  Dim bInButton As Boolean
  Dim x As Long
  Dim y As Long
                    
    If (m_hWndTrackbar <> 0) Then
        If (wParam = MK_LBUTTON) Then
            Call pvMakePoints(lParam, x, y)
            bInButton = (PtInRect(m_uRctButton, x, y) <> 0)
            If (m_bButtonPressed Xor bInButton) Then
                m_bButtonPressed = bInButton
                Call pvPaintCombo
            End If
        End If
      Else
        If (m_eStyle = [sThemed] And m_bIsLuna) Then
            If (m_bButtonHot = False) Then
                Call pvMakePoints(lParam, x, y)
                If (PtInRect(m_uRctButton, x, y) <> 0) Then
                    m_bButtonHot = True
                    Call pvPaintCombo
                    Call pvKillTimer(TIMERID_HOT)
                    Call pvSetTimer(TIMERID_HOT, TIMERDT_HOT)
                End If
            End If
        End If
    End If
End Sub

Private Sub pvOnMouseUp()
            
    If (m_hWndTrackbar <> 0) Then
        m_bButtonPressed = False
        Call pvPaintCombo
    End If
End Sub

Private Sub pvOnCtlColorEdit( _
            ByVal wParam As Long, _
            ByVal lParam As Long, _
            lReturn As Long _
            )
  
  Dim uPt As POINTAPI
    
    If (lParam = m_oEdit.hWnd) Then
        Call SetTextColor(wParam, pvTranslateColor(m_oEdit.ForeColor))
        Call SetBkMode(wParam, TRANSPARENT)
        Call SetBrushOrgEx(wParam, -m_oEdit.Left, -m_oEdit.Top, uPt)
        lReturn = m_hBackBrush
    End If
End Sub

Private Sub pvOnCtlColorStatic( _
            ByVal wParam As Long, _
            ByVal lParam As Long, _
            lReturn As Long _
            )

  Dim uPt As POINTAPI

    If (lParam = m_oEdit.hWnd) Then
        Call SetBkMode(wParam, TRANSPARENT)
        Call SetBrushOrgEx(wParam, -m_oEdit.Left, -m_oEdit.Top, uPt)
        lReturn = m_hBackBrush
    End If
End Sub

Private Sub pvOnTimer( _
            ByVal wParam As Long _
            )
                         
  Dim uPt As POINTAPI
    
    Select Case wParam
        
        Case TIMERID_HOT
            Call GetCursorPos(uPt)
            Call ScreenToClient(UserControl.hWnd, uPt)
            
            If (PtInRect(m_uRctButton, uPt.x, uPt.y) = 0) Then
                m_bButtonHot = False
                Call pvKillTimer(TIMERID_HOT)
                Call pvPaintCombo
            End If
            
        Case TIMERID_EDIT
            Call pvKillTimer(TIMERID_EDIT)
            m_oEdit.Text = Format$(m_lValue / m_lPrecisionFactor, m_sPrecisionFormat(m_eRangePrecision))
    End Select
End Sub

Private Sub pvOnThemeChanged()
    
    '-- Check OS and Luna theme
    Call pvCheckEnvironment
    RaiseEvent ThemeChanged
    
    '-- Update all
    Call pvResizeControl
    Call InvalidateRect(m_oEdit.hWnd, ByVal 0, 0)
End Sub

Private Sub pvOnSysColorChange()
    
    '-- Repaint all
    Call pvPaintCombo
    Call InvalidateRect(m_oEdit.hWnd, ByVal 0, 0)
End Sub

'----------------------------------------------------------------------------------------
' Sizing
'----------------------------------------------------------------------------------------

Private Sub pvResizeControl()
    
  Dim lScrollWidth As Long
    
    '-- Get combo metrics
    Call pvGetStyleMetrics
   
    '-- Get vertical-scroll button width
    lScrollWidth = GetSystemMetrics(SM_CXVSCROLL)
  
    '-- Get control rectangle (client)
    Call GetClientRect(UserControl.hWnd, m_uRctControl)
    
    '-- Calculate Edit rectangle
    Call CopyRect(m_uRctEdit, m_uRctControl)
    Call InflateRect(m_uRctEdit, -m_lEditEdge, -m_lEditEdge)
    m_uRctEdit.x2 = m_uRctEdit.x2 - lScrollWidth + m_lEditExtent + m_lButtonExtent
    
    '-- Calculate button rectangle
    Call CopyRect(m_uRctButton, m_uRctControl)
    Call InflateRect(m_uRctButton, -m_lButtonEdge, -m_lButtonEdge)
    m_uRctButton.x1 = m_uRctButton.x2 - lScrollWidth
    
    '-- Adjust control width
    If (UserControl.ScaleWidth < 2 * lScrollWidth) Then
        Let UserControl.Width = (2 * lScrollWidth) * Screen.TwipsPerPixelX
    End If
    
    '-- Adjust control height
    Select Case m_eStyle
        Case [sClassic], [sThemed]
            Let UserControl.Height = (UserControl.TextHeight(vbNullString) + 2 * FONT_EXTENT) * Screen.TwipsPerPixelY
        Case [sFlat], [sFlatMono]
            Let UserControl.Height = (UserControl.TextHeight(vbNullString) + 2 * FONT_EXTENT - 2 * EDGE_THIN) * Screen.TwipsPerPixelY
    End Select
    
    '-- Update Edit size and position
    If (Not m_oEdit Is Nothing) Then
        Call SetWindowPos(m_oEdit.hWnd, 0, m_lEditEdge + EDGE_THIN, m_lEditEdge + EDGE_THIN, m_uRctButton.x1 - 2 * m_lEditEdge - 2 * EDGE_THIN + m_lButtonExtent, m_uRctControl.y2 - 2 * m_lEditEdge - 2 * EDGE_THIN, SWP_NOACTIVATE Or SWP_NOREDRAW Or SWP_NOZORDER)
    End If
    
    '-- Repaint control
    Call pvPaintCombo
End Sub

Private Sub pvGetStyleMetrics()

    Select Case True
        
        Case m_eStyle = [sClassic] Or (m_eStyle = [sThemed] And m_bIsLuna = False)
            m_lEditEdge = EDGE_THICK
            m_lButtonEdge = EDGE_THICK
            m_lEditExtent = EDGE_NULL
            m_lButtonExtent = EDGE_NULL
        
        Case m_eStyle = [sFlat], m_eStyle = [sFlatMono]
            m_lEditEdge = EDGE_THIN
            m_lButtonEdge = EDGE_NULL
            m_lEditExtent = EDGE_THIN
            m_lButtonExtent = EDGE_NULL
        
        Case m_eStyle = [sThemed]
            m_lEditEdge = EDGE_THICK
            m_lButtonEdge = EDGE_THIN
            m_lEditExtent = EDGE_THIN
            m_lButtonExtent = EDGE_THIN
    End Select
End Sub

'----------------------------------------------------------------------------------------
' Painting
'----------------------------------------------------------------------------------------

Private Sub pvPaintCombo()
  
  Dim uRct As RECT
    
    Select Case True
        
        Case m_eStyle = [sClassic] Or (m_eStyle = [sThemed] And m_bIsLuna = False)
            
            Call DrawEdge(UserControl.hDC, m_uRctControl, BDR_SUNKEN, BF_RECT)
            Call FillRect(UserControl.hDC, m_uRctEdit, m_hBackBrush)
            Call DrawFrameControl(UserControl.hDC, m_uRctButton, DFC_SCROLL, DFCS_SCROLLCOMBOBOX Or DFCS_PUSHED * -(m_bButtonPressed) Or DFCS_INACTIVE * -(Not UserControl.Enabled))
       
        Case m_eStyle = [sFlat]
        
            Call DrawEdge(UserControl.hDC, m_uRctControl, BDR_SUNKEN, BF_RECT Or BF_FLAT)
            Call FillRect(UserControl.hDC, m_uRctEdit, m_hBackBrush)
            Call DrawFrameControl(UserControl.hDC, m_uRctButton, DFC_SCROLL, DFCS_SCROLLCOMBOBOX Or DFCS_FLAT Or DFCS_PUSHED * -(m_bButtonPressed) Or DFCS_INACTIVE * -(Not UserControl.Enabled))
        
        Case m_eStyle = [sFlatMono]
        
            Call DrawEdge(UserControl.hDC, m_uRctControl, BDR_SUNKEN, BF_RECT Or BF_MONO)
            Call FillRect(UserControl.hDC, m_uRctEdit, m_hBackBrush)
            Call DrawFrameControl(UserControl.hDC, m_uRctButton, DFC_SCROLL, DFCS_SCROLLCOMBOBOX Or DFCS_MONO Or DFCS_PUSHED * -(m_bButtonPressed) Or DFCS_INACTIVE * -(Not UserControl.Enabled))
        
        Case m_eStyle = [sThemed]
    
            If (UserControl.Enabled) Then
                Call pvDrawThemePart(CB_THEME, CP_BORDER, CBXS_NORMAL, m_uRctControl)
                Call FillRect(UserControl.hDC, m_uRctEdit, m_hBackBrush)
                If (m_bButtonPressed) Then
                    Call pvDrawThemePart(CB_THEME, CP_DROPDOWNBUTTON, CBXS_PRESSED, m_uRctButton)
                  Else
                    If (m_bButtonHot And m_hWndTrackbar = 0) Then
                        Call pvDrawThemePart(CB_THEME, CP_DROPDOWNBUTTON, CBXS_HOT, m_uRctButton)
                      Else
                        Call pvDrawThemePart(CB_THEME, CP_DROPDOWNBUTTON, CBXS_NORMAL, m_uRctButton)
                    End If
                End If
              Else
                Call pvDrawThemePart(CB_THEME, CP_BORDER, CBXS_DISABLED, m_uRctControl)
                Call FillRect(UserControl.hDC, m_uRctEdit, m_hBackBrush)
                Call pvDrawThemePart(CB_THEME, CP_DROPDOWNBUTTON, CBXS_DISABLED, m_uRctButton)
            End If
    End Select
    
    If (Ambient.UserMode = False) Then
        Call CopyRect(uRct, m_uRctEdit)
        Call InflateRect(uRct, -2, 0)
        Call DrawText(UserControl.hDC, Ambient.DisplayName, -1, uRct, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER)
    End If
    
    Call UserControl.Refresh
End Sub

Private Function pvPaintTrackbar( _
                 ByVal lParam As Long _
                 ) As Long
                 
  Dim uNMCD As NMCUSTOMDRAW

    Call CopyMemory(uNMCD, ByVal lParam, Len(uNMCD))
      
    With uNMCD
    
        Select Case .dwDrawStage
          
            Case CDDS_PREPAINT
            
                pvPaintTrackbar = CDRF_NOTIFYITEMDRAW
          
            Case CDDS_ITEMPREPAINT
                
                Select Case .dwItemSpec
                    
                    Case TBCD_TICS
                        
                        '-- Nothing to do here
                        
                    Case TBCD_THUMB
                        
                        '-- Paint button
                        Select Case m_eStyle
                            Case [sClassic]
                                Call DrawFrameControl(.hDC, .rc, 0, 0)
                            Case [sFlat]
                                Call DrawFrameControl(.hDC, .rc, 0, DFCS_FLAT)
                            Case [sFlatMono]
                                Call DrawFrameControl(.hDC, .rc, 0, DFCS_MONO)
                        End Select
                        
                        '-- Highlight it if selected
                        If (.uItemState And CDIS_SELECTED) Then
                            Call InflateRect(.rc, -EDGE_THICK, -EDGE_THICK)
                            Call SetTextColor(.hDC, pvTranslateColor(vb3DHighlight))
                            Call FillRect(.hDC, .rc, m_hPatternBrush)
                        End If
                        
                    Case TBCD_CHANNEL
                    
                        '-- Paint edge
                        Select Case m_eStyle
                            Case [sClassic]
                                Call DrawEdge(.hDC, .rc, BDR_SUNKEN, BF_RECT)
                            Case [sFlat]
                                Call DrawEdge(.hDC, .rc, BDR_SUNKEN, BF_RECT Or BF_FLAT)
                            Case [sFlatMono]
                                Call DrawEdge(.hDC, .rc, BDR_SUNKEN, BF_RECT Or BF_MONO)
                        End Select
                End Select
                
                pvPaintTrackbar = CDRF_SKIPDEFAULT
        End Select
    End With
End Function

Private Function pvDrawThemePart( _
                 ByVal sClass As String, _
                 ByVal lPart As Long, _
                 ByVal lState As Long, _
                 lpRect As RECT _
                 ) As Boolean
  
  Dim hTheme As Long
    
    On Error GoTo Catch
    
    hTheme = OpenThemeData(UserControl.hWnd, StrPtr(sClass))
    If (hTheme <> 0) Then
        pvDrawThemePart = (DrawThemeBackground(hTheme, UserControl.hDC, lPart, lState, lpRect, lpRect) = 0)
    End If
    
Catch:
    On Error GoTo 0
End Function

Private Function pvCreateBrushFromStdPicture( _
                 Image As StdPicture _
                 ) As Long

  Dim uBI       As BITMAP
  Dim uBIH      As BITMAPINFOHEADER
  Dim aBuffer() As Byte
    
  Dim lHDC      As Long
  Dim lhOldBmp  As Long
    
    '-- Valid source?
    If (GetObjectType(Image.Handle) = OBJ_BITMAP) Then
    
        '-- Get image info
        Call GetObject(Image.Handle, Len(uBI), uBI)
        
        '-- Prepare DIB header
        With uBIH
            .biSize = Len(uBIH)
            .biPlanes = 1
            .biBitCount = 24
            .biWidth = uBI.bmWidth
            .biHeight = uBI.bmHeight
            .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
        End With
        
        '-- Size byte-array
        ReDim aBuffer(1 To Len(uBIH) + uBIH.biSizeImage)
            
        '-- Create DIB brush
        lHDC = CreateCompatibleDC(0)
        If (lHDC <> 0) Then
        
            '-- Select our bitmap into the DC
            lhOldBmp = SelectObject(lHDC, Image.Handle)
                    
            '-- Set header bits
            Call CopyMemory(aBuffer(1), uBIH, Len(uBIH))
            '-- Set image bits
            Call GetDIBits(lHDC, Image.Handle, 0, uBI.bmHeight, aBuffer(Len(uBIH) + 1), uBIH, DIB_RGB_COLORS)
            
            '-- Clean up
            Call SelectObject(lHDC, lhOldBmp)
            Call DeleteDC(lHDC)
            
            '-- Finaly, create brush from packed DIB
            pvCreateBrushFromStdPicture = CreateDIBPatternBrushPt(aBuffer(1), DIB_RGB_COLORS)
        End If
    End If
End Function

Private Sub pvCreateBackBrush()
    
    '-- Destroy previous if any
    If (m_hBackBrush) Then
        Call DeleteObject(m_hBackBrush)
        m_hBackBrush = 0
    End If
    
    '-- Create solid/image brush
    Select Case m_eBackStyle
        Case [bsSolidColor]
            m_hBackBrush = CreateSolidBrush(pvTranslateColor(m_oleBackColor))
        Case [bsImage]
            If (m_oBackImage Is Nothing) Then
                m_hBackBrush = CreateSolidBrush(pvTranslateColor(m_oleBackColor))
              Else
                m_hBackBrush = pvCreateBrushFromStdPicture(m_oBackImage)
            End If
    End Select
End Sub

Private Sub pvCreatePatternBrush()

  Dim hBitmap          As Long
  Dim nPattern(1 To 8) As Integer
    
    '-- Brush pattern (8x8)
    nPattern(1) = &HAA
    nPattern(2) = &H55
    nPattern(3) = &HAA
    nPattern(4) = &H55
    nPattern(5) = &HAA
    nPattern(6) = &H55
    nPattern(7) = &HAA
    nPattern(8) = &H55
    
    '-- Create brush from bitmap
    hBitmap = CreateBitmap(8, 8, 1, 1, nPattern(1))
    m_hPatternBrush = CreatePatternBrush(hBitmap)
    Call DeleteObject(hBitmap)
End Sub

'----------------------------------------------------------------------------------------
' Controling range/value
'----------------------------------------------------------------------------------------

Private Sub pvCalcIntegerRangeValues()
    
    '-- Translate range not-integer values to
    '   trackbar integer ones
    m_lPrecisionFactor = (10 ^ m_eRangePrecision)
    m_lMin = m_snRangeMin * m_lPrecisionFactor
    m_lMax = m_snRangeMax * m_lPrecisionFactor
    m_lValue = m_lMin
End Sub

Private Sub pvValueInc()
    
    Select Case True
        Case m_lValue > m_lMax
            m_lValue = m_lMax
        Case m_lValue < m_lMin
            m_lValue = m_lMin
        Case m_lValue < m_lMax
            m_lValue = m_lValue + 1
    End Select
End Sub

Private Sub pvValueDec()
    
    Select Case True
        Case m_lValue > m_lMax
            m_lValue = m_lMax
        Case m_lValue < m_lMin
            m_lValue = m_lMin
        Case m_lValue > m_lMin
            m_lValue = m_lValue - 1
    End Select
End Sub

Private Sub pvUpdateEdit()
    
    On Error GoTo Catch
    
    If (m_oEdit.Text * m_lPrecisionFactor <> m_lValue) Then
        m_oEdit.Text = Format$(m_lValue / m_lPrecisionFactor, m_sPrecisionFormat(m_eRangePrecision))
        RaiseEvent Change
    End If

Catch:
    If (Err.Number) Then
        m_oEdit.Text = Format$(m_lValue / m_lPrecisionFactor, m_sPrecisionFormat(m_eRangePrecision))
        RaiseEvent Change
    End If
    On Error GoTo 0
End Sub

Private Sub pvSelectEditContents()

    With m_oEdit
        If (.SelLength <> Len(.Text)) Then
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
    End With
End Sub

'----------------------------------------------------------------------------------------
' Timing (hot button)
'----------------------------------------------------------------------------------------

Private Sub pvSetTimer(ByVal lTimerID As Long, ByVal ldT As Long)
    Call SetTimer(UserControl.hWnd, lTimerID, ldT, 0)
End Sub

Private Sub pvKillTimer(ByVal lTimerID As Long)
    Call KillTimer(UserControl.hWnd, lTimerID)
End Sub

'----------------------------------------------------------------------------------------
' Misc.
'----------------------------------------------------------------------------------------

Private Sub pvMakePoints( _
            ByVal lParam As Long, _
            x As Long, _
            y As Long _
            )

    x = (lParam And &HFFFF&)
    y = (lParam And &HFFFF0000) \ &H10000
End Sub

Private Function pvTranslateColor( _
                 ByVal clr As OLE_COLOR _
                 ) As Long
    
    Call OleTranslateColor(clr, 0, pvTranslateColor)
End Function



'========================================================================================
' UserControl persistent properties
'========================================================================================

Private Sub UserControl_InitProperties()
    
    '-- Initialize 'm_oEdit' and 'm_oFont' objects
    Call pvInitializeObjects
    
    '-- Set inherently-stored properties
    Set m_oFont = Ambient.Font
        Set m_oEdit.Font = m_oFont
        Set UserControl.Font = m_oFont
    Let m_oEdit.ForeColor = FORECOLOR_DEF
        Let UserControl.ForeColor = FORECOLOR_DEF
    
    '-- Set 'memory' properties
    Let m_oleBackColor = BACKCOLOR_DEF
    Set m_oBackImage = Nothing
    Let m_eBackStyle = BACKSTYLE_DEF
    Let m_bEditDelayedUpdate = EDITDELAYEDUPDATE_DEF
    Let m_lEditDelay = EDITDELAY_DEF
    Let m_snRangeMax = RANGEMAX_DEF
    Let m_snRangeMin = RANGEMIN_DEF
    Let m_eRangePrecision = RANGEPRECISION_DEF
    Let m_eStyle = STYLE_DEF
    Let m_lTrackbarHeight = TRACKBARHEIGHT_DEF
    Let m_eTrackbarPosition = TRACKBARPOSITION_DEF
    Let m_lTrackbarWidth = TRACKBARWIDTH_DEF
    
    '-- Initialize all
    Call pvCalcIntegerRangeValues
    Call pvCreateBackBrush
    Call pvCreatePatternBrush
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    '-- Initialize 'm_oEdit' and 'm_oFont' objects
    Call pvInitializeObjects
  
    With PropBag
        
        '-- Read inherently-stored properties
        Set m_oFont = .ReadProperty("Font", Ambient.Font)
            Set UserControl.Font = m_oFont
            Set m_oEdit.Font = m_oFont
        Let UserControl.ForeColor = .ReadProperty("ForeColor", FORECOLOR_DEF)
            Let m_oEdit.ForeColor = UserControl.ForeColor
        Let UserControl.Enabled = .ReadProperty("Enabled", ENABLED_DEF)
            Let m_oEdit.Enabled = UserControl.Enabled
        Let m_oEdit.Locked = .ReadProperty("Locked", LOCKED_DEF)
        
        '-- Read 'memory' properties
        Let m_oleBackColor = .ReadProperty("BackColor", BACKCOLOR_DEF)
        Set m_oBackImage = .ReadProperty("BackImage", Nothing)
        Let m_eBackStyle = .ReadProperty("BackStyle", BACKSTYLE_DEF)
        Let m_bEditDelayedUpdate = .ReadProperty("EditDelayedUpdate", EDITDELAYEDUPDATE_DEF)
        Let m_lEditDelay = .ReadProperty("EditDelay", EDITDELAY_DEF)
        Let m_snRangeMax = .ReadProperty("RangeMax", RANGEMAX_DEF)
        Let m_snRangeMin = .ReadProperty("RangeMin", RANGEMIN_DEF)
        Let m_eRangePrecision = .ReadProperty("RangePrecision", RANGEPRECISION_DEF)
        Let m_eStyle = .ReadProperty("Style", STYLE_DEF)
        Let m_lTrackbarHeight = .ReadProperty("TrackbarHeight", TRACKBARHEIGHT_DEF)
        Let m_eTrackbarPosition = .ReadProperty("TrackbarPosition", TRACKBARPOSITION_DEF)
        Let m_lTrackbarWidth = .ReadProperty("TrackbarWidth", TRACKBARWIDTH_DEF)
        
        '-- Initialize all
        Call pvCalcIntegerRangeValues
        Call pvCreateBackBrush
        Call pvCreatePatternBrush
    End With
    
    '-- Only on run-time
    If (Ambient.UserMode) Then
    
        '-- Check OS and Luna theme
        Call pvCheckEnvironment
            
        '-- Update and show Edit now
        Call pvUpdateEdit
        Let m_oEdit.Visible = True
        
        '-- Subclass Edit window
        Call Subclass_Start(m_oEdit.hWnd)
        Call Subclass_AddMsg(m_oEdit.hWnd, WM_SETFOCUS)
        Call Subclass_AddMsg(m_oEdit.hWnd, WM_KILLFOCUS)
        
        '-- Subclass UserControl window
        Call Subclass_Start(UserControl.hWnd)
        Call Subclass_AddMsg(UserControl.hWnd, WM_NOTIFY, [MSG_BEFORE])
        Call Subclass_AddMsg(UserControl.hWnd, WM_HSCROLL)
        Call Subclass_AddMsg(UserControl.hWnd, WM_MOUSEWHEEL)
        Call Subclass_AddMsg(UserControl.hWnd, WM_LBUTTONDOWN)
        Call Subclass_AddMsg(UserControl.hWnd, WM_LBUTTONUP)
        Call Subclass_AddMsg(UserControl.hWnd, WM_MOUSEMOVE)
        Call Subclass_AddMsg(UserControl.hWnd, WM_TIMER)
        Call Subclass_AddMsg(UserControl.hWnd, WM_CTLCOLOREDIT)
        Call Subclass_AddMsg(UserControl.hWnd, WM_CTLCOLORSTATIC)
        Call Subclass_AddMsg(UserControl.hWnd, WM_SYSCOLORCHANGE)
        If (m_bIsXP) Then
            Call Subclass_AddMsg(UserControl.hWnd, WM_THEMECHANGED)
        End If
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("BackColor", m_oleBackColor, BACKCOLOR_DEF)
        Call .WriteProperty("BackImage", m_oBackImage, Nothing)
        Call .WriteProperty("BackStyle", m_eBackStyle, BACKSTYLE_DEF)
        Call .WriteProperty("EditDelayedUpdate", m_bEditDelayedUpdate, EDITDELAYEDUPDATE_DEF)
        Call .WriteProperty("EditDelay", m_lEditDelay, EDITDELAY_DEF)
        Call .WriteProperty("Enabled", UserControl.Enabled, ENABLED_DEF)
        Call .WriteProperty("Font", m_oFont, Ambient.Font)
        Call .WriteProperty("ForeColor", m_oEdit.ForeColor, FORECOLOR_DEF)
        Call .WriteProperty("Locked", m_oEdit.Locked, LOCKED_DEF)
        Call .WriteProperty("RangeMax", m_snRangeMax, RANGEMAX_DEF)
        Call .WriteProperty("RangeMin", m_snRangeMin, RANGEMIN_DEF)
        Call .WriteProperty("RangePrecision", m_eRangePrecision, RANGEPRECISION_DEF)
        Call .WriteProperty("Style", m_eStyle, STYLE_DEF)
        Call .WriteProperty("TrackbarHeight", m_lTrackbarHeight, TRACKBARHEIGHT_DEF)
        Call .WriteProperty("TrackbarPosition", m_eTrackbarPosition, TRACKBARPOSITION_DEF)
        Call .WriteProperty("TrackbarWidth", m_lTrackbarWidth, TRACKBARWIDTH_DEF)
    End With
End Sub

'//

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_oleBackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_oleBackColor = New_BackColor
    Call pvCreateBackBrush
    Call pvPaintCombo
    Call InvalidateRect(m_oEdit.hWnd, ByVal 0, 0)
End Property

Public Property Get BackImage() As StdPicture
    Set BackImage = m_oBackImage
End Property

Public Property Set BackImage(ByVal New_BackImage As StdPicture)
    Set m_oBackImage = New_BackImage
    Call pvCreateBackBrush
    Call pvPaintCombo
    Call InvalidateRect(m_oEdit.hWnd, ByVal 0, 0)
End Property

Public Property Get BackStyle() As ctBackStyleCts
    BackStyle = m_eBackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As ctBackStyleCts)
    m_eBackStyle = New_BackStyle
    Call pvCreateBackBrush
    Call pvPaintCombo
    Call InvalidateRect(m_oEdit.hWnd, ByVal 0, 0)
End Property

Public Property Get EditDelayedUpdate() As Boolean
    EditDelayedUpdate = m_bEditDelayedUpdate
End Property

Public Property Let EditDelayedUpdate(ByVal New_EditDelayedUpdate As Boolean)
    m_bEditDelayedUpdate = New_EditDelayedUpdate
End Property

Public Property Get EditDelay() As Long
    EditDelay = m_lEditDelay
End Property

Public Property Let EditDelay(ByVal New_EditDelay As Long)
    If (New_EditDelay < TIMERMINDT_EDIT) Then
        New_EditDelay = TIMERMINDT_EDIT
    End If
    m_lEditDelay = New_EditDelay
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enable As Boolean)
    UserControl.Enabled = New_Enable
    m_oEdit.Enabled = New_Enable
    Call pvPaintCombo
    Call InvalidateRect(m_oEdit.hWnd, ByVal 0, 0)
End Property

Public Property Get Font() As StdFont
    Set Font = m_oFont
End Property

Public Property Set Font(ByVal New_Font As StdFont)
    With m_oFont
        .Charset = New_Font.Charset
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
        .Weight = New_Font.Weight
    End With
End Property

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    Set m_oEdit.Font = m_oFont
    Set UserControl.Font = m_oFont
    Call pvResizeControl
    Call InvalidateRect(m_oEdit.hWnd, ByVal 0, 0)
End Sub

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_oEdit.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_oEdit.ForeColor = New_ForeColor
    UserControl.ForeColor = New_ForeColor
    Call pvPaintCombo
    Call InvalidateRect(m_oEdit.hWnd, ByVal 0, 0)
End Property

Public Property Get Locked() As Boolean
    Locked = m_oEdit.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    m_oEdit.Locked = New_Locked
    Call InvalidateRect(m_oEdit.hWnd, ByVal 0, 0)
End Property

Public Property Get RangeMin() As Single
    RangeMin = m_snRangeMin
End Property

Public Property Let RangeMin(ByVal New_RangeMin As Single)
    If (New_RangeMin > m_snRangeMax) Then
        New_RangeMin = m_snRangeMax
    End If
    If (pvValidateRange(New_RangeMin, m_snRangeMax, m_eRangePrecision)) Then
        m_snRangeMin = New_RangeMin
        Call pvCalcIntegerRangeValues
        Call InvalidateRect(m_oEdit.hWnd, ByVal 0, 0)
      Else
        Call VBA.Err.Raise(ERR_OVERFLOW)
    End If
End Property

Public Property Get RangeMax() As Single
    RangeMax = m_snRangeMax
End Property

Public Property Let RangeMax(ByVal New_RangeMax As Single)
    If (New_RangeMax < m_snRangeMin) Then
        New_RangeMax = m_snRangeMin
    End If
    If (pvValidateRange(m_snRangeMin, New_RangeMax, m_eRangePrecision)) Then
        m_snRangeMax = New_RangeMax
        Call pvCalcIntegerRangeValues
        Call InvalidateRect(m_oEdit.hWnd, ByVal 0, 0)
      Else
        Call VBA.Err.Raise(ERR_OVERFLOW)
    End If
End Property

Public Property Get RangePrecision() As ctRangePrecisionCts
    RangePrecision = m_eRangePrecision
End Property

Public Property Let RangePrecision(ByVal New_RangePrecision As ctRangePrecisionCts)
    If (New_RangePrecision < [rpInteger]) Then
        New_RangePrecision = [rpInteger]
    ElseIf New_RangePrecision > [rpThousandth] Then
        New_RangePrecision = [rpThousandth]
    End If
    If (pvValidateRange(m_snRangeMin, m_snRangeMax, New_RangePrecision)) Then
        m_eRangePrecision = New_RangePrecision
        Call pvCalcIntegerRangeValues
        Call InvalidateRect(m_oEdit.hWnd, ByVal 0, 0)
      Else
        Call VBA.Err.Raise(ERR_OVERFLOW)
    End If
End Property

Public Property Get Style() As ctStyleCts
    Style = m_eStyle
End Property

Public Property Let Style(ByVal New_Style As ctStyleCts)
    m_eStyle = New_Style
    Call pvResizeControl
    Call InvalidateRect(m_oEdit.hWnd, ByVal 0, 0)
End Property

Public Property Get TrackbarPosition() As ctTrackbarPositionCts
    TrackbarPosition = m_eTrackbarPosition
End Property

Public Property Let TrackbarPosition(ByVal New_TrackbarPosition As ctTrackbarPositionCts)
    m_eTrackbarPosition = New_TrackbarPosition
End Property

Public Property Get TrackbarWidth() As Long
    TrackbarWidth = m_lTrackbarWidth
End Property

Public Property Let TrackbarWidth(ByVal New_TrackbarWidth As Long)
    If (New_TrackbarWidth < TRACKBAR_WIDTH_MIN) Then
        New_TrackbarWidth = TRACKBAR_WIDTH_MIN
    End If
    m_lTrackbarWidth = New_TrackbarWidth
End Property

Public Property Get TrackbarHeight() As Long
    TrackbarHeight = m_lTrackbarHeight
End Property

Public Property Let TrackbarHeight(ByVal New_TrackbarHeight As Long)
    If (New_TrackbarHeight < TRACKBAR_HEIGHT_MIN) Then
        New_TrackbarHeight = TRACKBAR_HEIGHT_MIN
    End If
    m_lTrackbarHeight = New_TrackbarHeight
End Property

'// Runtime read only

Public Property Get PrecisionFormat() As String
    PrecisionFormat = m_sPrecisionFormat(m_eRangePrecision)
End Property

Public Property Get Value() As Single
    Value = m_lValue / m_lPrecisionFactor
End Property

'// Runtime read-write only

Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "400"
    Text = m_oEdit.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    m_oEdit.Text = New_Text
End Property

'// Runtime read only (OS related)

Public Property Get IsXP() As Boolean
    IsXP = m_bIsXP
End Property

Public Property Get IsThemed() As Boolean
    IsThemed = m_bIsLuna
End Property



'========================================================================================
'Subclass routines below here - The programmer may call any of the following Subclass_??? routines
'========================================================================================

Private Sub Subclass_AddMsg(ByVal lhWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
  
    With sc_aSubData(zIdx(lhWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

Private Function Subclass_Start(ByVal lhWnd As Long) As Long

  Const CODE_LEN              As Long = 202
  Const FUNC_CWP              As String = "CallWindowProcA"
  Const FUNC_EBM              As String = "EbMode"
  Const FUNC_SWL              As String = "SetWindowLongA"
  Const MOD_USER              As String = "user32"
  Const MOD_VBA5              As String = "vba5"
  Const MOD_VBA6              As String = "vba6"
  Const PATCH_01              As Long = 18
  Const PATCH_02              As Long = 68
  Const PATCH_03              As Long = 78
  Const PATCH_06              As Long = 116
  Const PATCH_07              As Long = 121
  Const PATCH_0A              As Long = 186
  Static aBuf(1 To CODE_LEN)  As Byte
  Static pCWP                 As Long
  Static pEbMode              As Long
  Static pSWL                 As Long
  Dim i                       As Long
  Dim j                       As Long
  Dim nSubIdx                 As Long
  Dim sHex                    As String
  
    If (aBuf(1) = 0) Then
  
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))
            i = i + 2
        Loop
    
        If (Subclass_InIDE) Then
            aBuf(16) = &H90
            aBuf(17) = &H90
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)
            If (pEbMode = 0) Then
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)
            End If
        End If
    
        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)
        ReDim sc_aSubData(0 To 0) As tSubData
      Else
        nSubIdx = zIdx(lhWnd, True)
        If (nSubIdx = -1) Then
            nSubIdx = UBound(sc_aSubData()) + 1
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
        End If
    
        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hWnd = lhWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)
        Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))
    End With
End Function

Private Sub Subclass_Stop(ByVal lhWnd As Long)
  
    With sc_aSubData(zIdx(lhWnd))
        Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)
        Call zPatchVal(.nAddrSub, PATCH_05, 0)
        Call zPatchVal(.nAddrSub, PATCH_09, 0)
        Call GlobalFree(.nAddrSub)
        .hWnd = 0
        .nMsgCntB = 0
        .nMsgCntA = 0
        Erase .aMsgTblB()
        Erase .aMsgTblA()
    End With
End Sub

Private Sub Subclass_StopAll()
  
  Dim i As Long
  
    i = UBound(sc_aSubData())
    Do While i >= 0
        With sc_aSubData(i)
            If (.hWnd <> 0) Then
                Call Subclass_Stop(.hWnd)
            End If
        End With
        i = i - 1
    Loop
End Sub

'----------------------------------------------------------------------------------------
'These z??? routines are exclusively called by the Subclass_??? routines.
'----------------------------------------------------------------------------------------

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  
  Dim nEntry  As Long
  Dim nOff1   As Long
  Dim nOff2   As Long
  
    If (uMsg = ALL_MESSAGES) Then
        nMsgCnt = ALL_MESSAGES
      Else
        Do While nEntry < nMsgCnt
            nEntry = nEntry + 1
            If (aMsgTbl(nEntry) = 0) Then
                aMsgTbl(nEntry) = uMsg
                Exit Sub
            ElseIf (aMsgTbl(nEntry) = uMsg) Then
                Exit Sub
            End If
        Loop

        nMsgCnt = nMsgCnt + 1
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long
        aMsgTbl(nMsgCnt) = uMsg
    End If

    If (When = eMsgWhen.MSG_BEFORE) Then
        nOff1 = PATCH_04
        nOff2 = PATCH_05
      Else
        nOff1 = PATCH_08
        nOff2 = PATCH_09
    End If

    If (uMsg <> ALL_MESSAGES) Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc
End Function

Private Function zIdx(ByVal lhWnd As Long, Optional ByVal bAdd As Boolean = False) As Long

    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0
        With sc_aSubData(zIdx)
            If (.hWnd = lhWnd) Then
                If Not bAdd Then
                    Exit Function
                End If
            ElseIf (.hWnd = 0) Then
                If bAdd Then
                    Exit Function
                End If
            End If
        End With
        zIdx = zIdx - 1
    Loop
  
    If (Not bAdd) Then
        Debug.Assert False
    End If
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function

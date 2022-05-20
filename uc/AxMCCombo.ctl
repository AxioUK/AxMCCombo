VERSION 5.00
Begin VB.UserControl axMCCombo 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   166
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   209
   ToolboxBitmap   =   "AxMCCombo.ctx":0000
   Begin VB.Timer tmrRelease 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2805
      Top             =   600
   End
   Begin VB.Timer TmrMouseOver 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2820
      Top             =   135
   End
   Begin VB.TextBox oText 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   135
      Width           =   2505
   End
   Begin VB.PictureBox pList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1620
      Left            =   45
      ScaleHeight     =   108
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   186
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   2790
   End
End
Attribute VB_Name = "axMCCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Readme

Private Type POINTAPI
    X   As Long
    Y   As Long
End Type

Private Type Rect
    L   As Long
    T   As Long
    r   As Long
    B   As Long
End Type

Private Type RectF
  Left As Long
  Top As Long
  Width As Long
  Height As Long
End Type

Private Type GUID
  Data1   As Long
  Data2   As Integer
  Data3   As Integer
  Data4(7) As Byte
End Type

Private Const HWND_TOPMOST       As Long = -1
Private Const HWND_NOTOPMOST     As Long = -2
Private Const SWP_NOSIZE         As Long = &H1
Private Const SWP_NOMOVE         As Long = &H2
Private Const SWP_NOACTIVATE     As Long = &H10
Private Const SWP_SHOWWINDOW     As Long = &H40
Private Const SWP_HIDEWINDOW     As Long = &H80

Private Const EVENT_TIMEOUT         As Long = 500

Private Const WrapModeTileFlipXY As Long = &H3
Private Const UnitPixel          As Long = &H2&
Private Const LOGPIXELSX         As Long = 88
Private Const LOGPIXELSY         As Long = 90
Private Const SmoothingModeAntiAlias As Long = 4

Private Const C_NULL_RESULT      As Long = -1

'/Window
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As Rect) As Long
Private Declare Function GetClassLong Lib "user32.dll" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32.dll" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetFocusEx Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
'---

Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
  cbSize                             As Long
  dwFlags                            As TRACKMOUSEEVENT_FLAGS
  hwndTrack                          As Long
  dwHoverTime                        As Long
End Type

Private bTrack                       As Boolean
Private bTrackUser32                 As Boolean
Private mInCtrl                      As Boolean
Private mInFocus                     As Boolean

Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32.dll" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
'---
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal Color As Long, ByVal hPal As Long, ByRef RGBResult As Long) As Long

'/WindowMessages
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'/Theme
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Rect, pClipRect As Any) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long

'?Border
Private Declare Function ExcludeClipRect Lib "gdi32.dll" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long

'/Draw
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function OleTranslateColor2 Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'''''----------------------------------------------------------
Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipCreateSolidFill Lib "GdiPlus.dll" (ByVal argb As Long, ByRef brush As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathRectangle Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipAddPathArc Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal brush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal token As Long)
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal hGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
'-------------------------------------------------------------------
'Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DrawTextW Lib "user32.dll" (ByVal hDC As Long, lpStr As Long, ByVal nCount As Long, ByRef lpRect As Rect, ByVal wFormat As Long) As Long
'-------------------------------------------------------------------
Private Declare Function PtInRect Lib "user32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
'-------------------------------------------------------------------

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type tHeader
    Text    As String
    Image   As Long
    Width   As Long
    Aling   As Integer
    IAlign  As Integer
End Type

Private Type tSubItem
    Text    As String
    Icon    As Long
End Type

Private Type tItem
    Item()  As tSubItem
    Data    As Long
    Tag     As String
End Type

Public Enum eFlatSide
  rUp
  rBottom
  rLeft
  rRight
End Enum

Public Enum JComboStyle
    axJComboBox = 0
    axJListCombo = 1
End Enum

''EVENTS--------------------------------------------
Public Event ItemClick(Item As Long)
Public Event ListIndexChanged(ByVal Item As Integer)
Public Event Change()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

''VARIABLES-----------------------------------------
Private cSubClass As c_SubClass
Private WithEvents VScrollBar As c_ScrollBars
Attribute VScrollBar.VB_VarHelpID = -1

Private m_ItemH         As Long
Private m_HeaderH       As Long
Private m_GridLineColor As Long
Private m_GridStyle     As Integer
Private m_Striped       As Boolean
Private m_Header        As Boolean
Private m_DrawEmpty     As Boolean
Private m_StripedColor  As Long
Private m_ForeColor     As OLE_COLOR
Private m_SelColor      As OLE_COLOR
Private m_ForeSel       As OLE_COLOR
Private m_BorderColor   As OLE_COLOR
Private m_BackColor     As OLE_COLOR
Private m_ButtonColorPress   As OLE_COLOR
Private m_BackColorParent    As OLE_COLOR
Private m_VisibleRows   As Long
Private m_DropW         As Long

Private mScrollTick     As Long
Private pmTrack(3)      As Long
Private m_PhWnd         As Long
Private m_hWnd          As Long
Private m_cols()        As tHeader
Private m_items()       As tItem
Private m_Iml           As Long

Private m_SelRow        As Long
Private m_bTrack        As Boolean
Private m_img           As POINTAPI
Private m_GridW         As Long
Private m_RowH          As Long
Private t_Row           As Long

Private e_Scale         As Long
Private lnScale         As Long
Private m_Visible       As Boolean
Private Expanded        As Boolean
Private m_ColumnInBox   As Integer
Private bInFocus        As Boolean
Private bFocus          As Boolean
Private m_ComboStyle    As JComboStyle
Private isVisible       As Boolean

Private gdipToken     As Long
Private sOldColor     As Long
Private m_Text        As String

Private TextH         As Long
Private m_Font        As StdFont
Private RctButton     As RectF
Private m_IconFont    As StdFont
Private m_IconCharCode    As Long
Private m_IconForeColor   As Long
Private m_PadY            As Long
Private m_PadX            As Long
Private m_MouseOver   As Boolean

Private m_MultiLine As Boolean
Private mEnabled  As Boolean
Private m_BorderWidth As Long
Private m_CornerRound As Long

Public Sub AddColumn(ByVal Text As String, Optional ByVal Width As Long = 80, Optional ByVal Alignment As AlignmentConstants)
Dim L       As Long
Dim I       As Long
    
    Width = Width * e_Scale
    L = ColumnCount
    
    ReDim Preserve m_cols(L)
    With m_cols(L)
        .Text = Text
        .Width = Width
        .Aling = Alignment
    End With
    m_GridW = m_GridW + Width
End Sub

Public Function AddItem(ByVal Text As String, Optional ByVal IconIndex As Long = -1, Optional ByVal ItemData As Long, Optional ByVal ItemTag As String = "") As Long
On Local Error Resume Next
Dim L   As Long
Dim I   As Long

    L = ItemCount
    ReDim Preserve m_items(L)
    
    With m_items(L)
        ReDim .Item(ColumnCount - 1)
        .Item(0).Text = Text
        .Item(0).Icon = IconIndex
        .Data = ItemData
        .Tag = ItemTag
        
        For I = 1 To ColumnCount - 1: .Item(I).Icon = -1: Next

    End With
    AddItem = L
    UpdateScrollV
End Function

Public Sub ClearItems()
    Erase m_items
    m_SelRow = -1
    UpdateScrollV
End Sub

Public Sub ColWidthAutoSize(Optional ByVal lCol As Long = C_NULL_RESULT)
Dim lngC As Long
  
   If lCol = C_NULL_RESULT Then
      For lngC = 0 To UBound(m_cols)
         Call ColSizing(lngC)
      Next lngC
   Else
      Call ColSizing(lCol)
   End If
      
End Sub

Public Function GetWindowsDPI() As Double
    Dim hDC As Long, lPx  As Double
    hDC = GetDC(0)
    lPx = CDbl(GetDeviceCaps(hDC, LOGPIXELSX))
    ReleaseDC 0, hDC

    If (lPx = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = lPx / 96#
    End If
End Function


'*1
Public Sub ReDrawControl(Optional BorderOpacity As Long = 50)
With UserControl
    .Cls
    .BackColor = m_BackColorParent
    oText.BackColor = m_BackColor
    oText.ForeColor = m_ForeColor
    DrawBorder .hDC, m_BorderColor, 90, BorderOpacity, 0, 0, .ScaleWidth - 30, .ScaleHeight - 1, rRight
    DrawButton m_BorderColor, BorderOpacity
    AddIconChar
    
    .BackStyle = 0
    .MaskColor = .BackColor
    Set .MaskPicture = .Image

    .Refresh
End With
End Sub

Public Sub RemoveItem(ByVal Index As Long)
On Local Error Resume Next
Dim j As Integer

    If ItemCount = 0 Or Index > ItemCount - 1 Or ItemCount < 0 Or Index < 0 Then Exit Sub
    
    If ItemCount > 1 Then
         For j = Index To UBound(m_items) - 1
            LSet m_items(j) = m_items(j + 1)
         Next
        ReDim Preserve m_items(UBound(m_items) - 1)
    Else
        Erase m_items
    End If
    
    UpdateScrollV
    If m_SelRow <> -1 Then
        If m_SelRow = Index Then m_SelRow = -1
        If m_SelRow > Index Then m_SelRow = m_SelRow - 1
    End If
    DrawGrid
End Sub

Public Function RGBtoARGB(ByVal RGBColor As Long, Optional ByVal Opacity As Long = 100) As Long

    If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
    RGBtoARGB = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
    Opacity = CByte((Abs(Opacity) / 100) * 255)
    If Opacity < 128 Then
        If Opacity < 0& Then Opacity = 0&
        RGBtoARGB = RGBtoARGB Or Opacity * &H1000000
    Else
        If Opacity > 255& Then Opacity = 255&
        RGBtoARGB = RGBtoARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
    End If
    
End Function

Public Sub ShowList(ByVal Visible As Boolean)
Dim lW  As Long
Dim lH  As Long
Dim lT  As Long
Dim Rct As Rect
Dim Rgn As Long
'Dim Rctf As RECTF
'Dim PT  As POINTAPI
       
    If Visible Then
        
        GetWindowRect UserControl.hwnd, Rct
        lW = IIf(m_DropW, m_DropW, Rct.r - Rct.L)
        lH = ((m_VisibleRows * m_RowH) + lHeaderH)

        SetParent pList.hwnd, 0
               
        SetWindowPos pList.hwnd, HWND_TOPMOST, Rct.L, Rct.B, lW, lH, SWP_NOACTIVATE Or SWP_SHOWWINDOW
        Call SetWindowLong(pList.hwnd, -8, UserControl.hwnd)
        
        UpdateScrollV
        Call DrawGrid
                
        Rgn = CreateRoundRectRgn(0, 0, (pList.Width / Screen.TwipsPerPixelX), (pList.Height / Screen.TwipsPerPixelY), m_CornerRound, m_CornerRound)
        SetWindowRgn pList.hwnd, Rgn, True
        DeleteObject Rgn

        pList.Visible = True
        SetFocusEx UserControl.hwnd
        SetFocusEx pList.hwnd
        m_Visible = True
        'SetTimer True
        
    Else
        pList.Visible = False
        SetParent pList.hwnd, UserControl.hwnd
        m_Visible = False
        SetTimer False
    End If
  
End Sub

'*3
Private Function AddIconChar()
Dim Rct As Rect
Dim pFont       As IFont
Dim lFontOld    As Long
    
On Error GoTo ErrF
With UserControl
  .AutoRedraw = True
  Set pFont = IconFont
  lFontOld = SelectObject(.hDC, pFont.hFont)

  If m_MouseOver Then
    .ForeColor = m_ButtonColorPress
  Else
    .ForeColor = m_IconForeColor
  End If
  
  Rct.L = RctButton.Left + (IconFont.Size / 2) + m_PadX
  Rct.T = RctButton.Top + (IconFont.Size / 2) + m_PadY
  Rct.r = RctButton.Left + RctButton.Width
  Rct.B = RctButton.Top + RctButton.Height
  
  DrawTextW .hDC, IconCharCode, 1, Rct, 0
  
  Call SelectObject(.hDC, lFontOld)
  
ErrF:
  Set pFont = Nothing
End With
End Function

Private Sub ColSizing(ByVal lCol As Long)
Dim lngLW As Long, lngCW As Long, lRow As Long
Dim strTemp As String

For lRow = 0 To UBound(m_items) - 1
  strTemp = m_items(lRow).Item(lCol).Text
  lngLW = UserControl.TextWidth(strTemp)
  If lngCW < lngLW Then lngCW = lngLW
Next lRow

  m_cols(lCol).Width = lngCW + Screen.TwipsPerPixelX
End Sub

Private Sub DrawBack(lpDC As Long, Color As Long, Rct As Rect)
Dim hBrush  As Long

    hBrush = CreateSolidBrush(Color)
    Call FillRect(lpDC, Rct, hBrush)
    Call DeleteObject(hBrush)

End Sub

Private Sub DrawBorder(lpDC As Long, cBorderColor As Long, BackOpacity As Long, BorderOpacity As Long, X As Long, Y As Long, w As Long, h As Long, fSide As eFlatSide)
Dim Rct As RectF
'Dim mPath As Long, hPen As Long
'Dim hGraphics As Long

Rct.Left = X
Rct.Top = Y
Rct.Width = w
Rct.Height = h
GDIpRoundBox lpDC, Rct, RGBtoARGB(m_BackColor, BackOpacity), m_BorderWidth, RGBtoARGB(cBorderColor, BorderOpacity), m_CornerRound, fSide

'GdipCreateFromHDC hDC, hGraphics
'GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
'mPath = GDIpCreateRoundPath(X, Y, w, h, m_CornerRound)
'GdipCreatePen1 RGBtoARGB(m_BackColor, BackOpacity), m_BorderWidth * e_Scale, &H2, hPen
'GdipDrawPath hGraphics, hPen, mPath
'GdipDeletePen hPen
'GdipDeleteGraphics hGraphics

End Sub

Private Sub DrawButton(BorderColor As OLE_COLOR, Opacity As Long)

RctButton.Left = UserControl.ScaleWidth - 30
RctButton.Top = 0
RctButton.Width = 29
RctButton.Height = UserControl.ScaleHeight - 1

GDIpRoundBox UserControl.hDC, RctButton, RGBtoARGB(m_BackColor, 90), m_BorderWidth, RGBtoARGB(BorderColor, Opacity), m_CornerRound, rLeft
End Sub

Private Sub DrawGrid()
On Local Error Resume Next
Dim lCol    As Long
Dim lRow    As Long
Dim ly      As Long
Dim lx      As Long
Dim lColW   As Long
Dim dvc     As Long
Dim iRct    As Rect
Dim tRct    As Rect
Dim lPx     As Long
Dim lPx2    As Long

    pList.Cls
    
    lCol = 0
    lRow = 0

    ly = -GetScroll(1)
    dvc = pList.hDC
    
    ly = ly + lHeaderH
    'Debug.Print "DrawGrid ly=" & ly
    
    Do While lRow <= ItemCount - 1 And ly < pList.ScaleHeight
        
        If ly + m_RowH > 0 Then '?Visible
            
            SetRect iRct, 0, ly, pList.ScaleWidth, ly + m_ItemH
            If m_Striped And lRow Mod 2 Then _
                DrawBack dvc, SysColor(m_StripedColor), iRct
            
            '\ Seleccion
            lPx2 = m_GridW - lnScale
            If lPx2 > pList.ScaleWidth + (8 * lnScale) Then lPx2 = pList.ScaleWidth + (8 * lnScale)
            If lRow = m_SelRow Then
                DrawSelection dvc, 0, ly, lPx2, m_ItemH, 1
            ElseIf lRow = t_Row Then
                DrawSelection dvc, 0, ly, lPx2, m_ItemH, 1
            End If
                
            lPx2 = lnScale \ 2
            '?GridLines 0N,1H,2V,3B -> Horizontal
            If m_GridStyle = 1 Or m_GridStyle = 3 Then _
                DrawLine dvc, lx, ly + m_ItemH + lPx2, pList.ScaleWidth, ly + m_ItemH + lPx2, m_GridLineColor
            
             Do While lCol < ColumnCount And lx < pList.ScaleWidth
             
                If ColumnCount = 1 Then
                  lColW = pList.ScaleWidth - 2
                Else
                  lColW = m_cols(lCol).Width
                End If
                
                If m_GridStyle = 2 Or m_GridStyle = 3 Then lColW = lColW - lnScale
                
                    SetRect iRct, lx, ly, lx + lColW, ly + m_ItemH
                   
                    '?GridLines 0N,1H,2V,3B - > Vertical
                     If m_GridStyle = 2 Or m_GridStyle = 3 Then _
                       DrawLine dvc, lx + lColW + lPx2, ly, lx + lColW + lPx2, ly + m_ItemH, m_GridLineColor
                       
                    If Trim(m_items(lRow).Item(lCol).Text) <> vbNullString Then
                        
                        SetRect tRct, lx + 4 + lPx, ly, lx + lColW - 3, ly + m_ItemH
                        If tRct.r < tRct.L Then tRct.r = tRct.L
                        If tRct.r > tRct.L Then _
                        DrawText dvc, m_items(lRow).Item(lCol).Text, Len(m_items(lRow).Item(lCol).Text), tRct, GetTextFlag(lCol)
                    
                    End If
eDrawNext:
                lx = lx + m_cols(lCol).Width
                lCol = lCol + 1
                
             Loop
            '?Reset to Scroll Position
            lCol = 0
            lx = 0
        End If
        
        ly = ly + m_RowH
        lRow = lRow + 1
    Loop
    Call DrawHeader
    DrawBorder pList.hDC, m_BorderColor, 0, 90, 0, 0, pList.ScaleWidth - 1, pList.ScaleHeight - 2, rUp
End Sub

Private Function DrawHeader()
Dim uTheme  As Long
Dim uRct    As Rect
Dim Col    As Long
Dim lx      As Long
Dim lW      As Long

    uTheme = OpenThemeData(pList.hwnd, StrPtr("Header"))
    If uTheme = 0 Then Exit Function
    
    SetRect uRct, 0, 0, pList.ScaleWidth, lHeaderH
    Call DrawThemeBackground(uTheme, pList.hDC, 0, 0&, uRct, ByVal 0&)
    
    Do While Col < ColumnCount And lx < pList.ScaleWidth
    
        lW = m_cols(Col).Width
        SetRect uRct, lx, 0, lx + lW, lHeaderH
        
        Call DrawThemeBackground(uTheme, pList.hDC, 1, 1, uRct, ByVal 0&)
        
        uRct.r = uRct.r - (10 * 1)
        OffsetRect uRct, 5 * 1, 0
        
        DrawText pList.hDC, m_cols(Col).Text, Len(m_cols(Col).Text), uRct, GetTextFlag(Col)
        
        lx = lx + m_cols(Col).Width
        Col = Col + 1
        
    Loop
    
End Function

Private Sub DrawLine(lpDC As Long, X As Long, Y As Long, X2 As Long, Y2 As Long, Color As Long)
Dim PT      As POINTAPI
Dim hPen    As Long
Dim hPenOld As Long

    hPen = CreatePen(0, lnScale, Color)
    hPenOld = SelectObject(lpDC, hPen)
    Call MoveToEx(lpDC, X, Y, PT)
    Call LineTo(lpDC, X2, Y2)
    Call SelectObject(lpDC, hPenOld)
    Call DeleteObject(hPen)
    
End Sub

Private Sub DrawSelection(lpDC As Long, X As Long, Y As Long, w As Long, h As Long, lIndex As Long)
Dim hBmp    As Long
Dim DC      As Long
Dim hDCMem  As Long
Dim hPen    As Long
Dim Alpha1  As Long
Dim lColor  As Long
Dim lH      As Long
Dim Px      As Long
Dim out     As Long
Dim I       As Long
Dim DivValue    As Double


    Select Case lIndex
        Case 0: lColor = pvAlphaBlend(m_SelColor, vbWhite, 110)
        Case 1: lColor = pvAlphaBlend(m_SelColor, vbWhite, 190)
        Case 2: lColor = m_SelColor
    End Select

    Px = lnScale \ 2
    out = Px \ 2
    lH = h - (2 * lnScale)

    DC = GetDC(0)
    hDCMem = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(DC, 1, lH)
    Call SelectObject(hDCMem, hBmp)
    
    Alpha1 = pvAlphaBlend(lColor, vbWhite, 45)
    For I = 0 To lH
        DivValue = ((I * 100) / lH)
        SetPixelV hDCMem, 0, I, pvAlphaBlend(lColor, Alpha1, DivValue)
    Next
    
    StretchBlt lpDC, X + lnScale, Y + lnScale, w - (lnScale * 2), lH, hDCMem, 0, 0, 1, lH, vbSrcCopy
    
    hPen = CreatePen(0, lnScale, lColor)
    Call SelectObject(lpDC, hPen)
    RoundRect lpDC, X + Px, Y + Px, X + w - out, Y + h - out, 3 * lnScale, 3 * lnScale
    DeleteObject hPen
    
    hPen = CreatePen(0, lnScale, pvAlphaBlend(lColor, vbWhite, 18))
    Call SelectObject(lpDC, hPen)
    RoundRect lpDC, X + Px + lnScale, Y + Px + lnScale, X + w - (lnScale + out), Y + h - (lnScale + out), 3 * lnScale, 3 * lnScale
    
    DeleteObject hPen
    DeleteObject hBmp
    DeleteDC DC
    DeleteDC hDCMem
    
End Sub

Private Function GetSafeRound(Angle As Integer, Width As Long, Height As Long) As Integer
    Dim lRet As Integer
    lRet = Angle
    If lRet * 2 > Height Then lRet = Height \ 2
    If lRet * 2 > Width Then lRet = Width \ 2
    GetSafeRound = lRet
End Function

Private Sub SafeRange(ByVal Value As Long, ByVal Min As Long, ByVal Max As Long)
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
End Sub

Private Function GDIpRoundBox(ByVal hDC As Long, Rect As RectF, ByVal BackColor, ByVal BorderW As Long, ByVal BorderColor As Long, mRound As Long, Side As eFlatSide) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mPath As Long
    Dim hGraphics As Long
    Dim Round As Integer
    
    GdipCreateFromHDC hDC, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    GdipCreateSolidFill BackColor, hBrush
    GdipCreatePen1 BorderColor, BorderW * e_Scale, &H2, hPen '&H2 * e_Scale, &H2, hPen
    GdipCreatePath &H0, mPath   '&H0
    
    With Rect
        Round = GetSafeRound(mRound * e_Scale, .Width, .Height)
        Round = IIf(Round = 0, 1, Round)
        
        Select Case Side
          Case rUp
              GdipAddPathArcI mPath, .Left, .Top, 1, 1, 180, 90
              GdipAddPathArcI mPath, (.Left + .Width) - 1, .Top, 1, 1, 270, 90
              GdipAddPathArcI mPath, (.Left + .Width) - Round, (.Top + .Height) - Round, Round, Round, 0, 90
              GdipAddPathArcI mPath, .Left, (.Top + .Height) - Round, Round, Round, 90, 90
          Case rBottom
              GdipAddPathArcI mPath, .Left, .Top, Round, Round, 180, 90
              GdipAddPathArcI mPath, (.Left + .Width) - Round, .Top, Round, Round, 270, 90
              GdipAddPathArcI mPath, (.Left + .Width) - 1, (.Top + .Height) - 1, 1, 1, 0, 90
              GdipAddPathArcI mPath, .Left, (.Top + .Height) - 1, 1, 1, 90, 90
          Case rLeft
              GdipAddPathArcI mPath, .Left, .Top, 1, 1, 180, 90
              GdipAddPathArcI mPath, (.Left + .Width) - Round, .Top, Round, Round, 270, 90
              GdipAddPathArcI mPath, (.Left + .Width) - Round, (.Top + .Height) - Round, Round, Round, 0, 90
              GdipAddPathArcI mPath, .Left, (.Top + .Height) - 1, 1, 1, 90, 90
          Case rRight
              GdipAddPathArcI mPath, .Left, .Top, Round, Round, 180, 90
              GdipAddPathArcI mPath, (.Left + .Width) - 1, .Top, 1, 1, 270, 90
              GdipAddPathArcI mPath, (.Left + .Width) - 1, (.Top + .Height) - 1, 1, 1, 0, 90
              GdipAddPathArcI mPath, .Left, (.Top + .Height) - Round, Round, Round, 90, 90
        End Select
    End With
    
    GdipClosePathFigure mPath
    GdipFillPath hGraphics, hBrush, mPath
    GdipDrawPath hGraphics, hPen, mPath
    
    Call GdipDeletePath(mPath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)
    GdipDeleteGraphics hGraphics

    GDIpRoundBox = mPath
End Function

Private Function GDIpCreateRoundPath(ByVal Left As Single, ByVal Top As Single, ByVal Width As Single, ByVal Height As Single, ByVal Radius As Single) As Long
    Dim hPath As Long
    If GdipCreatePath(&H0, hPath) = 0& Then
    
        If Radius > Width / 2 Then Radius = Width / 2
        If Radius > Height / 2 Then Radius = Height / 2
    
        If Radius = 0 Then
            GdipAddPathRectangle hPath, Left, Top, Width, Height
        Else
            Radius = Radius * 2
            GdipAddPathArc hPath, Left, Top, Radius, Radius, 180, 90
            GdipAddPathArc hPath, Left + Width - Radius, Top, Radius, Radius, 270, 90
            GdipAddPathArc hPath, Left + Width - Radius, Top + Height - Radius, Radius, Radius, 0, 90
            GdipAddPathArc hPath, Left, Top + Height - Radius, Radius, Radius, 90, 90
            GdipClosePathFigure hPath
        End If
        GDIpCreateRoundPath = hPath
    End If
    
End Function

Private Function GDIpRoundRect(ByVal hDC As Long, Rect As RectF, ByVal BackColor, ByVal BorderW As Long, ByVal BorderColor As Long, Round As Long) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mPath As Long
    Dim hGraphics As Long
    
    GdipCreateFromHDC hDC, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    GdipCreateSolidFill BackColor, hBrush
    GdipCreatePen1 BorderColor, BorderW * e_Scale, &H2, hPen '&H2 * e_Scale, &H2, hPen
    GdipCreatePath &H0, mPath   '&H0
    
    With Rect
        If Round = 0 Then
            GdipDrawRectangleI hGraphics, hPen, .Left, .Top, .Width, .Height
            GdipAddPathLineI mPath, .Left, .Top, .Width, .Top       'Line-Top
            GdipAddPathLineI mPath, .Width, .Top, .Width, .Height   'Line-Left
            GdipAddPathLineI mPath, .Width, .Height, .Left, .Height 'Line-Bottom
            GdipAddPathLineI mPath, .Left, .Height, .Left, .Top     'Line-Right
        Else
            GdipAddPathArcI mPath, .Left, .Top, Round, Round, 180, 90
            GdipAddPathArcI mPath, (.Left + .Width) - Round, .Top, Round, Round, 270, 90
            GdipAddPathArcI mPath, (.Left + .Width) - Round, (.Top + .Height) - Round, Round, Round, 0, 90
            GdipAddPathArcI mPath, .Left, (.Top + .Height) - Round, Round, Round, 90, 90
        End If
    End With
    
    GdipClosePathFigure mPath
    GdipFillPath hGraphics, hBrush, mPath
    GdipDrawPath hGraphics, hPen, mPath
    
    Call GdipDeletePath(mPath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)
    GdipDeleteGraphics hGraphics

    GDIpRoundRect = mPath
End Function

Private Function GetRowFromY(ByVal Y As Long) As Long
'If m_ShowHeader Then
If m_Header Then
    If Y <= lHeaderH Then
      GetRowFromY = -1
      Exit Function
    End If
    Y = Y + GetScroll(1) - lHeaderH
Else
    Y = Y + GetScroll(1)
End If
    GetRowFromY = Y \ m_RowH
    If GetRowFromY >= ItemCount Then GetRowFromY = -1
End Function

Private Function GetScroll(eBar As EFSScrollBarConstants) As Long
    GetScroll = IIf(VScrollBar.Visible(eBar), VScrollBar.Value(eBar), 0)
End Function

Private Function GetTextFlag(Col As Long) As Long
    GetTextFlag = &H4 Or &H20 Or &H40000 '-> VCenter Or SingleLine Or WordElipsis
    Select Case m_cols(Col).Aling
        Case 1: GetTextFlag = GetTextFlag Or &H2
        Case 2: GetTextFlag = GetTextFlag Or &H1
    End Select
End Function

'Inicia GDI+
Private Sub InitGDI()
    Dim gdipStartupInput As GdiplusStartupInput
    gdipStartupInput.GdiplusVersion = 1&
    Call GdiplusStartup(gdipToken, gdipStartupInput, ByVal 0)
End Sub
Private Function IsCompleteVisibleRow(eRow As Long) As Boolean
On Local Error Resume Next
Dim Y       As Long
Dim bRow    As Boolean
    Y = (eRow * m_RowH) - GetScroll(1)
    bRow = (Y >= 0) And (Y + m_ItemH <= lGridH)
    IsCompleteVisibleRow = bRow
End Function

Private Function IsMouseOver(hwnd As Long) As Boolean
    Dim PT As POINTAPI
    GetCursorPos PT
    IsMouseOver = (WindowFromPoint(PT.X, PT.Y) = hwnd)
End Function

Private Function IsVisibleRow(ByVal eRow As Long) As Boolean
On Error Resume Next
Dim Y As Long
    If VScrollBar.Visible(1) = False Then IsVisibleRow = True: Exit Function
    Y = (eRow * m_RowH) - GetScroll(1)
    IsVisibleRow = (Y + m_ItemH > 0) And Y <= lGridH
End Function

Private Function pFindText(ByVal Text As String, Optional ByVal iStart As Integer = -1, Optional IgnoreCase As Boolean, Optional CompleteString As Boolean) As Integer
On Error Resume Next
Dim j As Integer, K As Integer
Dim iText    As String
Dim iRet      As Integer
Dim tLn       As Integer
            
  If ItemCount = 0 Then pFindText = -1: Exit Function
  'Debug.Print "B: " & Text
  If iStart > ItemCount - 1 Then iStart = 0
  If IgnoreCase Then Text = UCase(Text)
  
  iRet = -1
  tLn = Len(Text)
  
  For j = iStart To ItemCount - 1
    For K = 0 To ColumnCount - 1
      iText = IIf(CompleteString, m_items(j).Item(K).Text, Left(m_items(j).Item(K).Text, tLn))
    Next K
    
      If IgnoreCase Then iText = UCase(iText)
      If iText <> "" Then
              If Text = iText Then: iRet = j: Exit For
      End If
  Next j
  
  If iRet = -1 And iStart > 0 Then
      For j = 0 To iStart
          For K = 0 To ColumnCount - 1
            iText = IIf(CompleteString, m_items(j).Item(K).Text, Left(m_items(j).Item(K).Text, tLn))
          Next K
          
          If IgnoreCase Then iText = UCase(iText)
          If iText <> "" Then
                  If Text = iText Then: iRet = j: Exit For
          End If
      Next
  End If
  pFindText = iRet
End Function

Private Function pvAlphaBlend(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
Dim clrFore(3)      As Byte
Dim clrBack(3)      As Byte

    OleTranslateColor clrFirst, 0, VarPtr(clrFore(0))
    OleTranslateColor clrSecond, 0, VarPtr(clrBack(0))
    
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
    CopyMemory pvAlphaBlend, clrFore(0), 4
End Function

Private Sub ReplaceColor(ByVal PictureBox As Object, ByVal FromColor As Long, ByVal ToColor As Long)
  If PictureBox.Picture Is Nothing Then Err.Raise Number:=1, Description:="Picture not set"
  If PictureBox.Picture.Handle = 0 Then Err.Raise Number:=2, Description:="Picture handle is null"
  Dim WinFromColor As Long, WinToColor As Long, MemAutoRedraw As Boolean
  WinFromColor = WinColor(FromColor)
  WinToColor = WinColor(ToColor)
  With PictureBox
    MemAutoRedraw = .AutoRedraw
    .AutoRedraw = True
    Dim X As Long, Y As Long
    For X = 0 To CInt(.ScaleX(.Picture.Width, vbHimetric, vbPixels))
        For Y = 0 To CInt(.ScaleY(.Picture.Height, vbHimetric, vbPixels))
            If GetPixel(.hDC, X, Y) = WinFromColor Then SetPixel .hDC, X, Y, WinToColor
        Next Y
    Next X
    .Refresh
    .Picture = .Image
    .AutoRedraw = MemAutoRedraw
  End With
End Sub

Private Function SendText(Text As String)
    SendMessage m_hWnd, &HC, 0&, Text
    SendMessage m_hWnd, &HB1, Len(Text), Len(Text)
End Function
Private Sub SetVisibleItem(eRow As Long)
On Error GoTo zErr
Dim lx  As Integer
Dim ly  As Integer

    If eRow = -1 Then Exit Sub
    ly = eRow * m_RowH

    '?Vertical
    If (ly + m_RowH) - lGridH > GetScroll(1) Then
        VScrollBar.Value(1) = ((ly + m_RowH) + 4) - lGridH
    ElseIf ly < GetScroll(1) Then
        VScrollBar.Value(1) = ly
    End If
zErr:
    DrawGrid
End Sub

Private Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
'Funcion para combinar dos colores
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
    
    If (clrFirst And &H80000000) Then clrFirst = GetSysColor(clrFirst And &HFF&)
    If (clrSecond And &H80000000) Then clrSecond = GetSysColor(clrSecond And &HFF&)
  
    CopyMemory clrFore(0), clrFirst, 4
    CopyMemory clrBack(0), clrSecond, 4
  
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
  
    CopyMemory ShiftColor, clrFore(0), 4
  
End Function

'Private Sub SetVisibleItem(eRow As Long)
'On Error GoTo zErr
'Dim lx  As Integer
'Dim ly  As Integer
'
'
'    If eRow = -1 Then Exit Sub
'    ly = eRow * m_RowH
'
'    '?Vertical
'    If (ly + m_RowH) - lGridH > GetScroll(1) Then
'        VScrollBar.Value(1) = ((ly + m_RowH) + 2) - lGridH
'    ElseIf ly < GetScroll(1) Then
'        VScrollBar.Value(1) = ly
'    End If
'zErr:
'    DrawGrid
'End Sub

Private Function SysColor(oColor As Long) As Long
    OleTranslateColor2 oColor, 0, SysColor
End Function

'Termina GDI+
Private Sub TerminateGDI()
    Call GdiplusShutdown(gdipToken)
End Sub

''/Start tracking of mouse leave event
'Private Sub TrackMouseTracking(hWnd As Long)
'Dim tEventTrack As tTrackMouseEvent
'
'    With tEventTrack
'        .cbSize = Len(tEventTrack)
'        .dwFlags = TME_LEAVE
'        .hwndTrack = hWnd
'    End With
'    If (m_bTrackHandler32) Then
'        'TrackMouseEvent tEventTrack
'        TrackMouseEvent pmTrack(0)
'    Else
'        TrackMouseEvent2 tEventTrack
'    End If
'End Sub

Private Sub UpateValues()
Dim TH As Integer
Dim Px As Long
    
    Px = 4 * e_Scale
    TH = UserControl.TextHeight("ÀjFq")
    
    If TH + Px > m_ItemH Then m_ItemH = TH + Px
    If m_img.Y + Px > m_ItemH Then m_ItemH = m_img.Y + Px
    m_RowH = m_ItemH
    If m_GridStyle = 1 Or m_GridStyle = 3 Then m_RowH = m_RowH + (1 * lnScale)
    UpdateScrollV
End Sub

Private Sub UpdateScrollV()
On Local Error Resume Next
Dim lHeight     As Long
Dim lProportion As Long
Dim ly          As Long
Dim bFlag       As Boolean

    ly = lGridH
    'Debug.Print "UpdateScrollV ly=" & ly
    'lHeight = ((ItemCount * m_RowH) + (2 * e_Scale)) - ly
    lHeight = (ItemCount * m_RowH) - (m_VisibleRows * m_RowH)
    
    If (lHeight > 0) Then
      lProportion = lHeight \ ly
      VScrollBar.LargeChange(1) = lHeight \ lProportion
      VScrollBar.Max(1) = lHeight + 1
      VScrollBar.Visible(1) = True
    Else
      VScrollBar.Visible(1) = False
    End If
    
End Sub

Private Function WinColor(ByVal Color As Long, Optional ByVal hPal As Long) As Long
If OleTranslateColor2(Color, hPal, WinColor) <> 0 Then WinColor = -1
End Function

Private Sub oText_Change()
RaiseEvent Change
End Sub

Private Sub oText_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lT As Integer

        Select Case KeyCode
            Case 13                                       '{Enter}
                If pList.Visible Then
                    ShowList False
                Else
                    ShowList True
                End If
            Case 38                                       '{Up arrow}
                KeyCode = 0
                If t_Row > 0 Then ListIndex = ListIndex - 1
                oText.Text = ItemText(t_Row, m_ColumnInBox)
                SetVisibleItem t_Row
                
            Case 40                                       '{Down arrow}
                KeyCode = 0
                If t_Row < ItemCount - 1 Then ListIndex = ListIndex + 1
                oText.Text = ItemText(t_Row, m_ColumnInBox)
                SetVisibleItem t_Row
            
            Case 33                                       '{PageUp}
               ' If (m_ListIndex > m_VisibleRows) Then
                   ' ListIndex = ListIndex - (m_VisibleRows - 1)
                'Else                                      'NOT (M_ListIndex...
                    'ListIndex = 0
                'End If
            Case 34                                       '{PageDown}
               ' If (m_ListIndex < m_nItems - m_VisibleRows - 1) Then
                    'ListIndex = ListIndex + (m_VisibleRows - 1)
                'Else                                      'NOT (M_ListIndex...
                   ' ListIndex = m_nItems - 1
                'End If
            Case 36                                       '{Start}
                KeyCode = 0
                ListIndex = 0
                
            Case 35                                       '{End}
                KeyCode = 0
                ListIndex = ItemCount - 1
                
            Case 27
                    If pList.Visible Then ShowList False
                    
            Case Else
                Dim NewIndex As Integer
                
                If Chr(KeyCode) = "" Then Exit Sub
                  If m_ComboStyle = axJListCombo Then
                        NewIndex = pFindText(Chr(KeyCode), t_Row + 1, True)
                        If NewIndex <> -1 Then ListIndex = NewIndex
                  End If
        End Select
        
RaiseEvent ItemClick(t_Row)
End Sub

Private Sub pList_Click()
Dim C As Long, cString As String
On Error Resume Next
    If t_Row <> -1 Then
        If IsCompleteVisibleRow(t_Row) Then
          If m_MultiLine Then
            For C = 0 To ColumnCount - 1
              cString = cString & ItemText(t_Row, C) & vbCrLf
            Next C
            oText.Text = cString
          Else
            oText.Text = ItemText(t_Row, m_ColumnInBox)
          End If
          
          RaiseEvent ItemClick(t_Row)
          ShowList False
        Else
          SetVisibleItem t_Row
        End If
    End If
End Sub

Private Sub pList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = vbLeftButton) Then Call SetCapture(pList.hwnd)

End Sub

Private Sub pList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lRow    As Long

    lRow = GetRowFromY(Y)
    If X > m_GridW Then lRow = -1
    If lRow <> t_Row Then
        t_Row = lRow
        DrawGrid
    End If
    
    'Debug.Print "Row=" & t_Row
End Sub

Private Sub pList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then Call ReleaseCapture
End Sub

Private Sub TmrMouseOver_Timer()
If IsMouseOver(UserControl.hwnd) Or IsMouseOver(pList.hwnd) Then
    ReDrawControl 90
Else
    TmrMouseOver.Enabled = False
    ReDrawControl 50
End If
End Sub

Private Sub UserControl_EnterFocus()
mInFocus = True
End Sub

Private Sub UserControl_ExitFocus()
m_MouseOver = False
ShowList False
End Sub

Private Sub UserControl_Initialize()
    Set cSubClass = New c_SubClass
    Set VScrollBar = New c_ScrollBars
    
    InitGDI
    
    t_Row = -1: m_SelRow = -1
    e_Scale = GetWindowsDPI
    Select Case e_Scale
        Case 1, 2: lnScale = 1
        Case 3, 4: lnScale = 2
        Case 5: lnScale = 4
    End Select
End Sub

Private Sub UserControl_InitProperties()
    m_HeaderH = 24
    m_GridLineColor = &HF0F0F0
    m_GridStyle = 3
    m_Striped = True
    m_StripedColor = &HFDFDFD
    m_SelColor = vbHighlight  '&HDDAC84
    m_BorderColor = &H908782  '&HB2ACA5
    m_BackColor = vbWhite
    m_BackColorParent = Ambient.BackColor
    m_BorderWidth = 1
    m_CornerRound = 5
    m_Header = True
    m_VisibleRows = 8
    m_Visible = False
    m_MouseOver = False
    mEnabled = True
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With RctButton
  If X > .Left And X < (.Left + .Width) And Y > .Top And Y < (.Top + .Height) Then
    If m_Visible = False Then
      m_MouseOver = True
    Else
      m_MouseOver = False
    End If
  End If
  
  ReDrawControl
  ShowList m_MouseOver
End With
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TmrMouseOver.Enabled = True

If X > UserControl.ScaleWidth Or X < 0 Or Y > UserControl.ScaleHeight Or Y < 0 Then
    ReleaseCapture
    mInCtrl = False
ElseIf mInCtrl Then
    RaiseEvent MouseMove(Button, Shift, X, Y)
Else
    mInCtrl = True
    Call TrackMouseLeave(UserControl.hwnd)
    ReDrawControl 90
    RaiseEvent MouseMove(Button, Shift, X, Y)
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim h As Double
    With PropBag
        mEnabled = .ReadProperty("Enabled", True)
        m_HeaderH = .ReadProperty("HeaderH", 24)
        m_GridLineColor = .ReadProperty("LineColor", &HF0F0F0)
        m_GridStyle = .ReadProperty("GridStyle", 3)
        m_Striped = .ReadProperty("Striped", True)
        m_StripedColor = .ReadProperty("StripedColor", &HFDFDFD)
        m_SelColor = .ReadProperty("SelColor", vbHighlight)
        m_ItemH = .ReadProperty("ItemH", 0)
        m_BorderColor = .ReadProperty("BorderColor", &HB2ACA5)
        m_BorderWidth = .ReadProperty("BorderWidth", 1)
        m_CornerRound = .ReadProperty("CornerCurve", 5)
        m_Header = .ReadProperty("Header", True)
        m_ForeColor = .ReadProperty("ForeColor", vbBlack)
        oText.Text = .ReadProperty("Text", "")
        m_VisibleRows = .ReadProperty("VisibleRows", 8)
        m_DropW = .ReadProperty("DropWidth", UserControl.Width / Screen.TwipsPerPixelX)
        m_ComboStyle = .ReadProperty("ComboStyle", 0)
        Set oText.Font() = .ReadProperty("Font", oText.Font)
        m_BackColor = .ReadProperty("BackColor", vbWhite)
        m_BackColorParent = .ReadProperty("BackColorParent", Ambient.BackColor)
        m_ButtonColorPress = .ReadProperty("ButtonColorPress", &H40&)
        m_ColumnInBox = .ReadProperty("ColumnInBox", 0)
        m_MultiLine = .ReadProperty("MultiLine", False)
        
        Set m_IconFont = .ReadProperty("IconFont", Ambient.Font)
        m_IconCharCode = .ReadProperty("IconCharCode", "&Hea67")
        m_IconForeColor = .ReadProperty("IconForeColor", &H404040)
        
        m_PadX = .ReadProperty("IcoPaddingX", 0)
        m_PadY = .ReadProperty("IcoPaddingY", 0)
    End With

    oText.Locked = CBool(m_ComboStyle)
    

    If Ambient.UserMode Then
            
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
        
        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If
        Debug.Print "TrackMouseEvent:" & bTrack
      
        With VScrollBar
            .Create pList.hwnd
            .Style = efsFlat
            .SmallChange(0) = 20 '48
            .SmallChange(1) = 16
        End With
                        
        With cSubClass
            If .Subclass(UserControl.hwnd, , , Me) Then
                .AddMsg UserControl.hwnd, WM_KILLFOCUS, MSG_AFTER
                '.AddMsg UserControl.hWnd, WM_WINDOWPOSCHANGED, MSG_AFTER
                .AddMsg UserControl.hwnd, WM_MOUSELEAVE, MSG_AFTER
                .AddMsg UserControl.hwnd, WM_NCPAINT, MSG_AFTER
            End If
            
            If .Subclass(pList.hwnd, , , Me) Then
                .AddMsg pList.hwnd, WM_KILLFOCUS, MSG_AFTER
                '.AddMsg plist.hwnd, WM_SETFOCUS, MSG_AFTER
                .AddMsg pList.hwnd, WM_MOUSELEAVE, MSG_AFTER
                .AddMsg pList.hwnd, WM_NCPAINT, MSG_AFTER
            End If
                        
          m_PhWnd = UserControl.ContainerHwnd
          
            If .Subclass(m_PhWnd, , , Me) Then
                .AddMsg m_PhWnd, WM_WINDOWPOSCHANGING, MSG_AFTER
                .AddMsg m_PhWnd, WM_WINDOWPOSCHANGED, MSG_AFTER
                .AddMsg m_PhWnd, WM_GETMINMAXINFO, MSG_AFTER
                .AddMsg m_PhWnd, WM_LBUTTONDOWN, MSG_AFTER
                .AddMsg m_PhWnd, WM_SIZE, MSG_AFTER
                .AddMsg m_PhWnd, 516, MSG_BEFORE ' MouseDown
                .AddMsg m_PhWnd, WM_ACTIVATE, MSG_AFTER ' MouseUp
                .AddMsg m_PhWnd, 164, MSG_BEFORE ' Menu
                .AddMsg m_PhWnd, WM_SYSCOMMAND, MSG_BEFORE
                .AddMsg m_PhWnd, WM_LBUTTONUP, MSG_BEFORE
            End If
            
        End With
        
        'pmTrack(0) = 16&
        'pmTrack(1) = &H2
        'pmTrack(2) = pList.hWnd
               
        m_HeaderH = m_HeaderH * e_Scale
        h = UserControl.TextHeight("Ájq")
        If m_HeaderH < (h + (6 * e_Scale)) Then HeaderHeight = h + (6 * e_Scale)
        
        UpateValues
    End If
    
    ReDrawControl
    
End Sub

Private Sub UserControl_Resize()
Dim CtrlH As Integer

With UserControl
    TextH = .TextHeight("ÁjWHz")
    CtrlH = TextH + 6
    
    If CtrlH < 25 Then CtrlH = 25
    If .ScaleHeight < CtrlH Then .Height = CtrlH * 15
    
    If m_MultiLine Then
      oText.Move 4, 4, .ScaleWidth - 36, .ScaleHeight - 6
      'oText.ScrollBars = 2
    Else
      oText.Move 4, (.ScaleHeight - oText.Height) / 2, .ScaleWidth - 36, TextH
      'oText.ScrollBars = 0
    End If
    
    If .Height <= TextH Then
      .Height = TextH + 6
    End If
    
End With
        
 ReDrawControl
        
End Sub

Private Sub UserControl_Terminate()
On Error Resume Next

Call cSubClass.UnSubclass(m_PhWnd)
Call cSubClass.UnSubclass(UserControl.hwnd)
Call cSubClass.UnSubclass(oText.hwnd)
    
Erase m_items
Erase m_cols
Set cSubClass = Nothing
Set VScrollBar = Nothing

TerminateGDI
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Enabled", mEnabled
        .WriteProperty "HeaderH", m_HeaderH
        .WriteProperty "LineColor", m_GridLineColor
        .WriteProperty "GridStyle", m_GridStyle
        .WriteProperty "Striped", m_Striped
        .WriteProperty "StripedColor", m_StripedColor
        .WriteProperty "SelColor", m_SelColor
        .WriteProperty "ItemH", m_ItemH
        .WriteProperty "BorderColor", m_BorderColor
        .WriteProperty "BorderWidth", m_BorderWidth
        .WriteProperty "CornerCurve", m_CornerRound
        .WriteProperty "Header", m_Header
        .WriteProperty "ForeColor", m_ForeColor
        .WriteProperty "Font", oText.Font
        .WriteProperty "BackColor", m_BackColor, vbWhite
        .WriteProperty "BackColorParent", m_BackColorParent
        .WriteProperty "ComboStyle", m_ComboStyle, 0
        .WriteProperty "VisibleRows", m_VisibleRows
        .WriteProperty "DropWidth", m_DropW
        .WriteProperty "Text", oText.Text, ""
        .WriteProperty "ColumnInBox", m_ColumnInBox, 0
        .WriteProperty "MultiLine", m_MultiLine, False
        
        .WriteProperty "ButtonColorPress", m_ButtonColorPress
        
        .WriteProperty "IconFont", m_IconFont
        .WriteProperty "IconCharCode", m_IconCharCode, 0
        .WriteProperty "IconForeColor", m_IconForeColor, vbButtonText
        
        .WriteProperty "IcoPaddingX", m_PadX
        .WriteProperty "IcoPaddingY", m_PadY
    End With
End Sub

Private Sub VScrollBar_Change(eBar As EFSScrollBarConstants)
    DrawGrid
End Sub
Private Sub VScrollBar_MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)
    DrawGrid
End Sub

Property Get BackColor() As OLE_COLOR
BackColor = m_BackColor
End Property

Property Let BackColor(ByVal Value As OLE_COLOR)
    m_BackColor = Value
    ReDrawControl
    PropertyChanged "BackColor"
End Property

Property Get BackColorParent() As OLE_COLOR
BackColorParent = m_BackColorParent
End Property

Property Let BackColorParent(ByVal Value As OLE_COLOR)
    m_BackColorParent = Value
    ReDrawControl
    PropertyChanged "BackColor"
End Property

Property Get BorderColor() As OLE_COLOR
BorderColor = m_BorderColor
End Property

Property Let BorderColor(ByVal Value As OLE_COLOR)
    m_BorderColor = Value
    ReDrawControl
    PropertyChanged "BorderColor"
End Property

Property Get BorderWidth() As Long
BorderWidth = m_BorderWidth
End Property

Property Let BorderWidth(ByVal nBorder As Long)
  m_BorderWidth = nBorder
  PropertyChanged "BorderWidth"
  ReDrawControl
End Property

Property Get ButtonColorPress() As OLE_COLOR
ButtonColorPress = m_ButtonColorPress
End Property

Property Let ButtonColorPress(ByVal Value As OLE_COLOR)
    m_ButtonColorPress = Value
    ReDrawControl
    PropertyChanged "ButtonColorPress"
End Property

Property Get ColumnCount() As Long
On Local Error Resume Next
    ColumnCount = UBound(m_cols) + 1
End Property

Public Property Get ColumnInBox() As Integer
    ColumnInBox = m_ColumnInBox
End Property

Public Property Let ColumnInBox(ByVal NewColumnInBox As Integer)
On Error Resume Next
'Debug.Print "UBound(m_cols):" & UBound(m_cols)
If UBound(m_cols) >= NewColumnInBox Then
    m_ColumnInBox = NewColumnInBox
Else
    m_ColumnInBox = LBound(m_cols)
End If
    PropertyChanged "ColumnInBox"
End Property

Public Property Get ComboStyle() As JComboStyle
  ComboStyle = m_ComboStyle
End Property

Public Property Let ComboStyle(ByVal NewComboStyle As JComboStyle)
  m_ComboStyle = NewComboStyle
  oText.Locked = IIf(m_ComboStyle = axJListCombo, True, False)
  PropertyChanged "ComboStyle"
End Property

Public Property Get CornerRound() As Long
  CornerRound = m_CornerRound
End Property

Public Property Let CornerRound(ByVal NewCornerRound As Long)
  m_CornerRound = NewCornerRound  'IIf(NewCornerRound > 12, 12, NewCornerRound)
  PropertyChanged "CornerRound"
  UserControl_Resize
End Property

Public Property Get DropWidth() As Long
  DropWidth = m_DropW
End Property

Public Property Let DropWidth(ByVal NewDropWidth As Long)
  m_DropW = NewDropWidth
  PropertyChanged "DropWidth"
End Property

Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    mEnabled = NewValue
    oText.Enabled = mEnabled
    UserControl.Enabled = mEnabled
    If mEnabled Then ReDrawControl 30 Else ReDrawControl 50
End Property

Property Get Font() As StdFont: Set Font = oText.Font: End Property
Property Set Font(ByVal Value As StdFont)
    Set oText.Font() = Value
    PropertyChanged "Font"
    UserControl_Resize
End Property

Property Get ForeColor() As OLE_COLOR: ForeColor = m_ForeColor: End Property
Property Let ForeColor(ByVal Value As OLE_COLOR)
    m_ForeColor = Value
    PropertyChanged "ForeColor"
    ReDrawControl
End Property

Property Get GridLineColor() As OLE_COLOR: GridLineColor = m_GridLineColor: End Property
Property Let GridLineColor(ByVal Value As OLE_COLOR)
    m_GridLineColor = Value
    PropertyChanged "LineColor"
    DrawGrid
End Property
Property Get GridLineStyle() As ScrollBarConstants: GridLineStyle = m_GridStyle: End Property
Property Let GridLineStyle(ByVal Value As ScrollBarConstants)
    m_GridStyle = Value
    UpateValues
    PropertyChanged "GridStyle"
End Property
Property Get Header() As Boolean: Header = m_Header: End Property
Property Get HeaderHeight() As Long: HeaderHeight = m_HeaderH: End Property
Property Let HeaderHeight(ByVal Value As Long)
    m_HeaderH = Value
    UpateValues
    PropertyChanged "HeaderH"
End Property
Property Let Header(Value As Boolean)
    m_Header = Value
    PropertyChanged "Header"
End Property

''END ICONFONT----------------------------------------

Public Property Get hwnd()
  hwnd = UserControl.hwnd
End Property

Public Property Get IconCharCode() As String
    IconCharCode = "&H" & Hex(m_IconCharCode)
End Property

Public Property Let IconCharCode(ByVal New_IconCharCode As String)
    New_IconCharCode = UCase(Replace(New_IconCharCode, Space(1), vbNullString))
    New_IconCharCode = UCase(Replace(New_IconCharCode, "U+", "&H"))
    If Not Left(New_IconCharCode, 2) = "&H" And Not IsNumeric(New_IconCharCode) Then
        m_IconCharCode = "&H" & New_IconCharCode
    Else
        m_IconCharCode = New_IconCharCode
    End If
    PropertyChanged "IconCharCode"
    ReDrawControl
End Property


''ICONFONT--------------------------------------------
Public Property Get IconFont() As StdFont
    Set IconFont = m_IconFont
End Property

Public Property Set IconFont(New_Font As StdFont)
Set m_IconFont = New_Font
'    With m_IconFont
'        .name = New_Font.name
'        .Size = New_Font.Size
'        .Bold = New_Font.Bold
'        .Italic = New_Font.Italic
'        .Strikethrough = New_Font.Strikethrough
'        .Underline = New_Font.Underline
'        .Weight = New_Font.Weight
'    End With
    PropertyChanged "IconFont"
    ReDrawControl
End Property

Public Property Get IconForeColor() As OLE_COLOR
    IconForeColor = m_IconForeColor
End Property

Public Property Let IconForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_IconForeColor = New_ForeColor
    PropertyChanged "IconForeColor"
    ReDrawControl
End Property

Public Property Get IcoPaddingX() As Long
IcoPaddingX = m_PadX
End Property

Public Property Let IcoPaddingX(ByVal XpadVal As Long)
m_PadX = XpadVal
PropertyChanged "IcoPaddingX"
ReDrawControl
End Property

Public Property Get IcoPaddingY() As Long
IcoPaddingY = m_PadY
End Property

Public Property Let IcoPaddingY(ByVal YpadVal As Long)
m_PadY = YpadVal
PropertyChanged "IcoPaddingY"
ReDrawControl
End Property
Property Get ItemCount() As Long
On Local Error Resume Next
    ItemCount = UBound(m_items) + 1
End Property
Property Get ItemHeight() As Long: ItemHeight = m_ItemH: End Property
Property Let ItemHeight(ByVal Value As Long)
    m_ItemH = Value
    UpateValues
    PropertyChanged "ItemH"
End Property

Property Get ItemText(ByVal Item As Long, Optional ByVal Column As Long) As String
On Local Error Resume Next
    ItemText = m_items(Item).Item(Column).Text
End Property
Property Let ItemText(ByVal Item As Long, Optional ByVal Column As Long, Value As String)
On Local Error Resume Next
    If m_items(Item).Item(Column).Text = Value Then Exit Property
    m_items(Item).Item(Column).Text = Value
End Property

Property Get ListIndex() As Integer
  ListIndex = t_Row
End Property

Property Let ListIndex(NewIndex As Integer)
Dim lT As Integer
    If NewIndex > ItemCount - 1 Then Exit Property
    t_Row = NewIndex
    DrawGrid
    
    If t_Row < VScrollBar.Value(1) And t_Row > -1 Then
        VScrollBar.Value(1) = t_Row
    ElseIf (t_Row > VScrollBar.Value(1) + m_VisibleRows - 1) Then
        VScrollBar.Value(1) = t_Row - m_VisibleRows + 1
    End If
    RaiseEvent ListIndexChanged(t_Row)
    If pList.Visible Then DrawGrid
End Property

Public Property Get MultiLine() As Boolean
  MultiLine = m_MultiLine
End Property

Public Property Let MultiLine(ByVal NewMultiLine As Boolean)
  m_MultiLine = NewMultiLine
  PropertyChanged "MultiLine"
  UserControl_Resize
End Property

Property Get SelectionColor() As OLE_COLOR: SelectionColor = m_SelColor: End Property
Property Let SelectionColor(ByVal Value As OLE_COLOR)
    m_SelColor = Value
    PropertyChanged "SelColor"
End Property
Property Get StripBackColor() As OLE_COLOR: StripBackColor = m_StripedColor: End Property
Property Let StripBackColor(ByVal Value As OLE_COLOR)
    m_StripedColor = Value
    PropertyChanged "StripedColor"
End Property
Property Get StripedGrid() As Boolean: StripedGrid = m_Striped: End Property
Property Let StripedGrid(ByVal Value As Boolean)
    m_Striped = Value
    PropertyChanged "Striped"
End Property

Public Property Get Text() As String
  'm_Text = oText.Text
  Text = oText.Text
End Property

Public Property Let Text(ByVal newText As String)
  'm_Text = NewText
  oText.Text = newText
  PropertyChanged "Text"
End Property

Property Get VisibleRows() As Long: VisibleRows = m_VisibleRows: End Property
Property Let VisibleRows(ByVal Value As Long)
    m_VisibleRows = Value
    PropertyChanged "VisibleRows"
End Property
Private Property Get lGridH() As Long
    lGridH = (m_VisibleRows * m_RowH) + lHeaderH + (4 * e_Scale)
    'lGridH = UserControl.ScaleHeight - lHeaderH
End Property

Private Property Get lHeaderH() As Long
    'lHeaderH = 0
    'lHeaderH = IIf(m_ShowHeader, m_HeaderH, 0)
    lHeaderH = IIf(m_Header, m_HeaderH, 0)
End Property

''+++SUBCLASS++++++++++++++++++++++++++++++++++++
Private Sub tmrRelease_Timer()
    Dim uPoint  As POINTAPI
    Dim uRect As Rect
    Dim nLB As Integer
    Dim nRB As Integer
    
    '#############################################################################################################################
    'This is soley for detecting if we have clicked on a container which does not generate
    'WM_KILLFOCUS message for us to detect. i.e. the parent Form or a Frame
    
    'I don't like Timers in UserControls but wanted to make the Control behave as a normal Combo which
    'closes DropDown when the above situation occurs. I may still remove this "feature"!
    
    'NOTE: This Timer is only Enabled when we detect a WM_MOUSELEAVE so it does not fire unneccessarily
    'while the DropDown is displayed. It is Disabled as soon as the mouse re-enters the DropDown.
    '#############################################################################################################################
    
    Call GetCursorPos(uPoint)
    Call GetWindowRect(pList.hwnd, uRect)
        
    nLB = GetAsyncKeyState(VK_LBUTTON)
    nRB = GetAsyncKeyState(VK_RBUTTON)
    
    If (uPoint.X >= uRect.L) And (uPoint.X <= uRect.r) And (uPoint.Y >= uRect.T) And (uPoint.Y <= uRect.B) Then
        Debug.Print "The mouse pointer is within the Dropdown list"
    ElseIf nLB Or nRB Then
        Select Case WindowFromPoint(uPoint.X, uPoint.Y)
        Case UserControl.hwnd Or pList.hwnd
          Debug.Print "The mouse pointer is within the Control"
        Case Else
            If m_Visible Then
               ShowList False
            Else
               SetTimer False
            End If
        End Select
    End If
End Sub

Private Sub SetTimer(bEnabled As Boolean)
    If tmrRelease.Enabled <> bEnabled Then
        If bEnabled Then
            tmrRelease.Enabled = True
        Else
            tmrRelease.Enabled = False
        End If
    End If
End Sub

'Track the mouse hovering the indicated window
Private Sub TrackMouseHover(ByVal lng_hWnd As Long, lHoverTime As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
  If bTrack Then
    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_HOVER
      .hwndTrack = lng_hWnd
      .dwHoverTime = lHoverTime
    End With

    If bTrackUser32 Then
      Call TrackMouseEvent(tme)
    Else
      Call TrackMouseEventComCtl(tme)
    End If
  End If
End Sub

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
  If bTrack Then
    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hwndTrack = lng_hWnd
    End With

    If bTrackUser32 Then
      Call TrackMouseEvent(tme)
    Else
      Call TrackMouseEventComCtl(tme)
    End If
  End If
End Sub

'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
  Dim hMod        As Long
  Dim bLibLoaded  As Boolean

  hMod = GetModuleHandleA(sModule)

  If hMod = 0 Then
    hMod = LoadLibraryA(sModule)
    If hMod Then
      bLibLoaded = True
    End If
  End If

  If hMod Then
    If GetProcAddress(hMod, sFunction) Then
      IsFunctionExported = True
    End If
  End If

  If bLibLoaded Then
    Call FreeLibrary(hMod)
  End If
End Function

Private Sub WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, _
                    ByRef lReturn As Long, ByVal lng_hWnd As Long, _
                    ByVal uMsg As ssc_eMsg, ByVal wParam As Long, _
                    ByVal lParam As Long, ByRef lParamUser As Long)
    Select Case uMsg
        Case 516, 164, WM_SYSCOMMAND, WM_NCPAINT, WM_ACTIVATE
            If m_Visible Then ShowList False
            
        Case WM_KILLFOCUS, WM_WINDOWPOSCHANGING, WM_WINDOWPOSCHANGED, WM_GETMINMAXINFO, WM_SIZE, WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_LBUTTONUP
            If m_Visible Then ShowList False
                                    
        Case WM_MOUSELEAVE
          mInCtrl = False
          If pList.Visible Then
              Call GetAsyncKeyState(VK_LBUTTON)
              Call GetAsyncKeyState(VK_RBUTTON)
              Debug.Print "WndProc SetTimer"
              SetTimer True
          End If
        
        Case WM_MOUSEMOVE
            SetTimer False
        
            If Not mInCtrl Then
                mInCtrl = True
                Call TrackMouseLeave(lng_hWnd)
                Call TrackMouseHover(lng_hWnd, 0)
            End If

          
        Case WM_VSCROLL, WM_HSCROLL, WM_MOUSEWHEEL


    End Select
    
    Debug.Print "hWnd:" & lng_hWnd
    Debug.Print "uMsg:" & uMsg
End Sub


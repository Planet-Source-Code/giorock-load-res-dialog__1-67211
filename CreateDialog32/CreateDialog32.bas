Attribute VB_Name = "Module1"
Option Explicit
'***********************
'  Written by GioRock  *
'***********************
'***********************
'      Completely      *
'  Created by GioRock  *
'***********************

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Image Function
Private Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const IMAGE_ICON = 1
Private Const LR_COPYFROMRESOURCE = &H4000

' Paint Function
Private Type PAINTSTRUCT
    hDC As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved As Byte
End Type
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Const WM_ERASEBKGND = &H14
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Const OPAQUE = 2
Private Const TRANSPARENT = 1
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private hBrush As Long
Private hBrush2 As Long

' Font Function
Private Const LF_FACESIZE = 32
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * LF_FACESIZE
End Type
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private hFont As Long

' Window Function
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long

' Window Creation and Destruction Function
Private Declare Function DialogBoxIndirectParam Lib "user32" Alias "DialogBoxIndirectParamA" (ByVal hInstance As Long, hDialogTemplate As Any, ByVal hWndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EndDialog Lib "user32" (ByVal hDlg As Long, ByVal nResult As Long) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long

' Window Message Function
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' Window Message Constant
Private Const WM_INITDIALOG = &H110
Private Const WM_DESTROY = &H2
Private Const WM_COMMAND = &H111
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_CLOSE = &H10
Private Const WM_SYSCOMMAND = &H112
Private Const SC_CLOSE = &HF060&
Private Const WM_SIZE = &H5
Private Const STN_CLICKED = 0
Private Const STM_SETICON = &H170
Private Const WM_PAINT = &HF
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const WM_CTLCOLORDLG = &H136
Private Const WM_CTLCOLORSTATIC = &H138
Private Const WM_GETFONT = &H31
Private Const WM_SETFONT = &H30
Private Const WM_MOUSEMOVE = &H200
Private Const MK_ALT = (&H20)
Private Const MK_CONTROL = &H8
Private Const MK_LBUTTON = &H1
Private Const MK_MBUTTON = &H10
Private Const MK_RBUTTON = &H2
Private Const MK_SHIFT = &H4

' User Function
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_HIDE = 0
Private Const SW_MAXIMIZE = 3
Private Const SW_NORMAL = 1
Private Const SW_SHOW = 5
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GCL_HICON = -14

' Memory Manage Function
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' ScrollBar Function
Private Declare Function SetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Private Const SB_BOTH = 3
Private Const SB_CTL = 2
Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Const SB_THUMBPOSITION = 4
Private Const SB_THUMBTRACK = 5
Private Const SB_ENDSCROLL = 8
Private Const SB_LEFT = 6
Private Const SB_LINEDOWN = 1
Private Const SB_LINELEFT = 0
Private Const SB_LINERIGHT = 1
Private Const SB_LINEUP = 0
Private Const SB_PAGEDOWN = 3
Private Const SB_PAGELEFT = 2
Private Const SB_PAGERIGHT = 3
Private Const SB_PAGEUP = 2
Private Const SB_BOTTOM = 7
Private Const SB_TOP = 6
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Private Const SIF_DISABLENOSCROLL = &H8
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_RANGE = &H1
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

' Resource ID Constant
Private Const VB_RES_DIALOG As Long = 5
Private Const IDD_FORMVIEW As Long = 101
Private Const IDC_PICTURE1 As Long = 102
Private Const CONAPPICON As String = "CONAPP"

' Size of Image Object
Private Const PICTURESIZE As Long = 44

Private Function HIWORD(ByVal Value As Long) As Integer
    MoveMemory HIWORD, ByVal VarPtr(Value) + 2, 2
End Function

Private Function MAKELONG(ByVal wLow As Integer, ByVal wHi As Integer) As Long
Dim ml As Long
    ml = wLow
    MoveMemory ByVal VarPtr(ml) + 2, wHi, 2
    MAKELONG = ml
End Function

Private Function LOWORD(ByVal Value As Long) As Integer
    MoveMemory LOWORD, Value, 2
End Function


Public Sub Main()
Dim hDlgTest As Long
Dim sTplDlg As String

    sTplDlg = CStr(LoadResData(IDD_FORMVIEW, VB_RES_DIALOG))
    
    hDlgTest = DialogBoxIndirectParam(App.hInstance, ByVal StrPtr(sTplDlg), 0, GetAddr(AddressOf DialoProc), WM_INITDIALOG)
    If hDlgTest > 0 Then
        ShowWindow hDlgTest, SW_SHOW
        DeleteObject hBrush
        DeleteObject hBrush2
        DeleteObject hFont
        DestroyWindow hDlgTest
        hDlgTest = 0
    End If
    
    End

End Sub

Public Function DialoProc(ByVal hDlg As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Select Case uMsg
        Case WM_INITDIALOG
            hBrush = CreateSolidBrush(RGB(255, 255, 220))
            hBrush2 = CreateSolidBrush(RGB(255, 255, 255))
            SetClassLong hDlg, GCL_HICON, LoadResPicture(CONAPPICON, vbResIcon).Handle
            SetWindowText hDlg, "Test Create Dialog from VB"
            SendMessage GetDlgItem(hDlg, IDC_PICTURE1), STM_SETICON, CopyImage(LoadResPicture(CONAPPICON, vbResIcon).Handle, IMAGE_ICON, 0, 0, LR_COPYFROMRESOURCE), Null
            Dim LF As LOGFONT
            hFont = SendMessage(hDlg, WM_GETFONT, 0, 0)
            Call GetObject(hFont, Len(LF), LF)
            hFont = CreateFontIndirect(LF)
            DialoProc = True
            Exit Function
        Case WM_CTLCOLORDLG
            DialoProc = hBrush
            Exit Function
        Case WM_CTLCOLORSTATIC
            DialoProc = hBrush2
            Exit Function
        Case WM_COMMAND
            If LOWORD(wParam) = IDC_PICTURE1 Then
                If HIWORD(wParam) = STN_CLICKED Then
                    MsgBox "You have clicked me!"
                End If
            End If
            DialoProc = True
            Exit Function
        Case WM_LBUTTONDOWN
            SendMessage hDlg, WM_MOUSEMOVE, MK_LBUTTON, ByVal lParam
            DialoProc = 0
            Exit Function
        Case WM_PAINT
            Dim hDlgPic As Long, RCW As RECT, RCP As RECT, PS As PAINTSTRUCT
            GetClientRect hDlg, RCW
            InvalidateRect hDlg, RCW, True
            BeginPaint hDlg, PS
'                FillRect PS.hDC, RCW, hBrush
            EndPaint hDlg, PS
            SelectObject PS.hDC, hFont
            SetBkMode PS.hDC, TRANSPARENT
            SetTextColor PS.hDC, RGB(0, 128, 0)
            GetClientRect hDlg, RCW
            hDlgPic = GetDlgItem(hDlg, IDC_PICTURE1)
            GetWindowRect hDlgPic, RCP
            ScreenToClient hDlg, RCP
            RCP.Top = RCP.Top - 25
            TextOut PS.hDC, RCP.Left, RCP.Top, "I'm here!" + Chr$(0), 9
'        Case WM_ERASEBKGND
'            DialoProc = True
'            Exit Function
        Case WM_SYSCOMMAND
            If wParam = SC_CLOSE Then
                EndDialog hDlg, 0
                DialoProc = 0
                Exit Function
            End If
        Case WM_SIZE
            Dim hDlgP As Long, RC As RECT, SI As SCROLLINFO
            hDlgP = GetDlgItem(hDlg, IDC_PICTURE1)
            GetClientRect hDlg, RC
            MoveWindow hDlgP, Center(RC.Right, PICTURESIZE), Center(RC.Bottom, PICTURESIZE), PICTURESIZE, PICTURESIZE, True
            With SI
                .cbSize = Len(SI)
                .fMask = SIF_ALL
                .nMin = 0
                .nMax = RC.Right - PICTURESIZE
                .nPos = Center(RC.Right - PICTURESIZE, 0)
            End With
            SetScrollInfo hDlg, SB_HORZ, SI, True
            With SI
                .cbSize = Len(SI)
                .fMask = SIF_ALL
                .nMin = 0
                .nMax = RC.Bottom - PICTURESIZE
                .nPos = Center(RC.Bottom - PICTURESIZE, 0)
            End With
            SetScrollInfo hDlg, SB_VERT, SI, True
        Case WM_HSCROLL, WM_VSCROLL
            If LOWORD(wParam) = SB_THUMBPOSITION Or LOWORD(wParam) = SB_THUMBTRACK Then
                SetScrollPos hDlg, IIf(uMsg = WM_HSCROLL, SB_HORZ, SB_VERT), HIWORD(wParam), True
                MoveWindow GetDlgItem(hDlg, IDC_PICTURE1), GetScrollPos(hDlg, SB_HORZ), GetScrollPos(hDlg, SB_VERT), PICTURESIZE, PICTURESIZE, True
                InvalidateRect hDlg, Null, True
            ElseIf LOWORD(wParam) = SB_LINELEFT Or LOWORD(wParam) = SB_LINERIGHT Then
                SetScrollPos hDlg, IIf(uMsg = WM_HSCROLL, SB_HORZ, SB_VERT), GetScrollPos(hDlg, IIf(uMsg = WM_HSCROLL, SB_HORZ, SB_VERT)) + IIf(LOWORD(wParam) = SB_LINELEFT, -1, 1), True
                SendMessage hDlg, uMsg, MAKELONG(CInt(SB_THUMBPOSITION), GetScrollPos(hDlg, IIf(uMsg = WM_HSCROLL, SB_HORZ, SB_VERT))), ByVal 0&
            ElseIf LOWORD(wParam) = SB_PAGEUP Or LOWORD(wParam) = SB_PAGEDOWN Then
                SetScrollPos hDlg, IIf(uMsg = WM_HSCROLL, SB_HORZ, SB_VERT), GetScrollPos(hDlg, IIf(uMsg = WM_HSCROLL, SB_HORZ, SB_VERT)) + IIf(LOWORD(wParam) = SB_PAGEUP, -10, 10), True
                SendMessage hDlg, uMsg, MAKELONG(CInt(SB_THUMBPOSITION), GetScrollPos(hDlg, IIf(uMsg = WM_HSCROLL, SB_HORZ, SB_VERT))), ByVal 0&
            End If
        Case WM_MOUSEMOVE
            If wParam = MK_LBUTTON Then
                Dim X As Integer, Y As Integer
                X = LOWORD(lParam)
                Y = HIWORD(lParam)
                SendMessage hDlg, WM_HSCROLL, MAKELONG(CInt(SB_THUMBPOSITION), X), ByVal 0&
                SendMessage hDlg, WM_VSCROLL, MAKELONG(CInt(SB_THUMBPOSITION), Y), ByVal 0&
            End If
    End Select
    
    DialoProc = 0
    
End Function

Private Function GetAddr(ByVal lAddr As Long) As Long
    GetAddr = lAddr
End Function

Private Function Center(ByVal lMaxWidth As Long, ByVal lCurWidth As Long) As Long
    Center = (lMaxWidth - lCurWidth) / 2
End Function

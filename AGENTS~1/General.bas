Attribute VB_Name = "modGeneral"
Option Explicit
'
' Required Win32 API Declarations
'
' The AttachThreadInput function attaches the input processing
' mechanism of one thread to that of another thread.
' Windows created in different threads typically process input
' independently of each other - they have their own input states
' (focus, active, capture windows, key state, queue status, ...),
' and they are not synchronized with the input processing of other
' threads. By using the AttachThreadInput function, a thread can
' attach its input processing to another thread. This also allows
' threads to share their input states, so they can call the
' SetFocus function to set the keyboard focus to a window of a
' different thread. This also allows threads to get key-state
' information. These capabilities are not generally possible.
Public TaskbarOpen As Boolean
Public SubShown(0 To 1) As Boolean

Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Boolean
    Public Const EWX_FORCE = 4
    Public Const EWX_LOGOFF = 0
    Public Const EWX_REBOOT = 2
    Public Const EWX_SHUTDOWN = 1

Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRECT As RECT) As Long


Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRECT As RECT) As Long


Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long


Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long


Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
    Public Const RGN_AND = 1
    Public Const RGN_COPY = 5
    Public Const RGN_DIFF = 4
    Public Const RGN_OR = 2
    Public Const RGN_XOR = 3


Type POINTAPI
    X As Long
    Y As Long
    End Type


Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type
Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

'
' Constants used with APIs
'
Public Const SW_SHOW = 5
Public Const SW_RESTORE = 9
Public Const GW_OWNER = 4
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_APPWINDOW = &H40000
'
' Listbox messages
'
Public Const LB_ADDSTRING = &H180
Public Const LB_SETITEMDATA = &H19A

Declare Function AlphaBlending Lib "Alphablending.dll" _
                     (ByVal destHDC As Long, ByVal XDest As Long, ByVal YDest As Long, _
                      ByVal destWidth As Long, ByVal destHeight As Long, ByVal srcHDC As Long, _
                      ByVal xSrc As Long, ByVal ySrc As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal AlphaSource As Long) As Long
            

Public Sub pSetForegroundWindow(ByVal hwnd As Long)
Dim lForeThreadID As Long
Dim lThisThreadID As Long
Dim lReturn       As Long
'
' Make a window, specified by its handle (hwnd)
' the foreground window.
'
' If it is already the foreground window, exit.
'
If hwnd <> GetForegroundWindow() Then
    If IsIconic(hwnd) Then
       Call ShowWindow(hwnd, SW_RESTORE)
    Else
       Call ShowWindow(hwnd, SW_SHOW)
    End If
    '
    ' Get the threads for this window and the foreground window.
    '
    lForeThreadID = GetWindowThreadProcessId(GetForegroundWindow, ByVal 0&)
    lThisThreadID = GetWindowThreadProcessId(hwnd, ByVal 0&)
    '
    ' By sharing input state, threads share their concept of
    ' the active window.
    '
    If lForeThreadID <> lThisThreadID Then
        ' Attach the foreground thread to this window.
        Call AttachThreadInput(lForeThreadID, lThisThreadID, True)
        ' Make this window the foreground window.
        lReturn = SetForegroundWindow(hwnd)
        ' Detach the foreground window's thread from this window.
        Call AttachThreadInput(lForeThreadID, lThisThreadID, False)
    Else
       lReturn = SetForegroundWindow(hwnd)
    End If
    '
    ' Restore this window to its normal size.
    '

End If
End Sub

Public Function fEnumWindows(lst As ListBox) As Long
'
' Clear list, then fill it with the running
' tasks. Return the number of tasks.
'
' The EnumWindows function enumerates all top-level windows
' on the screen by passing the handle of each window, in turn,
' to an application-defined callback function. EnumWindows
' continues until the last top-level window is enumerated or
' the callback function returns FALSE.
'
With lst
    .Clear
    Call EnumWindows(AddressOf fEnumWindowsCallBack, .hwnd)
    fEnumWindows = .ListCount
End With
End Function

Private Function fEnumWindowsCallBack(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim lReturn     As Long
Dim lExStyle    As Long
Dim bNoOwner    As Boolean
Dim sWindowText As String
'
' This callback function is called by Windows (from
' the EnumWindows API call) for EVERY window that exists.
' It populates the listbox with a list of windows that we
' are interested in.
'
' Windows to display are those that:
'   -   are not this app's
'   -   are visible
'   -   do not have a parent
'   -   have no owner and are not Tool windows OR
'       have an owner and are App windows
'
If hwnd <> FrmTaskbar.hwnd And hwnd <> frmBstart.hwnd Then
    If IsWindowVisible(hwnd) Then
        If GetParent(hwnd) = 0 Then
            bNoOwner = (GetWindow(hwnd, GW_OWNER) = 0)
            lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
            
            If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or _
                ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
                '
                ' Get the window's caption.
                '
                sWindowText = Space$(256)
                lReturn = GetWindowText(hwnd, sWindowText, Len(sWindowText))
                If lReturn Then
                   '
                   ' Add it to our list.
                   '
                   sWindowText = Left$(sWindowText, lReturn)
                   lReturn = SendMessage(lParam, LB_ADDSTRING, 0, ByVal sWindowText)
                   Call SendMessage(lParam, LB_SETITEMDATA, lReturn, ByVal hwnd)
                End If
            End If
        End If
    End If
End If
fEnumWindowsCallBack = True
End Function


Public Sub Blend(Destination As Object, Source As Object, Amount As Integer, X, Y, X2, Y2)
AlphaBlending Destination.hdc, X, Y, X2, Y2, Source.hdc, X, Y, X2, Y2, Amount
End Sub

Public Sub HideStartMenu()
    Unload frmBstart
    HideSubs
    
    TaskbarOpen = False
    FrmTaskbar.picBstart.Picture = FrmTaskbar.picTbar.Picture
    Blend FrmTaskbar.picBstart, FrmTaskbar.picDesktopCapture, 80, 0, 0, FrmTaskbar.picBstart.Width, FrmTaskbar.picBstart.Height
    FrmTaskbar.picBstart.Refresh
End Sub

Public Sub ShowStartMenu()
    FrmTaskbar.picBstart.Picture = FrmTaskbar.picTbarDown.Picture
    Blend FrmTaskbar.picBstart, FrmTaskbar.picDesktopCapture, 80, 0, 0, FrmTaskbar.picBstart.Width, FrmTaskbar.picBstart.Height
    FrmTaskbar.picBstart.Refresh
    Load frmBstart
    TaskbarOpen = True
    frmBstart.showme
End Sub

Public Sub HideSubs()
            If SubShown(0) Then
                    Unload frmShutdownSubMenu
                    SubShown(0) = False
            End If
            If SubShown(1) Then
                    Unload frmHelpSubMenu
                    SubShown(1) = False
            End If
End Sub

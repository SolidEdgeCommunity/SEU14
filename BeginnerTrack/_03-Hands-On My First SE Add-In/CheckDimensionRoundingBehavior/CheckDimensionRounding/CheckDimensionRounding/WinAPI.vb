Imports System.Runtime.InteropServices

Module WinAPI

    Public MyResourceHandle As Int32
    Public MyResourceFilename As String

    Private Declare Function api_LoadString Lib "User32" Alias "LoadStringA" (ByVal hInstance As Int32, ByVal wID As Integer, ByVal lpBuffer As String, ByVal nBufferMax As Integer) As Integer
    Private Declare Function api_LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal H As IntPtr, ByVal flag As Int32) As Int32
    Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer

    <MarshalAsAttribute(UnmanagedType.ByValTStr, sizeconst:=512)> Private StringBuf As New String("", 512)

    Public Enum SetWindowPosFlags
        SWP_FRAMECHANGED = &H20S '  The frame changed: send WM_NCCALCSIZE
        SWP_HIDEWINDOW = &H80S
        SWP_NOACTIVATE = &H10S
        SWP_NOCOPYBITS = &H100S
        SWP_NOMOVE = &H2S
        SWP_NOOWNERZORDER = &H200S '  Don't do owner Z ordering
        SWP_NOREDRAW = &H8S
        SWP_NOSIZE = &H1S
        SWP_NOZORDER = &H4S
        SWP_SHOWWINDOW = &H40S
        SWP_DRAWFRAME = SetWindowPosFlags.SWP_FRAMECHANGED
        SWP_NOREPOSITION = SetWindowPosFlags.SWP_NOOWNERZORDER
    End Enum

    'Extended Window Messages
    Public Enum ExtendedWindowMessages
        GWL_EXSTYLE = (-20)
        GWL_HINSTANCE = (-6)
        GWL_HWNDPARENT = (-8)
        GWL_ID = (-12)
        GWL_STYLE = (-16)
        GWL_USERDATA = (-21)
        GWL_WNDPROC = (-4)
    End Enum

    Public Enum WindowStyles
        WS_BORDER = &H800000
        WS_CAPTION = &HC00000 '  WS_BORDER Or WS_DLGFRAME
        WS_CHILD = &H40000000
        WS_CLIPCHILDREN = &H2000000
        WS_CLIPSIBLINGS = &H4000000
        WS_DISABLED = &H8000000
        WS_DLGFRAME = &H400000
        WS_EX_ACCEPTFILES = &H10
        WS_EX_DLGMODALFRAME = &H1
        WS_EX_NOPARENTNOTIFY = &H4
        WS_EX_TOPMOST = &H8
        WS_EX_TRANSPARENT = &H20
        WS_GROUP = &H20000
        WS_HSCROLL = &H100000
        WS_MAXIMIZE = &H1000000
        WS_MAXIMIZEBOX = &H10000
        WS_MINIMIZE = &H20000000
        WS_MINIMIZEBOX = &H20000
        WS_OVERLAPPED = &H0
        WS_POPUP = &H80000000
        WS_SYSMENU = &H80000
        WS_TABSTOP = &H10000
        WS_THICKFRAME = &H40000
        WS_VISIBLE = &H10000000
        WS_VSCROLL = &H200000
        WS_CHILDWINDOW = (WindowStyles.WS_CHILD)
        WS_SIZEBOX = WindowStyles.WS_THICKFRAME
        WS_TILED = WindowStyles.WS_OVERLAPPED
        WS_ICONIC = WindowStyles.WS_MINIMIZE
        WS_POPUPWINDOW = (WindowStyles.WS_POPUP Or WindowStyles.WS_BORDER Or WindowStyles.WS_SYSMENU)
        WS_OVERLAPPEDWINDOW = (WindowStyles.WS_OVERLAPPED Or WindowStyles.WS_CAPTION Or WindowStyles.WS_SYSMENU Or WindowStyles.WS_THICKFRAME Or WindowStyles.WS_MINIMIZEBOX Or WindowStyles.WS_MAXIMIZEBOX)
        WS_TILEDWINDOW = WindowStyles.WS_OVERLAPPEDWINDOW
    End Enum
    Public Enum WindowMessages
        WM_COMMAND = &H111

    End Enum


    Public Sub ChangeParentWindow(ByVal hWndChild As Integer, ByVal hWndParent As Integer, ByVal bPopupStyle As Boolean)
        On Error Resume Next

        Dim dwStyle As Integer

        'Make sure we have valid hWnd's.
        If hWndChild = 0 Or hWndParent = 0 Then Exit Sub

        'Get\adjust the style settings of the child window.
        dwStyle = GetWindowLong(hWndChild, ExtendedWindowMessages.GWL_STYLE)
        dwStyle = dwStyle Or WindowStyles.WS_CHILD Or WindowStyles.WS_POPUP

        'Set window properties about the child window.
        SetWindowLong(hWndChild, ExtendedWindowMessages.GWL_STYLE, dwStyle)

        'Reparent the child window to the parent window.
        SetParent(hWndChild, hWndParent)

        'Set window properties about the child window.
        If Not bPopupStyle Then SetWindowLong(hWndChild, ExtendedWindowMessages.GWL_STYLE, dwStyle And Not WindowStyles.WS_POPUP)
    End Sub

    Public Sub SetResourceHandle(ByVal ResourceHandle As Int32)
        MyResourceHandle = ResourceHandle
    End Sub
    Public Sub SetResourceFilename(ByVal Name As String)
        MyResourceFilename = Name
    End Sub
    Public Function GetResourceHandle() As Int32
        GetResourceHandle = MyResourceHandle
    End Function
    Public Function GetResourceFilename() As String
        GetResourceFilename = MyResourceFilename
    End Function

    Public Function LoadResourceFile(ByVal Filename As String) As Int32
        LoadResourceFile = api_LoadLibraryEx(Filename, 0, &H2)
    End Function

    Public Function GetResourceString(ByVal ResourceId As Integer) As String

        GetResourceString = ""

        Dim ResHandle As Int32

        ResHandle = GetResourceHandle()
        If 0 <> ResHandle Then
            Dim RetVal As Integer
            RetVal = api_LoadString(ResHandle, ResourceId, StringBuf, 512)

            GetResourceString = StringBuf.Substring(0, Convert.ToInt32(RetVal))

        End If
    End Function
End Module

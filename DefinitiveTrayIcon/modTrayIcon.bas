Attribute VB_Name = "modTrayIcon"
'DEFINITIVE TRAY ICON v. 0.3
'By: Phil Ciebiera
'----------------------------
'If you like it give me a vote,
'-Phil









'API DECLARES
'--------------------------------
'***** BEGIN POPUP PROPER CODE *****'
'This is required so that when a popup menu is called the menu
'dismisses correctly if no item is chosen
Public Declare Function SetForegroundWindow Lib "user32" ( _
    ByVal hWnd As Long _
) As Long
'***** BEGIN POPUP PROPER CODE *****'

'***** BEGIN HOOKING CODE *****'
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long _
) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long
'***** END HOOKING CODE *****'

'***** BEGIN EXPLORER.EXE CRASH DETECTION CODE *****'
'This is used to find registered window messages, for explorer crash detection
'as well as to find the begining of program definable user messages ie: WM_APP
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" ( _
    ByVal lpString As String _
) As Long
'***** END EXPLORER.EXE CRASH DETECTION CODE *****'

'***** BEGIN TRAY ICON CODE *****'
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" ( _
    ByVal dwMessage As Long, _
    lpData As NOTIFYICONDATA _
) As Long
'***** END TRAY ICON CODE *****'

'***** BEGIN OTHER API CODE *****'
'This will be used to help generate a GUID for use with creating our
'WM_TRAYHOOK
Private Declare Function CoCreateGuid Lib "OLE32.DLL" ( _
    pGuid As GUID _
) As Long
'***** END OTHER API CODE *****'

'STRUCT DECLARES
'--------------------------------
'***** BEGIN TRAY ICON CODE *****'
Public Type NOTIFYICONDATA
    cbSize              As Long             'Size of NotifyIconData struct
    hWnd                As Long             'Window handle for the window handling the icon events
    uID                 As Long             'Icon ID (to allow multiple icons per application)
    uFlags              As Long             'NIF Flags
    uCallbackMessage    As Long             'The message received for the system tray icon
    hIcon               As Long             'The memory location of our icon if NIF_ICON is specifed
    szTip               As String * 128     'Tooltip if NIF_TIP is specified (64 characters max)
    dwState             As Long
    dwStateMask         As Long
    szInfo              As String * 256
    uTimeout            As Long
    szInfoTitle         As String * 64
    dwInfoFlags         As Long
End Type
'***** END TRAY ICON CODE *****'

'***** BEGIN BALLOON TIP CODE *****'
Public Enum BalloonIcon
    ICON_NONE = 0
    ICON_INFO = 1
    ICON_WARNING = 2
    ICON_ERROR = 3
    ICON_USER = 4
End Enum
'***** END BALLOON TIP CODE *****'

'***** BEGIN GUID HELPER *****'
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
'***** END GUID HELPER *****'



'VARIABLE DECLARES
'--------------------------------
'***** BEGIN HOOKING CODE *****'
'Public mWndProcNext As Long
Private bIsHooked As Boolean
'***** END HOOKING CODE *****'

'***** BEGIN TRAY ICON CODE *****'
Private TrayIcon As NOTIFYICONDATA

'used to indentify different tray icons if used
Public WM_APP As Long  'For user defined window messages
Public WM_TRAYHOOK As Long 'The tray icon window message
'***** END TRAY ICON CODE *****'

'***** BEGIN EXPLORER.EXE CRASH DETECTION CODE *****'
Public mTaskbarCreated As Long
'***** END EXPLORER.EXE CRASH DETECTION CODE *****'


'CONSTANT DECLARES
'--------------------------------
'***** BEGIN HOOKING CODE *****'
Private Const GWL_WNDPROC = (-4)
Private Const GWL_USERDATA = (-21)

'Window messages relating to balloon tips and the like branch from here
Private Const WM_USER As Long = &H400
'***** END HOOKING CODE *****'

'***** BEGIN TRAY ICON CODE *****'
'Here are some mouse "events" to play with
'we are only going to use two in our example,
'however you feel free to use whatever you'd like (:

'Left Button
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_LBUTTONDBLCLK As Long = &H203
' Middle Button
Public Const WM_MBUTTONDOWN As Long = &H207
Public Const WM_MBUTTONUP As Long = &H208
Public Const WM_MBUTTONDBLCLK As Long = &H209
' Right Button
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_RBUTTONDBLCLK As Long = &H206

' Shell_NotifyIconA() messages
Private Const NIM_ADD As Long = &H0     'Add icon to the System Tray
Private Const NIM_MODIFY As Long = &H1  'Modify System Tray icon
Private Const NIM_DELETE As Long = &H2  'Delete icon from System Tray

'NotifyIconData Flags
Private Const NIF_MESSAGE As Long = &H1 'uCallbackMessage in NOTIFYICONDATA is valid
Private Const NIF_ICON As Long = &H2 ' hIcon in NOTIFYICONDATA is valid
Private Const NIF_TIP As Long = &H4 'szTip in NOTIFYICONDATA is valid
Private Const NIF_INFO As Long = &H10 'for use with balloons
'***** END TRAY ICON CODE *****'

'***** BEGIN BALLOON TIP CODE *****'
'Balloon tip icon constants
Private Const NIIF_NONE As Long = &H0
Private Const NIIF_WARNING As Long = &H2
Private Const NIIF_ERROR As Long = &H3
Private Const NIIF_USER As Long = &H4
Private Const NIIF_INFO As Long = &H1

'Balloon tip sound constants
Private Const NIIF_NOSOUND As Long = &H10

'Balloon tip notification messages
Public Const NIN_BALLOONSHOW As Long = WM_USER + &H2 'when the balloon is drawn
Public Const NIN_BALLOONHIDE As Long = WM_USER + &H3 'when the balloon disappears—for example, when the icon is deleted. This message is not sent if the balloon is dismissed because of a timeout or a mouse click.
Public Const NIN_BALLOONTIMEOUT As Long = WM_USER + &H4 'when the balloon is dismissed because of a timeout
Public Const NIN_BALLOONUSERCLICK As Long = WM_USER + &H5 'when the balloon is dismissed because of a mouse click.
'***** END BALLOON TIP CODE *****'



'Always
Option Explicit


'***** BEGIN TRAY ICON CODE *****'

Public Function CreateTrayIcon( _
    ByRef Owner As Form, _
    ByVal luID As Long, _
    Optional ByRef ToolTip As String = "", _
    Optional ByRef tIcon As StdPicture _
) As Long
    
    With TrayIcon
        .cbSize = Len(TrayIcon) 'This size is always the len(NOTIFYICONDATA)
        .hWnd = Owner.hWnd  'Which form is this icon for
        .uID = luID
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE    'set valid data inputs
        
        'You see this is where most VB tray icon codes are bad, no offense
        'they use a hack that uses the message WM_MOUSEMOVE for notification
        'this way VB can handle the message using its built in
        'Form_MouseMove event. This is not the way to do it in my opinion
        'because you can't have multiple icons for one form, my method allows
        'for this, not to mention the inability to detect explorer.exe crashes
        'but using the hack method you don't need to use message hooking either...
        'What window message should be sent during an event
        .uCallbackMessage = WM_TRAYHOOK
        .szTip = Trim(ToolTip$) & vbNullChar   'set the tooltip
        If tIcon Is Nothing Then
            .hIcon = Owner.Icon
        Else
            .hIcon = tIcon
        End If

        
    End With
    'Create the tray icon with an API call
    CreateTrayIcon = Shell_NotifyIcon(NIM_ADD, TrayIcon)
End Function

Public Function ModifyTrayIcon( _
    ByRef Owner As Form, _
    ByVal luID As Long, _
    Optional ByRef ToolTip As String = "", _
    Optional ByRef tIcon As StdPicture _
) As Long
    
    With TrayIcon
        .cbSize = Len(TrayIcon)
        .hWnd = Owner.hWnd
        .uID = luID
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_TRAYHOOK
        If Not tIcon Is Nothing Then
            .hIcon = tIcon
        End If
        If ToolTip <> "" Then .szTip = Trim(ToolTip$) & vbNullChar
    End With
    'Update the tray icon with an API call
    ModifyTrayIcon = Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
End Function

Public Function DeleteTrayIcon( _
    ByVal luID As Long _
) As Long
    With TrayIcon
        .cbSize = Len(TrayIcon)
        .uID = luID
        .uFlags = NIM_DELETE
        .uCallbackMessage = WM_TRAYHOOK
    End With
    
    'Remove the tray icon with an API call
    DeleteTrayIcon = Shell_NotifyIcon(NIM_DELETE, TrayIcon)
End Function


'***** END TRAY ICON CODE *****'


'***** BEGIN MESSAGE HOOKING CODE *****'

Public Function InsertHook( _
    ByRef Owner As Form _
) As Long
    
    Dim lResult As Long
    
    'Remove preexisting hook
    'Call RemoveHook(Owner)
    
    InsertHook = SetWindowLong(Owner.hWnd, GWL_WNDPROC, AddressOf GlobalMessageCatcher)
    If InsertHook Then
        lResult = SetWindowLong(Owner.hWnd, GWL_USERDATA, ObjPtr(Owner))
        'bIsHooked = True
    End If
End Function

Public Sub RemoveHook( _
    ByRef Owner As Form, _
    ByVal lHookID As Long _
)
    'Remove the hook and revert control back to VB
    
    If lHookID Then    'Make sure we really are hooked
        SetWindowLong Owner.hWnd, GWL_WNDPROC, lHookID
        'bIsHooked = False
    End If
End Sub

Public Function GlobalMessageCatcher( _
    ByVal shWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

    'All this serves is to intercept window messages and pass them to our
    'message handling function for inspection before then being ultimately
    'passed back to VB, this has been modified as of v. 0.3 for use across
    'multiple forms
    
    If shWnd = frmSingle.hWnd Then
        GlobalMessageCatcher = frmSingle.WindowProc(shWnd, uMsg, wParam, lParam)
    ElseIf shWnd = frmMultiple.hWnd Then
        GlobalMessageCatcher = frmMultiple.WindowProc(shWnd, uMsg, wParam, lParam)
    End If
    
    
    'If wParam = 111& Then
    '    GlobalMessageCatcher = frmSingle.WindowProc(shWnd, uMsg, wParam, lParam)
    'Else
    '    GlobalMessageCatcher = CallWindowProc(mWndProcNext, shWnd, uMsg, wParam, lParam)
    'End If
    'ElseIf wParam = 112& Or wParam = 113& Or wParam = 114& Then
        'GlobalMessageCatcher = frmMultiple.WindowProc(shWnd, uMsg, wParam, lParam)

End Function

'***** END MESSAGE HOOKING CODE *****'


'***** BEGIN BALLOON TIP CODE *****'
Public Function PopupBalloon( _
    ByRef Owner As Form, _
    ByVal luID As Long, _
    ByRef Message As String, _
    ByRef Title As String, _
    Optional ByVal IconType As BalloonIcon = ICON_INFO, _
    Optional ByVal Sound As Boolean = True, _
    Optional ByRef tIcon As StdPicture _
) As Long


    'This line is optional, if you include it new balloon tips erase old ones
    'if you omit it a balloon tip queue so to speak is created, and as they timeout
    'new ones appear
    Call RemoveBalloon(Owner, luID)
    With TrayIcon
        .cbSize = Len(TrayIcon)
        .hWnd = Owner.hWnd
        .uID = luID
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIM_MODIFY
        .uCallbackMessage = WM_TRAYHOOK
        If tIcon Is Nothing Then
            .hIcon = Owner.Icon
        Else
            .hIcon = tIcon
        End If
        .dwState = 0
        .dwStateMask = 0
        .szInfo = Message & Chr(0)
        .szInfoTitle = Title & Chr(0)
        Select Case IconType
            Case ICON_NONE
                .dwInfoFlags = NIIF_NONE
            Case ICON_INFO
                .dwInfoFlags = NIIF_INFO
            Case ICON_WARNING
                .dwInfoFlags = NIIF_WARNING
            Case ICON_ERROR
                .dwInfoFlags = NIIF_ERROR
            Case ICON_USER
                .dwInfoFlags = NIIF_USER
        End Select
        If Not Sound Then .dwInfoFlags = .dwInfoFlags Or NIIF_NOSOUND
    End With
    
    PopupBalloon = Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
End Function

'This function removes an existing Balloon Tip
Public Function RemoveBalloon( _
    ByRef Owner As Form, _
    ByVal luID As Long _
) As Long

    With TrayIcon
        .cbSize = Len(TrayIcon)
        .hWnd = Owner.hWnd
        .uID = luID
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIM_MODIFY
        .uCallbackMessage = WM_TRAYHOOK
        .hIcon = Owner.Icon
        .dwState = 0
        .dwStateMask = 0
        .szInfo = Chr(0)
        .szInfoTitle = Chr(0)
        .dwInfoFlags = NIIF_NONE
    End With
    RemoveBalloon = Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
End Function
'***** END BALLOON TIP CODE *****'


'***** BEGIN GUID GENERATION CODE *****'
Public Function GetGUID() As String
    '(c) 2000 Gus Molina
    '*** SOURCE FROM MSDN (http://support.microsoft.com/kb/176790/EN-US/) ***
    Dim udtGUID As GUID

    If (CoCreateGuid(udtGUID) = 0) Then
        GetGUID = _
        String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
        String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
        String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
        IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
        IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
        IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
        IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
        IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
        IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
        IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
        IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If
End Function
'***** END GUID GENERATION CODE *****'





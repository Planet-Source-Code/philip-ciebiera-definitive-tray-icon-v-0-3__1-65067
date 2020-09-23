VERSION 5.00
Begin VB.Form frmSingle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definitive Tray Icon v. 0.3 - [Single Icon]"
   ClientHeight    =   8040
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      Height          =   1815
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   6000
      Width           =   8535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "OneForm_OneTrayIcon.frx":0000
      Top             =   120
      Width           =   8535
   End
   Begin VB.Label lblMessages 
      Caption         =   "Balloon Notifications:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   5760
      Width           =   3375
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Choices"
      Begin VB.Menu mnuSound 
         Caption         =   "Sound"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNoIconBalloon 
         Caption         =   "Test No Icon Balloon Tip"
      End
      Begin VB.Menu mnuInfoBalloon 
         Caption         =   "Test Info Balloon Tip"
      End
      Begin VB.Menu mnuWarningBalloon 
         Caption         =   "Test Warning Balloon Tip"
      End
      Begin VB.Menu mnuErrorBalloon 
         Caption         =   "Test Error Balloon Tip"
      End
      Begin VB.Menu mnuUserBalloon 
         Caption         =   "Test User Balloon Tip"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMultiIcon 
         Caption         =   "Multiple Icon Demo"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmSingle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'API DECLARES
'--------------------------------




'VARIABLE DECLARES
'--------------------------------

Private lHookID As Long


'CONSTANT DECLARES
'--------------------------------




'***** BEGIN SAMPLE PROGRAM RELATED SECTION *****'
Private bSound As Boolean

Option Explicit

Private Sub Form_Load()
    PrintInfo   'Output my message in the first text box
    bSound = True   'Set sounds to be enabled by default
    
    
    'This gets us a globaly unique ID so that we can be sure the message
    'we use for getting our programs messages is unique
    WM_TRAYHOOK = RegisterWindowMessage(GetGUID())
    
    'This retrieves the window message for when the taskbar is created
    'since usually the application is run after the taskbar is created
    'it is safe to assume that if your program receives this message
    'any icon in the tray that was there is now gone and needs to be
    'recreated with a call to Shell_NotifyIcon(NIM_ADD, x)
    mTaskbarCreated = RegisterWindowMessage("TaskbarCreated")
    
    CreateTrayIcon Me, 111&, "Sample ToolTip", Me.Icon 'Create the tray icon
    lHookID = InsertHook(Me)   'Start the message hook
    PopupBalloon Me, 111&, "Program running, hook installed.", App.Title, ICON_INFO, bSound
End Sub

Private Sub Form_Resize()
    'Catch the window being minimized and react to it.
    If Me.WindowState = vbMinimized Then
        ModifyTrayIcon Me, 111&, "Double Click to restore window."
        PopupBalloon Me, 111&, App.Title & " is still running!" & vbCrLf & vbCrLf & "Right Click > Exit to end" & vbCrLf & vbCrLf & "-OR-" & vbCrLf & vbCrLf & "Double Click to restore.", App.Title, ICON_INFO, bSound
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteTrayIcon 111&  'Remove the tray icon
    RemoveHook Me, lHookID   'Remove the message hook  <=!!!IMPORTANT!!!
End Sub

Friend Function WindowProc( _
    ByVal shWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

    'This is our message handler
    
    If shWnd = Me.hWnd Then 'First we check to see if the message is for this window
        Select Case uMsg    'Then we look at the message
            Case mTaskbarCreated    'This message is for when the taskbar is created
                'if the taskbar was created, chances are explorer.exe had crashed
                CreateTrayIcon Me, 111&, "Sample ToolTip", Me.Icon 'recreate the tray icon
                PopupBalloon Me, 111&, "Detected explorer.exe crash!", App.Title, ICON_WARNING, bSound
            
            

            
            
            Case WM_TRAYHOOK 'Our user defined window message
                'if we get this we know that lParam carries the "event"
                'that occured on the tray icon
                

                
                Select Case lParam
                    Case WM_LBUTTONDBLCLK   'Left button dbl clicked
                        If Me.WindowState = vbMinimized Then
                            ModifyTrayIcon Me, 111&, "Sample ToolTip"
                            SetForegroundWindow Me.hWnd
                            PopupBalloon Me, 111&, "Program Restored.", App.Title, ICON_INFO, bSound
                            Me.WindowState = vbNormal
                            Me.Show
                            Me.SetFocus
                        End If
                        'Me.SetFocus
                        
                    
                    Case WM_RBUTTONUP   'Right button released
                        SetForegroundWindow Me.hWnd
                        RemoveBalloon Me, 111&
                        PopupMenu Me.mnuPopup
                    
                    Case NIN_BALLOONUSERCLICK
                        AppendLog "-User clicked the balloon."
                        
                    Case NIN_BALLOONTIMEOUT
                        AppendLog "-Balloon disapeared floated away, or was dismissed."
                End Select
        
        End Select
    End If
    
    'also pass them to VB
    WindowProc = CallWindowProc(lHookID, shWnd, uMsg, wParam, lParam)
End Function

Private Sub mnuErrorBalloon_Click()
    PopupBalloon Me, 111&, "Sample Error Balloon.", App.Title, ICON_ERROR, bSound
End Sub

Private Sub mnuInfoBalloon_Click()
    PopupBalloon Me, 111&, "Sample Info Balloon.", App.Title, ICON_INFO, bSound
End Sub

Private Sub mnuMultiIcon_Click()
    frmMultiple.Show
End Sub

Private Sub mnuNoIconBalloon_Click()
    PopupBalloon Me, 111&, "Sample No Icon Balloon.", App.Title, ICON_NONE, bSound
End Sub

Private Sub mnuSound_Click()
    If mnuSound.Checked Then
        mnuSound.Checked = False
    Else
        mnuSound.Checked = True
    End If
    bSound = mnuSound.Checked
End Sub

Private Sub mnuUserBalloon_Click()
    PopupBalloon Me, 111&, "Sample User Balloon.", App.Title, ICON_USER, bSound
End Sub

Private Sub mnuWarningBalloon_Click()
    PopupBalloon Me, 111&, "Sample Warning Balloon.", App.Title, ICON_WARNING, bSound
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub


















Private Sub PrintInfo()
    Dim sText As String
    
    sText = _
        "This code demonstrates what I believe is the RIGHT WAY to utilize the System Tray Icon in VB.  " & _
        "It uses a form of subclassing to install a message hook to intercept window messages that are " & _
        "usually handled by VB through " & Chr(34) & "events" & Chr(34) & ".  Using my method gives a greater level " & _
        "of control over your tray icon.  For instance my method can detect when the TaskBar is recreated " & _
        "due to an explorer.exe crash, and react to that by recreating the icon in the system tray.  " & _
        "Most VB sample code does not do this.  I've also included native windows Balloon Tip support with " & _
        "multiple icons, this portion of the code has been modified from Mark Mokoski's example on PSC.  " & _
        "I'd like to thank him for showing us an excellent way to achive this.  Some things to watch out " & _
        "for are: subclassing in VB can cause the IDE to crash if you're not careful when unhooking, or " & _
        "debuging. SO BE WARNED SAVE YOUR CODE ALWAYS BEFORE RUNNING!!! I CAN'T STRESS THIS ENOUGH.    " & _
        "Have fun with it   ( :    VOTE IF YOU LIKE IT" & vbCrLf & vbCrLf & _
        "Things to try:" & vbCrLf & _
        "1.) Minimize the window." & vbCrLf & _
        "2.) Right Click the tray icon and play." & vbCrLf & _
        "3.) Ctrl + Alt + Del, find and end EXPLORER.EXE (if needed go to File > New Task and type explorer.exe)" & vbCrLf & _
        "4.) Click on balloons when they appear, not on the X" & vbCrLf & vbCrLf & _
        "-Phil"
    
    Text1.Text = sText
End Sub

Private Sub AppendLog( _
    ByRef sText As String _
)
    txtLog.Text = txtLog.Text & sText & vbCrLf
End Sub


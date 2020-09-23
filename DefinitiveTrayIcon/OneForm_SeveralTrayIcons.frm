VERSION 5.00
Begin VB.Form frmMultiple 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definitive Tray Icon v. 0.3 - [Multiple Icons]"
   ClientHeight    =   4485
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   8055
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   2
      Left            =   6720
      Picture         =   "OneForm_SeveralTrayIcons.frx":0000
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   1
      Left            =   6360
      Picture         =   "OneForm_SeveralTrayIcons.frx":014A
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   0
      Left            =   6000
      Picture         =   "OneForm_SeveralTrayIcons.frx":0294
      Top             =   0
      Width           =   240
   End
   Begin VB.Menu mnuIcon1 
      Caption         =   "Icon 1"
      Begin VB.Menu mnuSound1 
         Caption         =   "Sound"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNoIconBalloon1 
         Caption         =   "Test No Icon Balloon Tip"
      End
      Begin VB.Menu mnuInfoBalloon1 
         Caption         =   "Test Info Balloon Tip"
      End
      Begin VB.Menu mnuWarningBalloon1 
         Caption         =   "Test Warning Balloon Tip"
      End
      Begin VB.Menu mnuErrorBalloon1 
         Caption         =   "Test Error Balloon Tip"
      End
      Begin VB.Menu mnuUserBalloon1 
         Caption         =   "Test User Balloon Tip"
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit1 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuIcon2 
      Caption         =   "Icon 2"
      Begin VB.Menu mnuSound2 
         Caption         =   "Sound"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNoIconBalloon2 
         Caption         =   "Test No Icon Balloon Tip"
      End
      Begin VB.Menu mnuInfoBalloon2 
         Caption         =   "Test Info Balloon Tip"
      End
      Begin VB.Menu mnuWarningBalloon2 
         Caption         =   "Test Warning Balloon Tip"
      End
      Begin VB.Menu mnuErrorBalloon2 
         Caption         =   "Test Error Balloon Tip"
      End
      Begin VB.Menu mnuUserBalloon2 
         Caption         =   "Test User Balloon Tip"
      End
      Begin VB.Menu mnuSep22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit2 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuIcon3 
      Caption         =   "Icon 3"
      Begin VB.Menu mnuSound3 
         Caption         =   "Sound"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep31 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNoIconBalloon3 
         Caption         =   "Test No Icon Balloon Tip"
      End
      Begin VB.Menu mnuInfoBalloon3 
         Caption         =   "Test Info Balloon Tip"
      End
      Begin VB.Menu mnuWarningBalloon3 
         Caption         =   "Test Warning Balloon Tip"
      End
      Begin VB.Menu mnuErrorBalloon3 
         Caption         =   "Test Error Balloon Tip"
      End
      Begin VB.Menu mnuUserBalloon3 
         Caption         =   "Test User Balloon Tip"
      End
      Begin VB.Menu mnuSep32 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit3 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lHookID As Long

'***** BEGIN SAMPLE FORM ORGANIZATIONAL CODE *****'
Private Type MyTrayIcon
    ToolTip As String
    Icon As StdPicture
    uID As Long
    Sound As Boolean
End Type
Private aTrayIcons(2) As MyTrayIcon
'***** END SAMPLE FORM ORGANIZATIONAL CODE *****'



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
                Call CreateIcons
                PopupBalloon Me, aTrayIcons(1).uID, "Detected explorer.exe crash!", App.Title, ICON_WARNING, aTrayIcons(1).Sound, aTrayIcons(1).Icon
            
            
            Case WM_TRAYHOOK 'Our user defined window message
                'if we get this we know that lParam carries the "event"
                'that occured on the tray icon
                
                Select Case lParam
                    Case WM_LBUTTONDBLCLK   'Left button dbl clicked
                        If Me.WindowState = vbMinimized Then
                            Select Case wParam
                                Case aTrayIcons(0).uID
                                    ModifyTrayIcon Me, aTrayIcons(0).uID, aTrayIcons(0).ToolTip, aTrayIcons(0).Icon
                                    SetForegroundWindow Me.hWnd
                                    PopupBalloon Me, aTrayIcons(0).uID, "Program Restored.", App.Title, ICON_INFO, aTrayIcons(0).Sound, aTrayIcons(0).Icon
                                    Me.WindowState = vbNormal
                                    Me.Show
                                    Me.SetFocus
                                Case aTrayIcons(1).uID
                                    ModifyTrayIcon Me, aTrayIcons(1).uID, aTrayIcons(1).ToolTip, aTrayIcons(1).Icon
                                    SetForegroundWindow Me.hWnd
                                    PopupBalloon Me, aTrayIcons(1).uID, "Program Restored.", App.Title, ICON_INFO, aTrayIcons(1).Sound, aTrayIcons(1).Icon
                                    Me.WindowState = vbNormal
                                    Me.Show
                                    Me.SetFocus
                                Case aTrayIcons(2).uID
                                    ModifyTrayIcon Me, aTrayIcons(2).uID, aTrayIcons(2).ToolTip, aTrayIcons(2).Icon
                                    SetForegroundWindow Me.hWnd
                                    PopupBalloon Me, aTrayIcons(2).uID, "Program Restored.", App.Title, ICON_INFO, aTrayIcons(2).Sound, aTrayIcons(2).Icon
                                    Me.WindowState = vbNormal
                                    Me.Show
                                    Me.SetFocus
                            End Select
                        End If
                    
                    Case WM_RBUTTONUP   'Right button released
                        Select Case wParam
                            Case aTrayIcons(0).uID
                                SetForegroundWindow Me.hWnd
                                PopupMenu Me.mnuIcon1
                            Case aTrayIcons(1).uID
                                SetForegroundWindow Me.hWnd
                                PopupMenu Me.mnuIcon2
                            Case aTrayIcons(2).uID
                                SetForegroundWindow Me.hWnd
                                PopupMenu Me.mnuIcon3
                        End Select
                            
                    Case NIN_BALLOONUSERCLICK
                        AppendLog "-User clicked the balloon, with uID (" & CStr(wParam) & ")."
                        
                    Case NIN_BALLOONTIMEOUT
                        AppendLog "-Balloon with uID (" & CStr(wParam) & "), floated away, or was dismissed."
                End Select
        End Select
    End If
    
    'also pass them to VB
    WindowProc = CallWindowProc(lHookID, shWnd, uMsg, wParam, lParam)
End Function



'This just sets up some organizational stuff for the multi icon sample
Private Sub LoadIcons()
    Dim i As Byte
    For i = 0 To 2
        With aTrayIcons(i)
            .ToolTip = "Tray Icon " & CStr(i + 1)
            .uID = 112& + CLng(i)
            Set .Icon = imgIcon(i).Picture
            .Sound = True
        End With
    Next i
End Sub

Private Sub CreateIcons()
    CreateTrayIcon Me, aTrayIcons(0).uID, aTrayIcons(0).ToolTip, aTrayIcons(0).Icon
    CreateTrayIcon Me, aTrayIcons(1).uID, aTrayIcons(1).ToolTip, aTrayIcons(1).Icon
    CreateTrayIcon Me, aTrayIcons(2).uID, aTrayIcons(2).ToolTip, aTrayIcons(2).Icon
End Sub

Private Sub Form_Load()
    lHookID = InsertHook(Me)
    Call LoadIcons
    Call CreateIcons
End Sub


Private Sub AppendLog( _
    ByRef sText As String _
)
    txtLog.Text = txtLog.Text & sText & vbCrLf
End Sub







Private Sub Form_Resize()
    'Catch the window being minimized and react to it.
    If Me.WindowState = vbMinimized Then
        ModifyTrayIcon Me, aTrayIcons(1).uID, "Double Click to restore window.", aTrayIcons(1).Icon
        PopupBalloon Me, aTrayIcons(1).uID, App.Title & " is still running!" & vbCrLf & vbCrLf & "Right Click > Exit to end" & vbCrLf & vbCrLf & "-OR-" & vbCrLf & vbCrLf & "Double Click to restore.", App.Title, ICON_INFO, aTrayIcons(1).Sound, aTrayIcons(1).Icon
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RemoveHook Me, lHookID
    DeleteTrayIcon aTrayIcons(0).uID
    DeleteTrayIcon aTrayIcons(1).uID
    DeleteTrayIcon aTrayIcons(2).uID
    Set aTrayIcons(0).Icon = Nothing
    Set aTrayIcons(1).Icon = Nothing
    Set aTrayIcons(2).Icon = Nothing
End Sub

Private Sub mnuSound1_Click()
    If mnuSound1.Checked Then
        mnuSound1.Checked = 0
    Else
        mnuSound1.Checked = 1
    End If
    aTrayIcons(0).Sound = mnuSound1.Checked
End Sub

Private Sub mnuSound2_Click()
    If mnuSound2.Checked Then
        mnuSound2.Checked = 0
    Else
        mnuSound2.Checked = 1
    End If
    aTrayIcons(1).Sound = mnuSound2.Checked

End Sub

Private Sub mnuSound3_Click()
    If mnuSound3.Checked Then
        mnuSound3.Checked = 0
    Else
        mnuSound3.Checked = 1
    End If
    aTrayIcons(2).Sound = mnuSound3.Checked

End Sub






Private Sub mnuNoIconBalloon1_Click()
    PopupBalloon Me, aTrayIcons(0).uID, "Sample No Icon Balloon", App.Title, ICON_NONE, aTrayIcons(0).Sound, aTrayIcons(0).Icon
End Sub
Private Sub mnuNoIconBalloon2_Click()
    PopupBalloon Me, aTrayIcons(1).uID, "Sample No Icon Balloon", App.Title, ICON_NONE, aTrayIcons(1).Sound, aTrayIcons(1).Icon
End Sub
Private Sub mnuNoIconBalloon3_Click()
    PopupBalloon Me, aTrayIcons(2).uID, "Sample No Icon Balloon", App.Title, ICON_NONE, aTrayIcons(2).Sound, aTrayIcons(2).Icon
End Sub
'**********'
Private Sub mnuInfoBalloon1_Click()
    PopupBalloon Me, aTrayIcons(0).uID, "Sample Info Balloon", App.Title, ICON_INFO, aTrayIcons(0).Sound, aTrayIcons(0).Icon
End Sub
Private Sub mnuInfoBalloon2_Click()
    PopupBalloon Me, aTrayIcons(1).uID, "Sample Info Balloon", App.Title, ICON_INFO, aTrayIcons(1).Sound, aTrayIcons(1).Icon
End Sub
Private Sub mnuInfoBalloon3_Click()
    PopupBalloon Me, aTrayIcons(2).uID, "Sample Info Balloon", App.Title, ICON_INFO, aTrayIcons(2).Sound, aTrayIcons(2).Icon
End Sub
'**********'
Private Sub mnuWarningBalloon1_Click()
    PopupBalloon Me, aTrayIcons(0).uID, "Sample Warning Balloon", App.Title, ICON_WARNING, aTrayIcons(0).Sound, aTrayIcons(0).Icon
End Sub
Private Sub mnuWarningBalloon2_Click()
    PopupBalloon Me, aTrayIcons(1).uID, "Sample Warning Balloon", App.Title, ICON_WARNING, aTrayIcons(1).Sound, aTrayIcons(1).Icon
End Sub
Private Sub mnuWarningBalloon3_Click()
    PopupBalloon Me, aTrayIcons(2).uID, "Sample Warning Balloon", App.Title, ICON_WARNING, aTrayIcons(2).Sound, aTrayIcons(2).Icon
End Sub
'**********'
Private Sub mnuErrorBalloon1_Click()
    PopupBalloon Me, aTrayIcons(0).uID, "Sample Error Balloon", App.Title, ICON_ERROR, aTrayIcons(0).Sound, aTrayIcons(0).Icon
End Sub
Private Sub mnuErrorBalloon2_Click()
    PopupBalloon Me, aTrayIcons(1).uID, "Sample Error Balloon", App.Title, ICON_ERROR, aTrayIcons(1).Sound, aTrayIcons(1).Icon
End Sub
Private Sub mnuErrorBalloon3_Click()
    PopupBalloon Me, aTrayIcons(2).uID, "Sample Error Balloon", App.Title, ICON_ERROR, aTrayIcons(2).Sound, aTrayIcons(2).Icon
End Sub
'**********'
Private Sub mnuUserBalloon1_Click()
    PopupBalloon Me, aTrayIcons(0).uID, "Sample User Balloon", App.Title, ICON_USER, aTrayIcons(0).Sound, aTrayIcons(0).Icon
End Sub
Private Sub mnuUserBalloon2_Click()
    PopupBalloon Me, aTrayIcons(1).uID, "Sample User Balloon", App.Title, ICON_USER, aTrayIcons(1).Sound, aTrayIcons(1).Icon
End Sub
Private Sub mnuUserBalloon3_Click()
    PopupBalloon Me, aTrayIcons(2).uID, "Sample User Balloon", App.Title, ICON_USER, aTrayIcons(2).Sound, aTrayIcons(2).Icon
End Sub
'**********'
Private Sub mnuExit1_Click()
    Unload Me
End Sub
Private Sub mnuExit2_Click()
    Unload Me
End Sub
Private Sub mnuExit3_Click()
    Unload Me
End Sub


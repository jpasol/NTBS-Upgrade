VERSION 5.00
Begin VB.Form frmLogIn 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Log On"
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLogOn 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CheckBox chkNTLog 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&NT Log On"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtLogOn 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtLogOn 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtLogOn 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtLogOn 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblLogOn 
      BackStyle       =   0  'Transparent
      Caption         =   "Supervisor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Shape shpLogOn 
      BorderColor     =   &H00808080&
      Height          =   375
      Left            =   6480
      Top             =   45
      Width           =   465
   End
   Begin VB.Label lblMainStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Message Line"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   360
      TabIndex        =   12
      Top             =   3120
      Width           =   5115
   End
   Begin VB.Label lblLogExit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4440
      MouseIcon       =   "LogIn.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblLogOk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4440
      MouseIcon       =   "LogIn.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblLogOn 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblLogOn 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblLogOnTitle 
      BackColor       =   &H00000000&
      Caption         =   " Log On"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      MousePointer    =   15  'Size All
      TabIndex        =   11
      Top             =   0
      Width           =   6315
   End
   Begin VB.Label lblLogOn 
      BackStyle       =   0  'Transparent
      Caption         =   "Database"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblLogOn 
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================
' Project       : Bill2K
' Module        : Main
' Form          : frmLogIn
' Decription    : Main Log On Screen
' Author        : BOS
' Created       : Jan 13, 1999
' Revision      :
'=============================================

Option Explicit

' Constants
Const lcServer = 0
Const lcDatabase = 1
Const lcUserID = 2
Const lcPassword = 3
Const lcCopyright = "MIS Copyright © 1999"

Private mrcTitleBar As RECT
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'API declarations
Private Type RECT
    left As Long
    tOp As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40

Dim hWnd1 As Long

'---------------------------------------------
' Trigger Standard or NT Trusted connection
'---------------------------------------------
Private Sub chkNTlog_Click()
    txtLogOn(lcUserID).Text = zCurrentUser()
    txtLogOn(lcPassword).Text = ""
    If chkNTLog.Value = 1 Then
        txtLogOn(lcUserID).Locked = True
        txtLogOn(lcPassword).Locked = True
    Else
        txtLogOn(lcUserID).Locked = False
        txtLogOn(lcPassword).Locked = False
        txtLogOn(lcUserID).SetFocus
        txtLogOn(lcUserID).SelStart = 0
        txtLogOn(lcUserID).SelLength = txtLogOn(lcUserID).MaxLength
    End If
    txtLogOn(lcUserID).TabStop = Not txtLogOn(lcUserID).Locked
    txtLogOn(lcPassword).TabStop = Not txtLogOn(lcPassword).Locked
End Sub

'---------------------------------------------
' Act on Enter / Esc keypress
'---------------------------------------------
Private Sub chkNTLog_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            lConnectUser
        Case vbKeyEscape
            End
    End Select
End Sub

'---------------------------------------------
' Initial Form Load
'---------------------------------------------
Private Sub Form_Load()
    'Call lzHideTaskbar
    With shpLogOn
        .left = Me.left: .tOp = Me.tOp
        .Width = Me.Width: .Height = Me.Height
    End With
    'Call zShowInTaskbar(True, Me.hwnd)  - not working in windows2000 environment
    ' display copyright notice
    lblMainStatus.Caption = lcCopyright
    ' get INI values
    zGetINIVal ("")
    txtLogOn(lcServer) = Trim(gINIServer)
    txtLogOn(lcDatabase) = Trim(gINIDatabase)
    ' default to NT trusted connection
    txtLogOn(lcUserID).Text = zCurrentUser()
    txtLogOn(lcPassword).Text = ""
    gComputer = zCurrentComputer()
    gShutDown = False
    ' refresh form
    With mrcTitleBar
        .left = lblLogOnTitle.left
        .tOp = lblLogOnTitle.tOp
        .Right = lblLogOnTitle.Width
        .Bottom = lblLogOnTitle.Height
    End With
    'repaint
    Refresh
End Sub

'---------------------------------------------
' Hover Off on action labels
'---------------------------------------------
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblLogOk.BackColor = vbWindowBackground: lblLogOk.ForeColor = vbWindowText: lblLogOk.FontBold = False
    lblLogExit.BackColor = vbWindowBackground: lblLogExit.ForeColor = vbWindowText: lblLogExit.FontBold = False
End Sub

'---------------------------------------------
' End of Program
'---------------------------------------------
Private Sub Form_Terminate()
    On Error Resume Next
    'Call lzShowTaskbar
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    gcnnBilling.Close
End Sub

'---------------------------------------------
' End program on Exit action
'---------------------------------------------
Private Sub lblLogExit_Click()
    End
End Sub

'---------------------------------------------
' Hove on Exit action label
'---------------------------------------------
Private Sub lblLogExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblLogExit.BackColor = vbWindowFocus: lblLogExit.ForeColor = vbBlue: lblLogExit.FontBold = True
    lblLogOk.BackColor = vbWindowBackground: lblLogOk.ForeColor = vbWindowText: lblLogOk.FontBold = False
End Sub

'---------------------------------------------
' Connect to server
'---------------------------------------------
Private Sub lblLogOk_Click()
'  If Trim(txtLogOn(4)) <> "" Then
     lConnectUser
'  Else
'     MsgBox "Please Enter Supervisor name!", , "Required"
'     txtLogOn(4).SetFocus
'  End If
End Sub

'---------------------------------------------
' Hover on Ok action
'---------------------------------------------
Private Sub lblLogOk_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblLogOk.BackColor = vbWindowFocus: lblLogOk.ForeColor = vbBlue: lblLogOk.FontBold = True
    lblLogExit.BackColor = vbWindowBackground: lblLogExit.ForeColor = vbWindowText: lblLogExit.FontBold = False
End Sub

Private Sub lblLogOn_Click(Index As Integer)

End Sub

Private Sub lblLogOnTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        'Visual Basic calls SetCapture when the left mouse
        'button is pressed so call ReleaseCapture so mouse
        'messages will be sent to Windows
        ReleaseCapture
        'Tell Windows the mouse was pressed in the caption
        'area (initiates dragging)
        SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
    End If
End Sub

'---------------------------------------------
' Hover on focused textbox
'---------------------------------------------
Private Sub txtLogOn_GotFocus(Index As Integer)
    txtLogOn(Index).BackColor = vbWindowFocus
End Sub

'---------------------------------------------
' Act on Enter / Esc keypress
'---------------------------------------------
Private Sub txtLogOn_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            lConnectUser
        Case vbKeyEscape
            End
    End Select
End Sub

'---------------------------------------------
' Hover off unfocused textbox
'---------------------------------------------
Private Sub txtLogOn_LostFocus(Index As Integer)
    txtLogOn(Index).BackColor = vbWindowBackground
End Sub

'---------------------------------------------
' Connect to SQL Server
'---------------------------------------------
Private Sub lConnectUser()
Dim wait As New CWaitCursor

    'wait.SetCursor
If Trim(txtLogOn(4).Text) <> "" Then
    lblMainStatus.Caption = "Connecting..."
    lblMainStatus.Refresh
    
    gbConnected = gzConnected(txtLogOn(lcServer).Text, _
                         txtLogOn(lcDatabase).Text, _
                         chkNTLog.Value, _
                         txtLogOn(lcUserID).Text, _
                         txtLogOn(lcPassword).Text)
    If gbConnected Then
        gbSupervisor = txtLogOn(4).Text
        lblMainStatus.Caption = lcCopyright
        Me.Hide
        frmMain.Show 1
        On Error Resume Next
        gcnnBilling.Close
        If gShutDown Then
            Unload Me
        Else
            Me.Show
        End If
    Else
        lblMainStatus.Caption = lcCopyright
    End If
 Else
   MsgBox "Please Enter Supervisor name!", , "Required"
   txtLogOn(4).SetFocus
 End If
End Sub

Private Sub lzHideTaskbar()
    hWnd1 = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(hWnd1, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Sub

Private Sub lzShowTaskbar()
    Call SetWindowPos(hWnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End Sub

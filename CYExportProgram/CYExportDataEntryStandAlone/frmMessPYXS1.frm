VERSION 5.00
Begin VB.Form frmMessPYXS1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Message S1"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "IBM3270 - 1254"
      Size            =   16.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lblMessPYXS13 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   7575
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      Caption         =   "W A R N I N G"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label lblMessPYXS12 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   7575
   End
   Begin VB.Label lblMessPYXS11 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   7575
   End
End
Attribute VB_Name = "frmMessPYXS1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdNo_Click()
    strResponse = False
'    Timer1.Enabled = False
    Unload Me
End Sub
Private Sub cmdYes_Click()
    strResponse = True
'    Timer1.Enabled = False
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyY
            Call cmdYes_Click
        Case vbKeyN
            Call cmdNo_Click
        Case Else
            Beep
            KeyCode = 0
    End Select
End Sub
Private Sub Timer1_Timer()
    lblWarning.Visible = Not lblWarning.Visible
End Sub

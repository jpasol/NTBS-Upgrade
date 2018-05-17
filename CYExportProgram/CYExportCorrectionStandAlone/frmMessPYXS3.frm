VERSION 5.00
Begin VB.Form frmMessPYXS3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Message"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "IBM3270 - 1254"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   700
      Left            =   7200
      Top             =   2520
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   450
      Left            =   4320
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   450
      Left            =   2520
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblWarning1 
      Alignment       =   2  'Center
      Caption         =   "W A R N I N G"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label lblMessPYXS31 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   7575
   End
   Begin VB.Label lblMessPYXS3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   7575
   End
End
Attribute VB_Name = "frmMessPYXS3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNo_Click()
    Timer2.Enabled = False
    Unload Me
End Sub

Private Sub cmdYes_Click()
    strResponse = True
    Timer2.Enabled = False
    Unload Me
End Sub
Private Sub Timer2_Timer()
    lblWarning1.Visible = Not lblWarning1.Visible
End Sub

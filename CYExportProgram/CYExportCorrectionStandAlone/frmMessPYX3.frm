VERSION 5.00
Begin VB.Form frmMessPYX3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Message"
   ClientHeight    =   3705
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   450
      Left            =   480
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   7680
      Top             =   3240
   End
   Begin VB.Label lblMessPYX32 
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
      TabIndex        =   2
      Top             =   2040
      Width           =   7575
   End
   Begin VB.Label lblMessPYX31 
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
      TabIndex        =   1
      Top             =   1200
      Width           =   7575
   End
   Begin VB.Label lblWarning 
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
      TabIndex        =   0
      Top             =   360
      Width           =   7575
   End
End
Attribute VB_Name = "frmMessPYX3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
    lblWarning.Visible = Not lblWarning.Visible
End Sub
Private Sub cmdYes_Click()
    Timer1.Enabled = False
    Unload Me
End Sub

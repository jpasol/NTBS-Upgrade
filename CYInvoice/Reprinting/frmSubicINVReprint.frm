VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmSubicINVReprint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CY Invoice Reprint"
   ClientHeight    =   3135
   ClientLeft      =   4950
   ClientTop       =   2895
   ClientWidth     =   6180
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   3135
   ScaleWidth      =   6180
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   735
      Left            =   3120
      TabIndex        =   3
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txtInvNum 
      Height          =   420
      Left            =   3480
      MaxLength       =   8
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "&Display"
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txtNumDay 
      Height          =   420
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin Crystal.CrystalReport CYInvoice 
      Left            =   5520
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Invoice Preview"
      WindowBorderStyle=   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      ProgressDialog  =   0   'False
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   0   'False
      WindowShowCancelBtn=   0   'False
      WindowShowExportBtn=   0   'False
      WindowShowProgressCtls=   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   "Invoice Number"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "No. of Days (SA only)"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   3135
   End
End
Attribute VB_Name = "frmSubicINVReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gcnnBilling As ADODB.Connection
Dim gbConnected As Boolean

'MDC (20131205)
Dim sqlConBilling As String
Dim sqlConNavis As String

Private Sub cmdDisplay_Click()
Dim rsINVict As New ADODB.Recordset
Dim tmpInvNo As Long

'    If Len(Trim(txtRefNum)) = 0 And Len(Trim(txtInvNum)) = 0 Then
'        MsgBox "Please specify valid entries.", vbExclamation, "Error Message"
'        Exit Sub
'    End If
    
    Screen.MousePointer = vbHourglass
    
'        If Len(Trim(txtNumDay)) > 0 Then    ' SA
'           CYInvoice.ReportFileName = App.Path & "\SubicINVSA1.rpt"
'            CYInvoice.ParameterFields(1) = "InvoiceNo; " & Trim(txtInvNum) & ";TRUE"
'            CYInvoice.ParameterFields(2) = "NumDays; " & Trim(txtNumDay) & ";TRUE"
'            CYInvoice.PrintReport
'        Else
'             CYInvoice.ReportFileName = App.Path & "\SubicInvoice1.rpt"
'             CYInvoice.ParameterFields(1) = "InvoiceNo; " & Trim(txtInvNum) & ";TRUE"
'            CYInvoice.ReportFileName = "c:\ntbs\cyinvoice\reprinting\SubicInvoice1.rpt"
'            CYInvoice.ParameterFields(1) = "InvoiceNo; " & Trim(txtInvNum) & ";TRUE"
'            CYInvoice.PrintReport
'        End If
'
'    Else                             'Use of Invoice Number (for MR/Reg. bills only)
    
        If Not gbConnected Then ConnectToBilling
        With rsINVict
            .Open "SELECT * FROM INVICT WHERE (invnum = " & txtInvNum & ")", _
                   gcnnBilling, , , adCmdText
            If Not .EOF Then
                'tmpInvNo = .Fields("refnum")
                tmpInvNo = .Fields("invnum")
            End If
            .Close
        End With
        'CYInvoice.ReportFileName = App.Path & "\SubicInvoice.rpt"
        'CYInvoice.ParameterFields(1) = "InvoiceNo; " & Trim(tmpInvNo) & ";TRUE"
        CYInvoice.ReportFileName = App.Path & "\SubicInvoice1.rpt"
        CYInvoice.ParameterFields(1) = "InvoiceNo; " & Trim(tmpInvNo) & ";TRUE"
        CYInvoice.PrintReport
        
       
    'End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub ConnectToBilling()
Dim gConnStr As String
Call ReadConfig
    gConnStr = sqlConBilling '"Provider=SQLOLEDB;Data Source=sbitcbilling;Initial Catalog=Billing;Integrated Security=SSPI"
    Set gcnnBilling = New ADODB.Connection
    gcnnBilling.Open gConnStr
    gbConnected = True
End Sub

Public Sub ReadConfig()
Dim Xcnt As Integer
Open App.Path & "\" & "Conn.cfg" For Binary Access Read As #1

Do While Not EOF(1)
    Xcnt = Xcnt + 1
    Select Case Xcnt
        Case 1
            Line Input #1, sqlConBilling
        Case 2
            Line Input #1, sqlConNavis
    End Select
Loop
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gbConnected Then gcnnBilling.Close
End Sub

'Private Sub txtRefNum_GotFocus()
'    SendKeys "{HOME}": SendKeys "+{END}"
'End Sub
Private Sub txtInvNum_Change()

End Sub

Private Sub txtInvNum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     cmdDisplay.SetFocus
  End If
End Sub

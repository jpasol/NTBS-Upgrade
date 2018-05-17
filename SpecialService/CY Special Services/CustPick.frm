VERSION 5.00
Begin VB.Form frmCustPick 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer PickList"
   ClientHeight    =   8265
   ClientLeft      =   1305
   ClientTop       =   3570
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   10935
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1875
      TabIndex        =   1
      Top             =   75
      Width           =   8940
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   1815
   End
   Begin VB.CommandButton cmdName 
      Caption         =   "&Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1875
      TabIndex        =   4
      Top             =   525
      Width           =   8940
   End
   Begin VB.CommandButton cmdCode 
      Caption         =   "&Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   75
      TabIndex        =   3
      Top             =   525
      Width           =   1815
   End
   Begin VB.ListBox lstCustomer 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   75
      TabIndex        =   2
      Top             =   960
      Width           =   10740
   End
End
Attribute VB_Name = "frmCustPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--Key Code Constants
Const KEY_BACK = &H8
Const KEY_DELETE = &H2E
Const KEY_CLEAR = &HC

'---Focus constants
Const TEXTBOX_FOCUS = 1       ' currently in text box
Const LISTBOX_FOCUS = 2       ' currently in list box

'---module level variables
Dim miCtrlFocus As Integer      ' which control (textbox/listbox) has focus
Dim miNumKeys As Integer        ' number of keys pressed by user
Dim mfScrolling As Integer      ' True if textbox triggers listbox scroll.
Dim mfKeepKey As Integer        ' False if user hit delete/backspace

Dim blnByName As Boolean

Private Sub cmdCode_Click()
    blnByName = False
    lzOpenCustomer
    txtCode.SelStart = 0: txtCode.SelLength = Len(txtName): txtCode.SetFocus
End Sub

Private Sub cmdName_Click()
    blnByName = True
    lzOpenCustomer
    txtName.SelStart = 0: txtName.SelLength = Len(txtName)
End Sub

Private Sub Form_Load()
    gsCusCode = "": gsCusName = ""
    blnByName = True
    lzOpenCustomer
    txtName.SelStart = 0: txtName.SelLength = Len(txtName)
End Sub

Private Sub lzOpenCustomer()
Dim wait As New CWaitCursor
Dim rstCustomer As ADODB.Recordset
Dim SQLStmt As String
    
    ' clear list first
    lstCustomer.Clear
    txtCode.Enabled = Not blnByName: txtName.Enabled = blnByName
    
    wait.SetCursor

    On Error GoTo err_OpenCustomer
    Set rstCustomer = New ADODB.Recordset
    If blnByName Then
        SQLStmt = "SELECT * FROM Customer ORDER BY cusnam"
    Else
        SQLStmt = "SELECT * FROM Customer ORDER BY cuscde"
    End If
    rstCustomer.Open SQLStmt, gcnnBilling, adOpenStatic, adLockReadOnly, adCmdText
    
    With rstCustomer
        While Not .EOF
            lstCustomer.AddItem Left((!cuscde + "      "), 6) & "  " & !cusnam
            .MoveNext
        Wend
    End With
    lstCustomer.Selected(0) = True
    txtCode.Text = Left(lstCustomer.Text, 6)
    txtName.Text = Trim(Mid(lstCustomer.Text, 9))
        
    Exit Sub

err_OpenCustomer:
    MsgBox "Error accessing Customer table ...", vbCritical
End Sub

Private Sub lstCustomer_Click()
Dim szListText As String
Dim iListIndex As Integer
    
    On Error Resume Next
    With lstCustomer
        If .ListIndex >= 0 And miCtrlFocus = LISTBOX_FOCUS Then
            ' user has clicked on the liat box
            iListIndex = .ListIndex
            szListText = .List(iListIndex)
            txtCode.Text = Left(szListText, 6)
            txtName.Text = Trim(Mid(szListText, 9))
        End If
    End With
End Sub

Private Sub lstCustomer_DblClick()
    lzRetPick
End Sub

Private Sub lzRetPick()
    With lstCustomer
        If .ListIndex >= 0 Then
            gsCusCode = Left(.Text, 6)
            gsCusName = Trim(Mid(.Text, 9))
            frmCYSCCR.txtCusName.Text = gsCusName
            frmCYSCCR.Text1.Text = gsCusCode
            Unload Me
        End If
    End With
End Sub

Private Sub lzRetNull()
    gsCusCode = "": gsCusName = ""
    Unload Me
End Sub

Private Sub lstCustomer_GotFocus()
    lstCustomer.BackColor = vbInfoBackground
End Sub

Private Sub lstCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    miCtrlFocus = LISTBOX_FOCUS
    miNumKeys = 0
End Sub

Private Sub lstCustomer_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            lzRetPick
        Case vbKeyEscape
            lzRetNull
        Case Else
    End Select
End Sub

Private Sub lzShowPick()
    txtCode.Text = Left(lstCustomer.Text, 6)
    txtName.Text = Trim(Mid(lstCustomer.Text, 9))
End Sub

Private Sub lstCustomer_LostFocus()
    lstCustomer.BackColor = vbWindowBackground
End Sub



Private Sub lstCustomer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    miCtrlFocus = LISTBOX_FOCUS
    miNumKeys = 0

End Sub

Private Sub txtCode_Change()
Dim szSrchText As String    ' contents of text box
Dim iTxtLen As Integer      ' length of search string
Dim iListIndex As Integer   ' set by SearchListBox
Dim fReturn As Integer      ' ret. from SearchListBox
Dim szListText As String    ' contents of list box
  
    '---Start of Code
    On Error Resume Next
    
    If miCtrlFocus = TEXTBOX_FOCUS And mfKeepKey And Not mfScrolling Then
        iTxtLen = Len(txtCode.Text)
        If iTxtLen Then
            miNumKeys = IIf(miNumKeys < iTxtLen, miNumKeys, iTxtLen)
            szSrchText = txtCode.Text
            fReturn = SearchByCode(szSrchText, lstCustomer, iListIndex)
            
            mfScrolling = True
            If iListIndex = -1 Then
                lstCustomer.ListIndex = -1
            Else
                ' perfect match was found
                lstCustomer.Selected(iListIndex) = True
                szListText = lstCustomer.List(lstCustomer.ListIndex)
                txtCode.Text = Left(szListText, 6)
                txtName.Text = Trim(Mid(szListText, 9))
                txtCode.SelStart = miNumKeys
                txtCode.SelLength = (Len(txtCode.Text) - miNumKeys)
            End If
            mfScrolling = False
        End If
    End If
End Sub

Private Sub txtCode_GotFocus()
    txtCode.BackColor = vbInfoBackground
    miNumKeys = 0
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtName.Text)
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_BACK Or KeyCode = KEY_DELETE Or KeyCode = KEY_CLEAR Then
        mfKeepKey = False
        If KeyCode = KEY_BACK Then
            ' unhilight current item; next search
            ' will start at top of list
            lstCustomer.ListIndex = -1
        End If
    Else
        mfKeepKey = True
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            lzRetPick
        Case vbKeyEscape
            lzRetNull
        Case Else
    End Select
    miCtrlFocus = TEXTBOX_FOCUS
    If mfKeepKey Then
        miNumKeys = Len(txtCode.Text) + 1
    End If
End Sub

Private Sub txtCode_LostFocus()
    txtCode.BackColor = vbWindowBackground
End Sub

Private Sub txtName_Change()
Dim szSrchText As String    ' contents of text box
Dim iTxtLen As Integer      ' length of search string
Dim iListIndex As Integer   ' set by SearchListBox
Dim fReturn As Integer      ' ret. from SearchListBox
Dim szListText As String    ' contents of list box
  
    '---Start of Code
    On Error Resume Next
    
    If miCtrlFocus = TEXTBOX_FOCUS And mfKeepKey And Not mfScrolling Then
        iTxtLen = Len(txtName.Text)
        If iTxtLen Then
            miNumKeys = IIf(miNumKeys < iTxtLen, miNumKeys, iTxtLen)
            szSrchText = txtName.Text
            fReturn = SearchByName(szSrchText, lstCustomer, iListIndex)
            
            mfScrolling = True
            If iListIndex = -1 Then
                lstCustomer.ListIndex = -1
            Else
                ' perfect match was found
                lstCustomer.Selected(iListIndex) = True
                szListText = lstCustomer.List(lstCustomer.ListIndex)
                txtCode.Text = Left(szListText, 6)
                txtName.Text = Trim(Mid(szListText, 9))
                txtName.SelStart = miNumKeys
                txtName.SelLength = (Len(txtName.Text) - miNumKeys)
            End If
            mfScrolling = False
        End If
    End If
End Sub

Private Sub txtName_GotFocus()
    txtName.BackColor = vbInfoBackground
    miNumKeys = 0
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_BACK Or KeyCode = KEY_DELETE Or KeyCode = KEY_CLEAR Then
        mfKeepKey = False
        If KeyCode = KEY_BACK Then
            ' unhilight current item; next search
            ' will start at top of list
            lstCustomer.ListIndex = -1
        End If
    Else
        mfKeepKey = True
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            lzRetPick
            Exit Sub
        Case vbKeyEscape
            lzRetNull
        Case Else
    End Select
    miCtrlFocus = TEXTBOX_FOCUS
    If mfKeepKey Then
        miNumKeys = Len(txtName.Text) + 1
    End If
End Sub

Private Function SearchByName(ByVal szSearchText As String, lbScroll As ListBox, iListIndex As Integer) As Integer
'=======================================================
' Simple function to create a scrolling list box.
' The procedure will select the first item in the list
' box in which Left(List box text,size of search string)
' matches the search string.
'==========================================================
'---Constants (returned from StrComp)
Const FOUND = 0
Const LT = -1
Const GT = 1

'---Variable declarations
Dim iListStart As Integer     ' starting point in list
Dim iListCount As Integer     ' no. of items in list box
Dim iTxtLen As Integer
Dim szListText As String      ' current list item
Dim vCompResult               ' result of string comp function
Dim fFound As Integer         ' match found?
Dim fDone As Integer          ' Terminates search if true

    '---Start of Code
    fFound = False
    iTxtLen = Len(szSearchText)
    
    If iTxtLen > 0 And lbScroll.ListCount > 0 Then
        iListStart = lbScroll.ListIndex
        If iListStart = -1 Then iListStart = 0
        iListIndex = iListStart
        iListCount = lbScroll.ListCount
        szListText = Mid(lbScroll.List(iListStart), 9, iTxtLen)
        
        ' check to see if current item matches
        fFound = CInt(StrComp(szSearchText, szListText, 1))
        
        If fFound <> FOUND Then
            fDone = False
        
            If (fFound = LT) Then
                iListIndex = 0
                fFound = False
            Else
                iListIndex = iListIndex + 1
            End If
        
            Do While (iListIndex <= iListCount) And Not fDone
                szListText = Mid(lbScroll.List(iListIndex), 9, iTxtLen)
                vCompResult = StrComp(szSearchText, szListText, 1)
                If IsNull(vCompResult) Then
                  iListIndex = -1
                  Exit Do
                End If
                fFound = (CInt(vCompResult) = FOUND)
                fDone = fFound Or (CInt(vCompResult) = -1)
                If Not fDone Then
                  iListIndex = iListIndex + 1
                End If
            Loop
        
            If Not fFound Then
              iListIndex = -1
            End If
        End If
    End If
    
    SearchByName = fFound
End Function ' ScrollListBox

Private Sub txtName_LostFocus()
    txtName.BackColor = vbWindowBackground
End Sub

Private Function SearchByCode(ByVal szSearchText As String, lbScroll As ListBox, iListIndex As Integer) As Integer
'=======================================================
' Simple function to create a scrolling list box.
' The procedure will select the first item in the list
' box in which Left(List box text,size of search string)
' matches the search string.
'==========================================================
'---Constants (returned from StrComp)
Const FOUND = 0
Const LT = -1
Const GT = 1

'---Variable declarations
Dim iListStart As Integer     ' starting point in list
Dim iListCount As Integer     ' no. of items in list box
Dim iTxtLen As Integer
Dim szListText As String      ' current list item
Dim vCompResult               ' result of string comp function
Dim fFound As Integer         ' match found?
Dim fDone As Integer          ' Terminates search if true

    '---Start of Code
    fFound = False
    iTxtLen = Len(szSearchText)
    
    If iTxtLen > 0 And lbScroll.ListCount > 0 Then
        iListStart = lbScroll.ListIndex
        If iListStart = -1 Then iListStart = 0
        iListIndex = iListStart
        iListCount = lbScroll.ListCount
        szListText = Left(lbScroll.List(iListStart), iTxtLen)
        
        ' check to see if current item matches
        fFound = CInt(StrComp(szSearchText, szListText, 1))
        
        If fFound <> FOUND Then
            fDone = False
        
            If (fFound = LT) Then
                iListIndex = 0
                fFound = False
            Else
                iListIndex = iListIndex + 1
            End If
        
            Do While (iListIndex <= iListCount) And Not fDone
                szListText = Left(lbScroll.List(iListIndex), iTxtLen)
                vCompResult = StrComp(szSearchText, szListText, 1)
                If IsNull(vCompResult) Then
                  iListIndex = -1
                  Exit Do
                End If
                fFound = (CInt(vCompResult) = FOUND)
                fDone = fFound Or (CInt(vCompResult) = -1)
                If Not fDone Then
                  iListIndex = iListIndex + 1
                End If
            Loop
        
            If Not fFound Then
              iListIndex = -1
            End If
        End If
    End If
    
    SearchByCode = fFound
End Function ' ScrollListBox


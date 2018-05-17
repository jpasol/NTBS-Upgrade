VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubicINVDE01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarUserid As Variant 'local copy
Public Property Let Userid(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Userid = 5
    mvarUserid = vData
    gUserID = mvarUserid
End Property
Public Property Set Userid(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Userid = Form1
    Set mvarUserid = vData
End Property
Public Property Get Userid() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Userid
    If IsObject(mvarUserid) Then
        Set Userid = mvarUserid
    Else
        Userid = mvarUserid
    End If
End Property

Public Sub Execute()
    frmSubicINVDE01.Show vbModal
End Sub

Public Sub Disconnect()
    On Error Resume Next
    gcnnBilling.Close
    Set gcnnBilling = Nothing
End Sub

Public Function ConnectByStr(ByVal pCnnStr As String) As Boolean
Dim errBilling As ADODB.Error
Dim lsErrStr As String
   
    ' Open the database.
    On Error GoTo err_Connect
    Set gcnnBilling = New ADODB.Connection
    gcnnBilling.Open pCnnStr
    gbConnected = True
    ConnectByStr = True
   
    Exit Function
    
err_Connect:
    ConnectByStr = False: gbConnected = False
    For Each errBilling In gcnnBilling.Errors
        With errBilling
            lsErrStr = "Connection Error. " & .Description & vbLf & _
            "Verify Log On then retry.  Contact MIS for assistance."
        End With
        MsgBox lsErrStr, vbCritical
    Next
End Function

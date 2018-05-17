VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubicINVReports"
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
    frmSubicINVReports.Show vbModal
End Sub

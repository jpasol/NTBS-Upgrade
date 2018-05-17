VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCYXINQ01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarUserid As String 'local copy
Private mvarUserType As String 'local copy
Public Sub Execute()
    frmCYXINQ01.Show vbModal
End Sub
Public Property Let UserType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserType = 5
    mvarUserType = vData
End Property
Public Property Get UserType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserType
    UserType = mvarUserType
End Property
Public Property Let Userid(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Userid = 5
    mvarUserid = UCase(vData)
    gUserid = UCase(vData)
End Property
Public Property Get Userid() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Userid
    Userid = mvarUserid
End Property

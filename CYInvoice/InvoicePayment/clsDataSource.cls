VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' TODO: Declare local ADO Recordset object. For example:
Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
'    Select Case DataMember
'    Case ""             ' Default
'        Set Data = Nothing
'    Case Else           ' Default
Set rs = New ADODB.Recordset
        rs.Open "invict", gConnection, adOpenDynamic, adLockOptimistic, adCmdTable
        Set Data = rs
'    End Select
End Sub



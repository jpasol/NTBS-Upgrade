VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim CustomerFilled As Boolean
Dim mvarreturnedName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarAdrBalance As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarCode As Variant 'local copy
Public Property Let Code(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Code = 5
    mvarCode = vData
End Property
Public Property Get Code() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Code
    Code = mvarCode
End Property
Public Property Let AdrBalance(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AdrBalance = 5
    mvarAdrBalance = vData
End Property
Public Property Get AdrBalance() As Variant
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.AdrBalance

If Len(Trim(mvarCode)) <> 0 Then
    DE.GetADRBal mvarCode, mvarAdrBalance
    DE.getCustomerName mvarCode, mvarreturnedName
'    If mvarAdrBalance <> 0 Then
'        mvarAdrBalance = AdrBalance
''    Else
''        mvarAdrBalance = 0
'    End If
Else
    mvarAdrBalance = 0
    mvarreturnedName = ""
End If
AdrBalance = mvarAdrBalance
'    Dim AdrBal As Recordset
'    If Len(Trim(mvarCode)) <> 0 Then
'        DE.GetADRBal mvarCode
'        Set AdrBal = DE.rsgetAdrBal
'        If AdrBal.RecordCount = 0 Then
'            mvarAdrBalance = 0
''            mvarreturnedName = ""
'        Else
'            mvarAdrBalance = AdrBal.Fields(0)
''            mvarreturnedName = AdrBal.Fields(1)
'        End If
'        AdrBal.Close
'        Set AdrBal = Nothing
'    Else
'        mvarAdrBalance = 0
'    End If
'    AdrBalance = mvarAdrBalance
End Property
Public Function ApplyAdr(AmtToApply As Single) As Boolean
' ** Apllication of ADR Amount
End Function
Public Property Get Name() As String
    Name = mvarreturnedName
End Property
Public Property Let Name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AdrBalance = 5
    mvarreturnedName = vData
End Property
Public Sub FillCustomer()
    Dim x As Integer
    Dim rstCustomer As Recordset
    DE.Customer
    Set rstCustomer = DE.rsCustomer
    x = 0
    frmCostumer.flxCustomer.AllowUserResizing = flexResizeColumns
    frmCostumer.flxCustomer.ColWidth(0) = 1
    frmCostumer.flxCustomer.ColWidth(1) = 7000
    frmCostumer.flxCustomer.ColWidth(2) = 1400
    frmCostumer.fillingCustomer = True
    With rstCustomer
        Do While Not .EOF
            If x > 0 Then
                frmCostumer.flxCustomer.AddItem ""
            End If
            frmCostumer.flxCustomer.Row = x
            If Not Len(Trim(.Fields("cusnam"))) > 40 Then
                frmCostumer.flxCustomer.TextMatrix(x, 1) = Trim(.Fields("cusnam")) & Space(42 - Len(Trim(.Fields("cusnam"))))
            Else
                frmCostumer.flxCustomer.TextMatrix(x, 1) = Trim(.Fields("cusnam"))
            End If
            frmCostumer.flxCustomer.TextMatrix(x, 2) = Trim(.Fields("cuscde"))
            x = x + 1
            .MoveNext
        Loop
    End With
    frmCostumer.ListCount = x
    rstCustomer.Close
    Set rstCustomer = Nothing
    CustomerFilled = True
    frmCostumer.fillingCustomer = False
End Sub
Public Function ShowList() As Boolean
    Dim pAmount As Single
    If Not CustomerFilled Then
        Call FillCustomer
    End If
    If Len(Trim(frmCostumer.txtCode)) = 0 Then
        frmCostumer.txtCode = frmCostumer.flxCustomer.TextMatrix(0, 2)
        frmCostumer.txtName = frmCostumer.flxCustomer.TextMatrix(0, 1)
        frmCostumer.flxCustomer.Row = 0
        frmCostumer.flxCustomer.Col = 0
        frmCostumer.flxCustomer.ColSel = 2
        frmCostumer.flxCustomer.TopRow = 0
    Else
        frmCostumer.flxCustomer.Col = 0
        frmCostumer.flxCustomer.ColSel = 2
        frmCostumer.flxCustomer.TopRow = frmCostumer.flxCustomer.Row
    End If
    frmCostumer.Show vbModal
    If Len(Trim(frmCostumer.CustomerCode)) <> 0 Then
        mvarCode = Format(frmCostumer.CustomerCode, "000000")
        mvarreturnedName = Trim(frmCostumer.CustomerName)
    End If
    If frmCostumer.TransCancelled = True Then
        ShowList = False
    Else
        ShowList = True
    End If
End Function
Private Sub Initialize()
    frmCostumer.CustomerCode = 0
    frmCostumer.CustomerName = ""
    CustomerFilled = False
    Load frmCostumer
End Sub

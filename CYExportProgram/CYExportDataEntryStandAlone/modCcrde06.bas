Attribute VB_Name = "modCcrde06"
Option Explicit

Public gcnnBilling As ADODB.Connection
Public gcnnNavis As ADODB.Connection

' Public Variables
Public gPassword As String * 10
Public gConnStr As String
Public gbConnected As Boolean
Public gbNavis As Boolean
Public gsCaption As String

' INI variables

Public gINIServer As String * 20
Public gINIDatabase As String * 20

' API Declares

Public Declare Function WNetGetUser Lib "mpr" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'***********************************************************
Public DomesticMode As Boolean
Public strResponse As Boolean
Public gUserID As String * 10
Public gSuprvsr As String * 15
Public PrinterRef As Integer
Public blnExit As Boolean
Public strWrkstn As String
Public strPrinter As String
Public NumberOfCCR As Integer
Public lngTagNewCCR As Long  '   ** Temporary CCR
Public lngCCR As Long   'CCR Number
Public Refnum As Long
Public Seqnum As Long
Public Customer As String
Public strCommodity As String
Public AdrAmount As Single
Public DE As New deCCRDE06
Public strAdrAmt As String * 12
Public strCashAmt As String * 12

'   ** Temporary Variables

Public AdrAmt As Currency
Public DetailTl As Currency
Public DetailAmt As Currency
Public TotalAmt As Currency
Public ChkTotal As Currency
Public CashAmt As Currency
Public CashAmount As Currency

Public ChkAmt1 As Currency
Public ChkAmt2 As Currency
Public ChkAmt3 As Currency
Public ChkAmt4 As Currency
Public ChkAmt5 As Currency

Public ChkAmount As Currency
Public ChkAmount1 As Currency
Public ChkAmount2 As Currency
Public ChkAmount3 As Currency
Public ChkAmount4 As Currency
Public ChkAmount5 As Currency

Public sngTempAmt As Currency
Public blnChkno1 As Boolean
Public blnChkno2 As Boolean
Public blnChkno3 As Boolean
Public blnChkno4 As Boolean
Public blnChkno5 As Boolean

Public lngRcount As Long
'   **  Parameter Passed
Public strChqAmt As String * 12
Public strChqAmt1 As String * 12
Public strChqAmt2 As String * 12
Public strChqAmt3 As String * 12
Public strChqAmt4 As String * 12
Public strChqAmt5 As String * 12
Public strCshAmt As String * 12

Public Type Rates
    Rtecode As String
    CntSze As Integer
    RteAmt As Currency
End Type

' ** Used for Printing
Public Type CList
    Refnum As Long
    Seqnum As Long
    CCRnum As Long
    Cusnam As String
    UGCode As String * 1
End Type
' ** End of Printing Public Variables

Public RateArr(1 To 300) As Rates

Public Function ComputeWeighing() As Currency
    'If chkWeighing.Value = 1 Then
        ComputeWeighing = GetRate2("MCTRUS", "20")
    'End If
End Function

Public Function Comp_Arrastre(ByRef sngdangramt As Currency, ByRef sngBscArr As Currency, ByRef sngRton As Currency, _
            ByRef sngOvzAmt As Currency, cnt_Size As Integer, cnt_Dngr As String, _
            cnt_Length As Single, cnt_Width As Single, _
            cnt_Height As Single, cnt_UMS As String) As Currency
    Dim arrAmt As Currency
    Dim sngRevton As Currency
    Const CntRton20 As Single = 27.95
    Const CntRton40 As Single = 63.75
    Const CntRton45 As Single = 76.38
' ** COMPUTATION FOR BASIC ARRASTRE
    If DomesticMode Then
        Select Case cnt_Size
        Case 20
            arrAmt = GetRate("CBDOM1", cnt_Size)
            sngBscArr = GetRate("CBDOM1", cnt_Size)
        Case 40
            arrAmt = GetRate("CBDOM2", cnt_Size)
            sngBscArr = GetRate("CBDOM2", cnt_Size)
        Case 45
            arrAmt = GetRate("CBDOM3", cnt_Size)
            sngBscArr = GetRate("CBDOM3", cnt_Size)
        End Select
    Else
        Select Case cnt_Size
        Case 20
            arrAmt = GetRate("CBEXP1", cnt_Size)
            sngBscArr = GetRate("CBEXP1", cnt_Size)
        Case 40
            arrAmt = GetRate("CBEXP2", cnt_Size)
            sngBscArr = GetRate("CBEXP2", cnt_Size)
        Case 45
            arrAmt = GetRate("CBEXP3", cnt_Size)
            sngBscArr = GetRate("CBEXP3", cnt_Size)
        End Select
    End If
    sngRton = 0
    sngRevton = 0
    sngOvzAmt = 0

' ** COMPUTATION FOR OVERSIZE AMOUNTS

    
    If (cnt_Length <> 0) And (cnt_Width <> 0) And (cnt_Height <> 0) Then
        If cnt_UMS <> "I" Then
            cnt_Length = Round((cnt_Length / 2.54), 2)
            cnt_Width = Round((cnt_Width / 2.54), 2)
            cnt_Height = Round((cnt_Height / 2.54), 2)
        Else
            cnt_Length = Round(cnt_Length, 2)
            cnt_Width = Round(cnt_Width, 2)
            cnt_Height = Round(cnt_Height, 2)
        End If
        sngRevton = ((cnt_Length * cnt_Width * cnt_Height) / 1728) / 40
        Select Case cnt_Size
            Case 20
                If (sngRevton > CntRton20) Or (sngRevton = CntRton20) Then
                    sngRton = sngRevton - CntRton20
                Else
                    sngRton = sngRevton
                End If
            Case 40
                If (sngRevton > CntRton40) Or (sngRevton = CntRton40) Then
                    sngRton = sngRevton - CntRton40
                Else
                    sngRton = sngRevton
                End If
            Case 45
                If (sngRevton > CntRton45) Or (sngRevton = CntRton45) Then
                    sngRton = sngRevton - CntRton45
                Else
                    sngRton = sngRevton
                End If
        End Select
        sngRton = Round(sngRton, 2)
        'sngOvzAmt = sngRton * GetRate("RTAREX", 0)
        sngOvzAmt = sngRton * GetRate("CBEXPA", 0)
        sngOvzAmt = Round(sngOvzAmt, 2)
        arrAmt = arrAmt + sngOvzAmt
    End If
' ** COMPUTATION FOR DANGER CLASS
    If cnt_Dngr <> "" Then
        Select Case cnt_Dngr
            Case "1", "6", "8"
                sngdangramt = arrAmt * 0.5
            Case "2", "3", "4", "7"
                sngdangramt = arrAmt * 0.25
            Case "5", "9"
                sngdangramt = arrAmt * 0.1
        End Select
        sngdangramt = Round(sngdangramt, 2)
        arrAmt = arrAmt + sngdangramt
    End If
    Comp_Arrastre = arrAmt
End Function

Public Function GetRate(Rtecode As String, cntsize As Integer) As Currency
Dim ctrArr As Integer
ctrArr = 0
    For ctrArr = 1 To 300
     If Rtecode = Trim(RateArr(ctrArr).Rtecode) And cntsize = RateArr(ctrArr).CntSze Then
        GetRate = RateArr(ctrArr).RteAmt
        Exit For
    End If
    Next ctrArr
End Function

Public Function GetRate2(Rtecode As String, cntsize As String) As Currency
Dim ctrArr As Integer
ctrArr = 0
    For ctrArr = 1 To 300
     If Rtecode = Trim(RateArr(ctrArr).Rtecode) Then
        GetRate2 = RateArr(ctrArr).RteAmt
        Exit For
    End If
    Next ctrArr
End Function

Public Function UpdateIsN4BillingPermissionGrantedStatus(ByVal unitId As String) As String
    Dim rstN4Status As ADODB.Recordset
    Dim strUpdateStatus As String
    
    Set rstN4Status = New ADODB.Recordset
    
    strUpdateStatus = ""
    strUpdateStatus = "UPDATE CCRCYX " & _
                    "SET IsN4BillingPermissionGranted = 1 " & _
                    "WHERE cntnum = '" & unitId & "' and IsN4BillingPermissionGranted = 0 "
                    
    rstN4Status.Open strUpdateStatus, gcnnBilling, adOpenForwardOnly, adLockReadOnly
    
   ' GetGKey = rstGKey.Fields(0)
End Function


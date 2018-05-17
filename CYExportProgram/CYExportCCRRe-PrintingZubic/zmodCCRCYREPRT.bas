Attribute VB_Name = "zmodCCRCYREPRT"
Option Explicit

Public DE As New zdeCCRCYREPRT
Public gUserid As String * 10
Public DomesticMode As Boolean
Public RefNum As Long
Public SeqNum As Long
Public Customer As String
Public strCommodity As String
Public AdrAmount As Single

Public strAdrAmt As String * 12
Public strCashAmt As String * 12

'   ** Temporary Variables

Public AdrAmt As Single
Public DetailTl As Single
Public DetailAmt As Single
Public TotalAmt As Single
Public ChkTotal As Single
Public CashAmt As Single
Public CashAmount As Single

Public ChkAmt1 As Single
Public ChkAmt2 As Single
Public ChkAmt3 As Single
Public ChkAmt4 As Single
Public ChkAmt5 As Single

Public ChkAmount As Single

Public ChkAmount1 As Single
Public ChkAmount2 As Single
Public ChkAmount3 As Single
Public ChkAmount4 As Single
Public ChkAmount5 As Single

Public sngTempAmt As Single

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

Public PrinterRef As Integer

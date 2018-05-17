Attribute VB_Name = "modPaySettings"
Option Explicit
Public Const FileName As String = "\PaySetting.ini"
Public Const LostFocusBackColor As Long = &H80000016
Public Const GotFocusBackColor As Long = &H80000005
Public PaymentINI As FileSetting
Public gSysDirectory As String

Private Type FileSetting
  RecNo As Integer
  TransCode  As String * 3
  WithADR As Boolean
  WithPOS As Boolean
  With_ePay As Boolean
  WithBankFund As Boolean
  SpecialGpCYImp As Boolean
  POSFee As Single
  CFSCode As String * 1
  DateSet As Date
End Type

' API Declares
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function WNetGetUser Lib "mpr" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Function GetWindowsSystemDirectory() As String
    Dim buffer As String * 512, length As Integer
    length = GetSystemDirectory(buffer, Len(buffer))
    GetWindowsSystemDirectory = Left$(buffer, length)
End Function

Public Sub Retrieve_Values_frmFile(ByVal pTransType As String)
    Dim n As Integer
    Dim RecCTR As Integer
     
     n = FreeFile
     RecCTR = 1
     
      Open App.Path & FileName For Random As #n Len = Len(PaymentINI)
      Get #n, RecCTR, PaymentINI
      
    If pTransType <> "" Then
        Do While Not EOF(n)
             If Trim(pTransType) = Trim(PaymentINI.TransCode) Then
                    Exit Do
              Else
                    RecCTR = RecCTR + 1
                    Get #n, RecCTR, PaymentINI
              End If
        Loop
   End If
   Close #n
End Sub

Public Sub Retrieve_Values_frmFileIMP(ByVal pTransType As String)
    Dim n As Integer
    Dim RecCTR As Integer
    Dim gWindowsSystemDIR As Object
    n = FreeFile
     RecCTR = 1

        gWindowsSystemDIR = GetWindowsSystemDirectory
      Open gWindowsSystemDIR & FileName For Random As #n Len = Len(PaymentINI)
      Get #n, RecCTR, PaymentINI

    If Trim(pTransType) <> "" Then
        Do While Not EOF(n)
             If Trim(pTransType) = Trim(PaymentINI.TransCode) Then
                    Close #n
                    Exit Do
                    
              Else
                    RecCTR = RecCTR + 1
                    Get #n, RecCTR, PaymentINI
              End If
        Loop
   End If
   Close #n
End Sub

Public Function zCurrentUser() As String
 Dim lpUserName As String * 64
    If WNetGetUser("", lpUserName, Len(lpUserName)) Then
        zCurrentUser = ""
    Else
        zCurrentUser = "AGALLARDO" 'UCase(Trim(Left(lpUserName, InStr(lpUserName, Chr(0)) - 1)))
    End If
End Function

Public Function zCurrentComputer() As String
Dim lSize As Long
Dim sBuffer As String
    sBuffer = Space$(15& + 1)
    lSize = Len(sBuffer)
    zCurrentComputer = ""
    If GetComputerName(sBuffer, lSize) Then
        zCurrentComputer = UCase(Trim(Left$(sBuffer, lSize)))
    End If
End Function



Attribute VB_Name = "ModCCRDE08"
Option Explicit
Public retKey As Boolean
Public strResponse As Boolean
Public gUserid As String
Public DE As New deCCRde08

Public gcnnBilling As ADODB.Connection

' Public Variables

Public gPassword As String * 10
Public gConnStr As String
Public gbConnected As Boolean
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


Public strContainer As String * 12

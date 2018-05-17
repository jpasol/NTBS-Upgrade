VERSION 5.00
Begin VB.UserControl utxtEntry 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   BeginProperty Font 
      Name            =   "IBM3270 - 1254"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   390
   ScaleWidth      =   2865
   ToolboxBitmap   =   "utxtEntry.ctx":0000
   Begin VB.TextBox txtNumeric 
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "utxtEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Miramed P. Puelong Feb. 26, 1999 - Numeric Only Textbox
'Event Declarations:
Event Change() 'MappingInfo=txtNumeric,txtNumeric,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event Click() 'MappingInfo=txtNumeric,txtNumeric,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=txtNumeric,txtNumeric,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtNumeric,txtNumeric,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtNumeric,txtNumeric,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtNumeric,txtNumeric,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtNumeric,txtNumeric,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtNumeric,txtNumeric,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtNumeric,txtNumeric,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
'Default Property Values:
'Const m_def_Value = "0"
'Const m_def_Value = ""
'Const m_def_Value = " "
'Const m_def_Value = 0
Const m_def_First = 0
Const m_def_Last = 0
Const m_def_DecimalPlaces = 0
Const m_def_IncludeDecimal = 1
Const m_def_Maskformat = ""
'Property Variables:
'Dim m_Value As String
'Dim m_Value As String
'Dim m_Value As String
'Dim m_Value As Variant
Dim m_First As Boolean
Dim m_Last As Boolean
Dim m_DecimalPlaces As Integer
Dim m_IncludeDecimal As Boolean
Dim m_Maskformat As String
Private Sub txtNumeric_GotFocus()
    txtNumeric.BackColor = &HFFFFFF
    txtNumeric.SelStart = 0
    txtNumeric.SelLength = Len(txtNumeric.Text)
    txtNumeric.Refresh
End Sub
Private Sub txtNumeric_LostFocus()
Dim Numbertochange As Double
txtNumeric.BackColor = &H8000000F
If Len(Trim(m_Maskformat)) > 0 And Len(Trim(txtNumeric.Text)) > 0 Then
    Numbertochange = CDec(txtNumeric.Text)
    txtNumeric.Text = Format(Numbertochange, m_Maskformat)
End If
End Sub
Private Sub UserControl_Initialize()
txtNumeric.Width = UserControl.Width
txtNumeric.Height = UserControl.Height
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
'    Case 40
'        If m_Last = False Then
'            SendKeys "{Tab}", True
'        End If
    Case 38
        If m_First = False Then
            SendKeys "+{Tab}", True
        End If
End Select
End Sub

Private Sub UserControl_Resize()
txtNumeric.Width = UserControl.Width
txtNumeric.Height = UserControl.Height
End Sub
'MappingInfo=txtNumeric,txtNumeric,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = txtNumeric.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtNumeric.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
'MappingInfo=txtNumeric,txtNumeric,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = txtNumeric.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtNumeric.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
'MappingInfo=txtNumeric,txtNumeric,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = txtNumeric.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtNumeric.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
'MappingInfo=txtNumeric,txtNumeric,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = txtNumeric.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set txtNumeric.Font = New_Font
    PropertyChanged "Font"
End Property
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property
Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property
'MappingInfo=txtNumeric,txtNumeric,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = txtNumeric.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    txtNumeric.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
'MappingInfo=txtNumeric,txtNumeric,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    txtNumeric.Refresh
End Sub
Private Sub txtNumeric_Click()
    RaiseEvent Click
End Sub
Private Sub txtNumeric_DblClick()
    RaiseEvent DblClick
End Sub
Private Sub txtNumeric_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub txtNumeric_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If KeyAscii <> 8 Then
        Select Case KeyAscii
            Case 13
                SendKeys "{Tab}", True
            Case 27
                SendKeys "+{Tab}", True
        Case 46
            If m_IncludeDecimal And m_DecimalPlaces > 0 Then
                If InStr(1, txtNumeric.Text, ".") <> 0 Then
                    KeyAscii = 0
                    Beep
                End If
            Else
                KeyAscii = 0
                Beep
            End If
        Case Else
            If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
                KeyAscii = 0
                Beep
            Else
                If Len(txtNumeric.Text) = txtNumeric.SelLength Then
                    txtNumeric.Text = ""
                    txtNumeric.SetFocus
'                    txtNumeric.SelStart = 1
                    Exit Sub
                End If
                If m_IncludeDecimal And m_DecimalPlaces > 0 And InStr(1, txtNumeric.Text, ".") > 0 Then
                    If InStr(1, txtNumeric.Text, ".") = Len(txtNumeric.Text) - m_DecimalPlaces Then
                        KeyAscii = 0
                        Beep
                    End If
                End If
            End If
        End Select
    End If
End Sub
Private Sub txtNumeric_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub txtNumeric_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub txtNumeric_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub txtNumeric_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    txtNumeric.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtNumeric.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtNumeric.Enabled = PropBag.ReadProperty("Enabled", True)
    Set txtNumeric.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    txtNumeric.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
'    txtNumeric.Text = PropBag.ReadProperty("Value", "")
'    txtNumeric.Alignment = PropBag.ReadProperty("Alignment", 0)
    txtNumeric.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    txtNumeric.Alignment = PropBag.ReadProperty("Alignment", 0)
    m_Maskformat = PropBag.ReadProperty("Maskformat", m_def_Maskformat)
    m_Maskformat = PropBag.ReadProperty("Maskformat", m_def_Maskformat)
    m_IncludeDecimal = PropBag.ReadProperty("IncludeDecimal", m_def_IncludeDecimal)
    m_DecimalPlaces = PropBag.ReadProperty("DecimalPlaces", m_def_DecimalPlaces)
    m_First = PropBag.ReadProperty("First", m_def_First)
    m_Last = PropBag.ReadProperty("Last", m_def_Last)
'    m_Value = PropBag.ReadProperty("Value", m_def_Value)
'    m_Value = PropBag.ReadProperty("Value", m_def_Value)
'    m_Value = PropBag.ReadProperty("Value", m_def_Value)
'    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    txtNumeric.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtNumeric.SelStart = PropBag.ReadProperty("SelStart", 0)
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", txtNumeric.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", txtNumeric.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", txtNumeric.Enabled, True)
    Call PropBag.WriteProperty("Font", txtNumeric.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", txtNumeric.BorderStyle, 1)
'    Call PropBag.WriteProperty("Value", txtNumeric.Text, "")
'    Call PropBag.WriteProperty("Alignment", txtNumeric.Alignment, 0)
    Call PropBag.WriteProperty("MaxLength", txtNumeric.MaxLength, 0)
    Call PropBag.WriteProperty("Alignment", txtNumeric.Alignment, 0)
    Call PropBag.WriteProperty("Maskformat", m_Maskformat, m_def_Maskformat)
    Call PropBag.WriteProperty("Maskformat", m_Maskformat, m_def_Maskformat)
    Call PropBag.WriteProperty("IncludeDecimal", m_IncludeDecimal, m_def_IncludeDecimal)
    Call PropBag.WriteProperty("DecimalPlaces", m_DecimalPlaces, m_def_DecimalPlaces)
    Call PropBag.WriteProperty("First", m_First, m_def_First)
    Call PropBag.WriteProperty("Last", m_Last, m_def_Last)
'    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
'    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
'    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
'    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("SelLength", txtNumeric.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", txtNumeric.SelStart, 0)
End Sub
''MappingInfo=txtNumeric,txtNumeric,-1,Text
'Public Property Get Value() As String
'    Value = txtNumeric.Text
'End Property
'Public Property Let Value(ByVal New_Value As String)
'    txtNumeric.Text() = New_Value
'    PropertyChanged "Value"
'End Property
'MappingInfo=txtNumeric,txtNumeric,-1,Alignment
'Public Property Get Alignment() As Integer
'    Alignment = txtNumeric.Alignment
'End Property
'Public Property Let Alignment(ByVal New_Alignment As Integer)
'    txtNumeric.Alignment() = New_Alignment
'    PropertyChanged "Alignment"
'End Property
'MappingInfo=txtNumeric,txtNumeric,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = txtNumeric.MaxLength
End Property
Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtNumeric.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property
'MappingInfo=txtNumeric,txtNumeric,-1,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = txtNumeric.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    txtNumeric.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Maskformat = m_def_Maskformat
    m_IncludeDecimal = m_def_IncludeDecimal
    m_DecimalPlaces = m_def_DecimalPlaces
    m_First = m_def_First
    m_Last = m_def_Last
'    m_Value = m_def_Value
'    m_Value = m_def_Value
'    m_Value = m_def_Value
'    m_Value = m_def_Value
End Sub
'MemberInfo=13,0,0,
Public Property Get Maskformat() As String
    Maskformat = m_Maskformat
End Property
Public Property Let Maskformat(ByVal New_Maskformat As String)
    m_Maskformat = New_Maskformat
    PropertyChanged "Maskformat"
End Property
'MemberInfo=0,0,0,1
Public Property Get IncludeDecimal() As Boolean
    IncludeDecimal = m_IncludeDecimal
End Property

Public Property Let IncludeDecimal(ByVal New_IncludeDecimal As Boolean)
    m_IncludeDecimal = New_IncludeDecimal
    PropertyChanged "IncludeDecimal"
End Property
'MemberInfo=7,0,0,0
Public Property Get DecimalPlaces() As Integer
    DecimalPlaces = m_DecimalPlaces
End Property
Public Property Let DecimalPlaces(ByVal New_DecimalPlaces As Integer)
    m_DecimalPlaces = New_DecimalPlaces
    PropertyChanged "DecimalPlaces"
End Property
Private Sub txtNumeric_Change()
    RaiseEvent Change
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get First() As Boolean
    First = m_First
End Property

Public Property Let First(ByVal New_First As Boolean)
    m_First = New_First
    PropertyChanged "First"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Last() As Boolean
    Last = m_Last
End Property

Public Property Let Last(ByVal New_Last As Boolean)
    m_Last = New_Last
    PropertyChanged "Last"
End Property
'''
''''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''''MemberInfo=14,0,0,0
'''Public Property Get Value() As Variant
'''    Value = m_Value
'''End Property
'''
'''Public Property Let Value(ByVal New_Value As Variant)
'''    m_Value = New_Value
'''    PropertyChanged "Value"
'''End Property
'''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=13,0,0,
''Public Property Get Value() As String
''    Value = m_Value
''End Property
''
''Public Property Let Value(ByVal New_Value As String)
''    m_Value = New_Value
''    PropertyChanged "Value"
''End Property
''
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=13,0,0,
'Public Property Get Value() As String
'    Value = m_Value
'End Property
'
'Public Property Let Value(ByVal New_Value As String)
'    m_Value = New_Value
'    PropertyChanged "Value"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
'Public Property Get Value() As String
'    Value = m_Value
'End Property
'
'Public Property Let Value(ByVal New_Value As String)
'    m_Value = New_Value
'    PropertyChanged "Value"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNumeric,txtNumeric,-1,Text
Public Property Get Value() As String
Attribute Value.VB_Description = "Returns/sets the text contained in the control."
    Value = txtNumeric.Text
End Property

Public Property Let Value(ByVal New_Value As String)
    txtNumeric.Text() = New_Value
    PropertyChanged "Value"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNumeric,txtNumeric,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = txtNumeric.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtNumeric.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNumeric,txtNumeric,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = txtNumeric.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtNumeric.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property


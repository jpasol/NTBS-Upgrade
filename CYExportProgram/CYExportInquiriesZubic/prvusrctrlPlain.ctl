VERSION 5.00
Begin VB.UserControl prvusrctrlPlain 
   ClientHeight    =   1245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3930
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
   ScaleHeight     =   1245
   ScaleWidth      =   3930
   ToolboxBitmap   =   "prvusrctrlPlain.ctx":0000
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   450
      Left            =   0
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "prvusrctrlPlain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim PreviousBackColor As Variant
Dim PreviousForeColor As Variant
'Event Declarations:

Event Click() ' FLR Feb 2 ' **** MappingInfo=Text1,Text1,-1,Click
Event DblClick() ' FLR Feb 2 ' **** MappingInfo=Text1,Text1,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) ' FLR Feb 2 ' **** MappingInfo=Text1,Text1,-1,KeyDown
Event KeyPress(KeyAscii As Integer) ' FLR Feb 2 ' **** MappingInfo=Text1,Text1,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) ' FLR Feb 2 ' **** MappingInfo=Text1,Text1,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) ' FLR Feb 2 ' **** MappingInfo=Text1,Text1,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) ' FLR Feb 2 ' **** MappingInfo=Text1,Text1,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) ' FLR Feb 2 ' **** MappingInfo=Text1,Text1,-1,MouseUp
Event Change() ' FLR Feb 2 ' **** MappingInfo=Text1,Text1,-1,Change

'Default Property Values:

Const m_def_First = 0
Const m_def_Last = 0

'Property Variables:
Dim m_First As Boolean
Dim m_Last As Boolean
Private Sub Text1_GotFocus()
PreviousBackColor = Text1.BackColor
PreviousForeColor = Text1.ForeColor
Text1.BackColor = &HFFFFFF

'Text1.SelStart = 0
'Text1.SelLength = 30

End Sub

Private Sub Text1_LostFocus()
Text1.BackColor = &H8000000F
End Sub

Private Sub UserControl_Initialize()
Text1.Width = UserControl.Width
Text1.Height = UserControl.Height
End Sub

Private Sub UserControl_Resize()
Text1.Width = UserControl.Width
Text1.Height = UserControl.Height
End Sub
'MappingInfo=Text1,Text1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Text1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Text1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
'MappingInfo=Text1,Text1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Text1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Text1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
'MappingInfo=Text1,Text1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Text1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Text1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
'MappingInfo=Text1,Text1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Text1.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set Text1.Font = New_Font
    PropertyChanged "Font"
End Property
'MappingInfo=Text1,Text1,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Text1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Text1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
'MappingInfo=Text1,Text1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    Text1.Refresh
End Sub

Private Sub Text1_Click()
    RaiseEvent Click
End Sub

Private Sub Text1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Text1_Change()
    RaiseEvent Change
End Sub
'MappingInfo=Text1,Text1,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = Text1.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    Text1.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property
'MappingInfo=Text1,Text1,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = Text1.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    Text1.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property
'MappingInfo=Text1,Text1,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = Text1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Text1.Text() = New_Text
    PropertyChanged "Text"
End Property
'MappingInfo=Text1,Text1,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = Text1.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    Text1.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property
'MappingInfo=Text1,Text1,-1,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = Text1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    Text1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Text1.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Text1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Text1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Text1.SelLength = PropBag.ReadProperty("SelLength", 0)
    Text1.SelStart = PropBag.ReadProperty("SelStart", 0)
    Text1.Text = PropBag.ReadProperty("Text", "")
    Text1.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Text1.Alignment = PropBag.ReadProperty("Alignment", 0)
    m_First = PropBag.ReadProperty("First", m_def_First)
    m_Last = PropBag.ReadProperty("Last", m_def_Last)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", Text1.Enabled, True)
    Call PropBag.WriteProperty("Font", Text1.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", Text1.BorderStyle, 1)
    Call PropBag.WriteProperty("SelLength", Text1.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", Text1.SelStart, 0)
    Call PropBag.WriteProperty("Text", Text1.Text, "")
    Call PropBag.WriteProperty("MaxLength", Text1.MaxLength, 0)
    Call PropBag.WriteProperty("Alignment", Text1.Alignment, 0)
    Call PropBag.WriteProperty("First", m_First, m_def_First)
    Call PropBag.WriteProperty("Last", m_Last, m_def_Last)
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 40
        If m_Last = 0 Then
            SendKeys "{Tab}", True
        End If
    Case 38
        If m_First = 0 Then
            SendKeys "+{Tab}", True
        End If
End Select
End Sub
'MemberInfo=0,0,0,0
Public Property Get First() As Boolean
    First = m_First
End Property

Public Property Let First(ByVal New_First As Boolean)
    m_First = New_First
    PropertyChanged "First"
End Property
'MemberInfo=0,0,0,0
Public Property Get Last() As Boolean
    Last = m_Last
End Property

Public Property Let Last(ByVal New_Last As Boolean)
    m_Last = New_Last
    PropertyChanged "Last"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_First = m_def_First
    m_Last = m_def_Last
End Sub
'MappingInfo=Text1,Text1,-1,MultiLine
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Returns/sets a value that determines whether a control can accept multiple lines of text."
    MultiLine = Text1.MultiLine
End Property
'MappingInfo=Text1,Text1,-1,ScrollBars
Public Property Get ScrollBars() As Integer
Attribute ScrollBars.VB_Description = "Returns/sets a value indicating whether an object has vertical or horizontal scroll bars."
    ScrollBars = Text1.ScrollBars
End Property


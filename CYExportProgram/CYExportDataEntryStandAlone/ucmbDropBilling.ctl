VERSION 5.00
Begin VB.UserControl ucmbDropBilling 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2505
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   420
   ScaleWidth      =   2505
   ToolboxBitmap   =   "ucmbDropBilling.ctx":0000
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "ucmbDropBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'FLR March 3, 1999

Option Explicit

'Event Declarations:
Event Click() 'MappingInfo=Combo1,Combo1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=Combo1,Combo1,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Combo1,Combo1,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Combo1,Combo1,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Combo1,Combo1,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event Change() 'MappingInfo=Combo1,Combo1,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
'Default Property Values:
Const m_def_Last = 0
Const m_def_First = 0
'Property Variables:
Dim m_Last As Boolean
Dim m_First As Boolean
Sub SmartType(c As Object, n As Integer)
    Dim i As Integer
    Dim l As Integer
    Dim s As Integer
    Dim t As String
    s = Combo1.SelStart
    l = Combo1.SelLength
    t = Combo1.Text
    t = Left(t, s) & Chr(n) & Right(t, Len(t) - s)
    s = s + 1
    i = 0
    Do While (i < Combo1.ListCount - 1) _
        And StrComp(Left(t, s), Left(Combo1.List(i), s), vbTextCompare) <> 0
        i = i + 1
    Loop
    If UCase(Left(t, s)) = UCase(Left(Combo1.List(i), s)) Then
        t = Combo1.List(i)
        l = Len(t) - s
    Else
        t = Left(t, s)
        l = 0
    End If
    Combo1.Text = t
    Combo1.SelStart = s
    Combo1.SelLength = l
End Sub
Private Sub Combo1_GotFocus()
    Combo1.BackColor = &HFFFFFF
    Combo1.SelStart = 0
    Combo1.SelLength = Len(Combo1.Text)
    SendKeys "%{down}", True
End Sub
Private Sub Combo1_LostFocus()
    Combo1.BackColor = &H8000000F
End Sub
Private Sub UserControl_Initialize()
    UserControl.height = Combo1.height
    Combo1.width = UserControl.width
End Sub
Private Sub UserControl_Resize()
    UserControl.height = Combo1.height
    Combo1.width = UserControl.width
End Sub
'MappingInfo=Combo1,Combo1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Combo1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Combo1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
'MappingInfo=Combo1,Combo1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Combo1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Combo1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Combo1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Combo1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
'MappingInfo=Combo1,Combo1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Combo1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Combo1.Font = New_Font
    PropertyChanged "Font"
End Property
'MappingInfo=Combo1,Combo1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    Combo1.Refresh
End Sub

Private Sub Combo1_Click()
    RaiseEvent Click
End Sub

Private Sub Combo1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If Chr(KeyAscii) >= " " And Chr(KeyAscii) <= "~" Then
        SmartType Combo1, KeyAscii
        KeyAscii = 0
    End If
End Sub
Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
'MappingInfo=Combo1,Combo1,-1,AddItem
Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "Adds an item to a Listbox or ComboBox control or a row to a Grid control."
    Combo1.AddItem Item, Index
End Sub
Private Sub Combo1_Change()
    RaiseEvent Change
End Sub
'MappingInfo=Combo1,Combo1,-1,Clear
Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of a control or the system Clipboard."
    Combo1.Clear
End Sub
''MappingInfo=Combo1,Combo1,-1,DataFormat
'Public Property Get DataFormat() As IStdDataFormatDisp
'    Set DataFormat = Combo1.DataFormat
'End Property
'Public Property Set DataFormat(ByVal New_DataFormat As IStdDataFormatDisp)
'    Set Combo1.DataFormat = New_DataFormat
'    PropertyChanged "DataFormat"
'End Property
'MappingInfo=Combo1,Combo1,-1,DataMember
Public Property Get DataMember() As String
Attribute DataMember.VB_Description = "Returns/sets a value that describes the DataMember for a data connection."
    DataMember = Combo1.DataMember
End Property

Public Property Let DataMember(ByVal New_DataMember As String)
    Combo1.DataMember() = New_DataMember
    PropertyChanged "DataMember"
End Property
''MappingInfo=Combo1,Combo1,-1,DataSource
'Public Property Get DataSource() As DataSource
'    Set DataSource = Combo1.DataSource
'End Property
'
'Public Property Set DataSource(ByVal New_DataSource As DataSource)
'    Set Combo1.DataSource = New_DataSource
'    PropertyChanged "DataSource"
'End Property
'MappingInfo=Combo1,Combo1,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = Combo1.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Combo1.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property
'MappingInfo=Combo1,Combo1,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = Combo1.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    Combo1.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property
'MappingInfo=Combo1,Combo1,-1,ListCount
Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = Combo1.ListCount
End Property
'MappingInfo=Combo1,Combo1,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = Combo1.Text
End Property
Public Property Let Text(ByVal New_Text As String)
    Combo1.Text() = New_Text
    PropertyChanged "Text"
End Property
'MappingInfo=Combo1,Combo1,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = Combo1.SelLength
End Property
Public Property Let SelLength(ByVal New_SelLength As Long)
    Combo1.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property
'MappingInfo=Combo1,Combo1,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = Combo1.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    Combo1.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property
'MappingInfo=Combo1,Combo1,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = Combo1.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    Combo1.SelText() = New_SelText
    PropertyChanged "SelText"
End Property
'MappingInfo=Combo1,Combo1,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = Combo1.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    Combo1.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'    Combo1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
'    Combo1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Combo1.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Combo1.Font = PropBag.ReadProperty("Font", Ambient.Font)
'    Set DataFormat = PropBag.ReadProperty("DataFormat", Nothing)
    Combo1.DataMember = PropBag.ReadProperty("DataMember", "")
'    Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
    Combo1.FontBold = PropBag.ReadProperty("FontBold", 0)
    Combo1.FontItalic = PropBag.ReadProperty("FontItalic", 0)
'    Combo1.FontName = PropBag.ReadProperty("FontName", "")
'    Combo1.FontSize = PropBag.ReadProperty("FontSize", 0)
'    Combo1.ListIndex = PropBag.ReadProperty("ListIndex", 0)
    Combo1.Text = PropBag.ReadProperty("Text", "Combo1")
    Combo1.SelLength = PropBag.ReadProperty("SelLength", 0)
    Combo1.SelStart = PropBag.ReadProperty("SelStart", 0)
    Combo1.SelText = PropBag.ReadProperty("SelText", "")
    Combo1.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_Last = PropBag.ReadProperty("Last", m_def_Last)
    m_First = PropBag.ReadProperty("First", m_def_First)
'    Combo1.ListIndex = PropBag.ReadProperty("ListIndex", 0)
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'    Call PropBag.WriteProperty("BackColor", Combo1.BackColor, &H80000005)
'    Call PropBag.WriteProperty("ForeColor", Combo1.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", Combo1.Enabled, True)
    Call PropBag.WriteProperty("Font", Combo1.Font, Ambient.Font)
'    Call PropBag.WriteProperty("DataFormat", DataFormat, Nothing)
    Call PropBag.WriteProperty("DataMember", Combo1.DataMember, "")
'   Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("FontBold", Combo1.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", Combo1.FontItalic, 0)
'    Call PropBag.WriteProperty("FontName", Combo1.FontName, "")
'    Call PropBag.WriteProperty("FontSize", Combo1.FontSize, 0)
'    Call PropBag.WriteProperty("ListIndex", Combo1.ListIndex, 0)
    Call PropBag.WriteProperty("Text", Combo1.Text, "Combo1")
    Call PropBag.WriteProperty("SelLength", Combo1.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", Combo1.SelStart, 0)
    Call PropBag.WriteProperty("SelText", Combo1.SelText, "")
    Call PropBag.WriteProperty("ToolTipText", Combo1.ToolTipText, "")
    Call PropBag.WriteProperty("Last", m_Last, m_def_Last)
    Call PropBag.WriteProperty("First", m_First, m_def_First)
    Call PropBag.WriteProperty("ListIndex", Combo1.ListIndex, 0)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Last() As Boolean
    Last = m_Last
End Property

Public Property Let Last(ByVal New_Last As Boolean)
    m_Last = New_Last
    PropertyChanged "Last"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get First() As Boolean
    First = m_First
End Property

Public Property Let First(ByVal New_First As Boolean)
    m_First = New_First
    PropertyChanged "First"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Last = m_def_Last
    m_First = m_def_First
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ListIndex
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    ListIndex = Combo1.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    Combo1.ListIndex() = New_ListIndex
    PropertyChanged "ListIndex"
End Property


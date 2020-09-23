VERSION 5.00
Begin VB.UserControl UltiFrame 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2580
   ControlContainer=   -1  'True
   ForwardFocus    =   -1  'True
   ScaleHeight     =   1560
   ScaleWidth      =   2580
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "UltiFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'************************************************'
'*                 Public Enums                 *'
'************************************************'
Public Enum ufAlignment     ' Returns or sets a value that determines how the caption of the control will be aligned.
    ufTopLeft = 0           ' (Default) Left align text - top of frame.
    ufTopCenter = 1         ' Center align text - top of frame.
    ufTopRight = 2          ' Right text - top of frame.
    ufBottomLeft = 3        ' Right align text - bottom of frame.
    ufBottomCenter = 4      ' Center align text - bottom of frame.
    ufBottomRight = 5       ' Left align text - bottom of frame.
End Enum

Public Enum ufAppearance    ' Returns or sets a value that specifies weather the frame will be drawn 2D or 3D.
    ufFlat = 0              ' The frame appears as a 2D flat rectangle.
    ufEtched = 1            ' (Default) The frame appears sunken into the background.
    ufBump = 2              ' The frame appears raised above the background.
End Enum

Public Enum ufBackStyle     ' Returns or sets a value that determines whether the background of the control will be opaque or transparent.
   ufTransparent = 0        ' Transparent background. The area of the control will allow what is behind the control to show through.
   ufOpaque = 1             ' (Default) Opaque background. The area of the control will be filled with the control’s BackColor.
End Enum

Public Enum ufBorderStyle   ' Returns or sets a value that specifes how the 2D border will be drawn.
    ufSolid = 0             ' (Default) The pen is solid.
    ufDash = 1              ' The pen is dashed. This style is valid only when the pen width is one or less in device units.
    ufDot = 2               ' The pen is dotted. This style is valid only when the pen width is one or less in device units.
    ufDashDot = 3           ' The pen has alternating dashes and dots. This style is valid only when the pen width is one or less in device units.
    ufDashDotDot = 4        ' The pen has alternating dashes and double dots. This style is valid only when the pen width is one or less in device units.
End Enum

Public Enum ufCaptionStyle  ' Returns or sets a value that specifies how caption text will be displayed on the control.
    ufStandard = 0          ' Default) Standard.  The caption displays as static text on a single line.
    ufWrapped = 1           ' Wrapped.  The caption displays as static text on multiple lines.
End Enum

Public Enum ufFont3D        ' Returns or sets a value that specifies the 3-D style of the control’s caption text.
    ufNoneFont3D = 0        ' (Default) None.  Caption is displayed flat (not 3-dimensional).
    ufRaisedLight = 1       ' Raised w/light shading.  Caption appears as if it is raised slightly above the background.
    ufRaisedHeavy = 2       ' Raised w/heavy shading.  Caption appears even more raised.
    ufInsetLight = 3        ' Inset w/light shading.  Caption appears as if it is inset slightly into the background.
    ufInsetHeavy = 4        ' Inset w/heavy shading.  Caption appears even more inset.
    ufDropShadow = 5        ' Drop Shadow. Caption appears with a dark gray drop shadow slightly below and to the right of the text.
End Enum

Public Enum ufMousePointer  ' Returns or sets a value specifying the type of mouse pointer displayed when the mouse is over a particular part of an object at run time.
    ufDefault = 0           ' (Default) Shape determined by the object.
    ufArrow = 1             ' Arrow.
    ufCrosshair = 2         ' Cross (cross-hair pointer).
    ufIBeam = 3             ' I Beam.
    ufIconPointer = 4       ' Icon (small square within a square).
    ufSizePointer = 5       ' Size (four-pointed arrow pointing north, south, east, and west).
    ufSizeNESW = 6          ' Size NE SW (double arrow pointing northeast and southwest).
    ufSizeNS = 7            ' Size N S (double arrow pointing north and south).
    ufSizeNWSE = 8          ' Size NW, SE (double arrow pointing northwest and southeast).
    ufSizeWE = 9            ' Size WE (double arrow pointing west and east).
    ufUpArrow = 10          ' Up Arrow.
    ufHourglass = 11        ' Hourglass (wait).
    ufNoDrop = 12           ' No Drop.
    ufArrowHourglass = 13   ' Arrow and hourglass.
    ufArrowQuestion = 14    ' Arrow and question mark.
    ufSizeAll = 15          ' Size all.
    ufCustom = 99           ' Custom icon specified by the MouseIcon property.
End Enum

Public Enum ufOLEDropMode   ' Returns or sets a value that determines whether the control can be a drop target for OLE drag-and-drop operations.
    ufOLEDropNone = 0       ' (Default) None. The target component does not accept OLE drops and displays the No Drop cursor.
    ufOLEDropManual = 1     ' Manual. The target component triggers the OLE drop events, allowing the programmer to handle the OLE drop operation in code.
    ufOLEDropAutomatic = 2  ' Automatic. The target component automatically accepts OLE drops if the DataObject object contains data in a format it recognizes. No mouse or OLE drag/drop events on the target will occur when OLEDropMode is set to ufOLEDropAutomatic.
End Enum

Public Enum ufSoundType
    ufPlaySoundFile = 0
    ufPlaySystemSound = 1
End Enum

'************************************************'
'*              Property Variables              *'
'************************************************'
Private m_Alignment As ufAlignment
Private m_Appearance As ufAppearance
Private m_BackColor As OLE_COLOR
Private m_BackStyle As ufBackStyle
Private m_BorderColor As OLE_COLOR
Private m_BorderHighlightColor As OLE_COLOR
Private m_BorderShadowColor As OLE_COLOR
Private m_BorderStyle As ufBorderStyle
Private m_Caption As String
Private m_CaptionStyle As ufCaptionStyle
Private m_ClipControls As Boolean
Private m_CornerRadius As Long
Private m_Enabled As Boolean
Private m_Font3D As ufFont3D
Private m_ForeColor As OLE_COLOR
Private m_hWnd As Long
Private m_MousePointer As ufMousePointer
Private m_OLEDropMode As ufOLEDropMode

'************************************************'
'*        Properties with Default-Values        *'
'************************************************'
Private Const m_def_Alignment As Long = ufTopLeft
Private Const m_def_Appearance As Long = ufEtched
Private Const m_def_BackColor As Long = vbButtonFace
Private Const m_def_BackStyle As Long = ufOpaque
Private Const m_def_BorderColor As Long = vbWindowFrame
Private Const m_def_BorderHighLightColor As Long = vb3DHighlight
Private Const m_def_BorderShadowColor As Long = vbButtonShadow
Private Const m_def_BorderStyle As Long = ufSolid
Private Const m_def_Caption As String = vbNullString
Private Const m_def_CaptionStyle As Long = ufStandard
Private Const m_def_ClipControls As Boolean = True
Private Const m_def_CornerRadius As Long = 0
Private Const m_def_Enabled As Boolean = True
Private Const m_def_Font3D As Long = ufNoneFont3D
Private Const m_def_ForeColor As Long = vbButtonText
Private Const m_def_MousePointer As Long = ufDefault
Private Const m_def_OLEDropMode As Long = ufOLEDropNone

'************************************************'
'*                 Local Variables              *'
'************************************************'
Private m_rcFrame As RECT
Private m_rcCaption As RECT
Private m_siCaption As Size

'************************************************'
'*           Default Control Constants          *'
'************************************************'
Private Const m_def_Height As Long = 750    ' Height in pixels
Private Const m_def_Width As Long = 1500    ' Width in pixles

'************************************************'
'*                    Events                    *'
'************************************************'
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user clicks the left mouse button and releases it over the control or, in the case of the SSCommand and SSCheck controls, when the user presses the spacebar while the control has focus."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user double-clicks the mouse while the mouse pointer is over the control."
Attribute DblClick.VB_UserMemId = -601
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses a mouse button while the mouse pointer is within the boundary of the control."
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Occurs when the mouse pointer is moved while within the boundaries of the control."
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases a mouse button while the mouse pointer is within the boundaries of the control."
Attribute MouseUp.VB_UserMemId = -607
Public Event OLECompleteDrag(Effect As Long)
Attribute OLECompleteDrag.VB_Description = "Occurs when a source component is dropped onto a target component, informing the source component that a drag action was either performed or canceled."
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute OLEDragDrop.VB_Description = "Occurs when a source component is dropped onto a target component when the source component determines that a drop can occur."
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when one component is dragged over another."
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Attribute OLEGiveFeedback.VB_Description = "Occurs after every OLEDragOver event. OLEGiveFeedback allows the source component to provide visual feedback to the user, such as changing the mouse cursor to indicate what will happen if the user drops the object, or provide visual feedback on the selec"
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Attribute OLESetData.VB_Description = "Occurs on a source component when a target component performs the GetData method on the source’s ssDataObject object, but the data for the specified format has not yet been loaded."
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Attribute OLEStartDrag.VB_Description = "Occurs when a component's OLEDrag method is performed. This event specifies the data formats and drop effects that the source component supports. It can also be used to insert data into the ssDataObject object."
Public Event Resize()
Attribute Resize.VB_Description = "Event fired when the control is resized."

Public Property Get Alignment() As ufAlignment
Attribute Alignment.VB_Description = "Returns or sets a value that determines how the caption of the control will be aligned."
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal vNewValue As ufAlignment)
    m_Alignment = vNewValue
    PropertyChanged "Alignment"
    DrawControl
End Property

Public Property Get Appearance() As ufAppearance
Attribute Appearance.VB_Description = "Returns or sets a value that specifies how the frame will be drawn."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal vNewValue As ufAppearance)
    m_Appearance = vNewValue
    PropertyChanged "Appearance"
    DrawControl
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns or sets the background color of the control."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    m_BackColor = vNewValue
    picFrame.BackColor = m_BackColor
    PropertyChanged "BackColor"
    DrawControl
End Property

Public Property Get BackStyle() As ufBackStyle
Attribute BackStyle.VB_Description = "Returns or sets a value that determines whether the background of the control will be opaque or transparent."
Attribute BackStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackStyle.VB_UserMemId = -502
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal vNewValue As ufBackStyle)
    m_BackStyle = vNewValue
    UserControl.BackStyle = m_BackStyle
    PropertyChanged "BackStyle"
    DrawControl
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns or sets the color of the broder when Appearance is ufFlat"
Attribute BorderColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal vNewValue As OLE_COLOR)
    m_BorderColor = vNewValue
    PropertyChanged "BorderColor"
    DrawControl
End Property

Public Property Get BorderHighLightColor() As OLE_COLOR
Attribute BorderHighLightColor.VB_Description = "Returns or sets the Highlight color of the broder when Appearance is ufEtched or ufBump."
Attribute BorderHighLightColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderHighLightColor = m_BorderHighlightColor
End Property

Public Property Let BorderHighLightColor(ByVal vNewValue As OLE_COLOR)
    m_BorderHighlightColor = vNewValue
    PropertyChanged "BorderHighLightColor"
    DrawControl
End Property

Public Property Get BorderShadowColor() As OLE_COLOR
Attribute BorderShadowColor.VB_Description = "Returns or sets the Shadow color of the broder when Appearance is ufEtched or ufBump."
Attribute BorderShadowColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderShadowColor = m_BorderShadowColor
End Property

Public Property Let BorderShadowColor(ByVal vNewValue As OLE_COLOR)
    m_BorderShadowColor = vNewValue
    PropertyChanged "BorderShadowColor"
    DrawControl
End Property

Public Property Get BorderStyle() As ufBorderStyle
Attribute BorderStyle.VB_Description = "Returns or sets a value that specifies how the border will be drawn when Appearance if ufFlat"
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal vNewValue As ufBorderStyle)
    m_BorderStyle = vNewValue
    PropertyChanged "BorderStyle"
    DrawControl
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns or sets the caption text of the control."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
    m_Caption = vNewValue
    PropertyChanged "Caption"
    DrawControl
End Property

Public Property Get CaptionStyle() As ufCaptionStyle
Attribute CaptionStyle.VB_Description = "Returns or sets a value that specifies how caption text will be displayed on the control.\r\n\r\n"
Attribute CaptionStyle.VB_ProcData.VB_Invoke_Property = ";Text"
    CaptionStyle = m_CaptionStyle
End Property

Public Property Let CaptionStyle(ByVal vNewValue As ufCaptionStyle)
    m_CaptionStyle = vNewValue
    PropertyChanged "CaptionStyle"
    DrawControl
End Property

Public Property Get ClipControls() As Boolean
Attribute ClipControls.VB_Description = "Returns or sets a value that determines whether graphics methods in Paint events repaint the entire object or only newly exposed areas.  Also determines whether the Microsoft Windows operating environment creates a clipping region that excludes nongraphi"
Attribute ClipControls.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ClipControls = m_ClipControls
End Property

Public Property Let ClipControls(ByVal vNewValue As Boolean)
    m_ClipControls = vNewValue
    PropertyChanged "ClipControls"
End Property

Public Property Get CornerRadius() As Long
Attribute CornerRadius.VB_Description = "Returns or sets a value that specifies the radius of the corners of the frame."
Attribute CornerRadius.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CornerRadius = m_CornerRadius
End Property

Public Property Let CornerRadius(ByVal vNewValue As Long)
    m_CornerRadius = vNewValue
    PropertyChanged "CornerRadius"
    DrawControl
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns or sets a value that determines whether the object can be selected by the user."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    m_Enabled = vNewValue
    UserControl.Enabled = m_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns or sets the properties of the Font object at design time and run time."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = picFrame.Font
End Property

Public Property Set Font(ByVal vNewValue As Font)
    Set picFrame.Font = vNewValue
    PropertyChanged "Font"
    DrawControl
End Property

Public Property Get Font3D() As ufFont3D
Attribute Font3D.VB_Description = "Returns or sets a value that specifies the 3-D style of the control’s caption text."
Attribute Font3D.VB_ProcData.VB_Invoke_Property = ";Font"
    Font3D = m_Font3D
End Property

Public Property Let Font3D(ByVal vNewValue As ufFont3D)
    m_Font3D = vNewValue
    PropertyChanged "Font3D"
    DrawControl
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns or sets the foreground (text) color of the control."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
    m_ForeColor = vNewValue
    picFrame.ForeColor = m_ForeColor
    PropertyChanged "ForeColor"
    DrawControl
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns the window handle of the control."
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Returns or sets the icon that will be used for the mouse pointer when the MousePointer property is set to Custom."
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Picture"
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal vNewValue As Picture)
    Set UserControl.MouseIcon = vNewValue
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As ufMousePointer
Attribute MousePointer.VB_Description = "Returns or sets a value specifying the type of mouse pointer displayed when the mouse is over a particular part of an object at run time."
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal vNewValue As ufMousePointer)
    m_MousePointer = vNewValue
    UserControl.MousePointer = m_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get OLEDropMode() As ufOLEDropMode
Attribute OLEDropMode.VB_Description = "Returns or sets a value that determines whether the control can be a drop target for OLE drag-and-drop operations"
Attribute OLEDropMode.VB_ProcData.VB_Invoke_Property = ";Behavior"
    OLEDropMode = m_OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal vNewValue As ufOLEDropMode)
    m_OLEDropMode = vNewValue
    UserControl.OLEDropMode = m_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Public Sub About()
Attribute About.VB_Description = "Displays version information about the control."
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
    frmAbout.Show vbModal
End Sub

Public Sub AboutBox()
Attribute AboutBox.VB_Description = "Displays version information about the control."
    frmAbout.Show vbModal
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "This method causes the control to enter OLE Drag mode. OLE Drag mode ends when mouse button is released."
    UserControl.OLEDrag
End Sub

Public Sub PlaySoundFile(ByVal lpSound As String, _
                         Optional lFlag As ufSoundType = ufPlaySoundFile)
Attribute PlaySoundFile.VB_Description = "This method causes the control to play the sound file specified."

Dim dwFlag As Long

dwFlag = IIf(lFlag, SND_ALIAS, SND_FILENAME) Or SND_ASYNC Or SND_NODEFAULT Or SND_PURGE
PlaySound lpSound, 0&, dwFlag

End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_InitProperties()
    m_Alignment = m_def_Alignment
    m_Appearance = m_def_Appearance
    m_BackColor = m_def_BackColor
    m_BackStyle = m_def_BackStyle
    m_BorderColor = m_def_BorderColor
    m_BorderHighlightColor = m_def_BorderHighLightColor
    m_BorderShadowColor = m_def_BorderShadowColor
    m_BorderStyle = m_def_BorderStyle
    m_Caption = Ambient.DisplayName
    m_CaptionStyle = m_def_CaptionStyle
    m_ClipControls = m_def_ClipControls
    m_CornerRadius = m_def_CornerRadius
    m_Enabled = m_def_Enabled
    m_Font3D = m_def_Font3D
    m_ForeColor = m_def_ForeColor
    m_MousePointer = m_def_MousePointer
    m_OLEDropMode = m_def_OLEDropMode
    Set picFrame.Font = Ambient.Font
    Set UserControl.MouseIcon = Nothing
    picFrame.ForeColor = m_ForeColor
    UserControl.BackColor = m_BackColor
    UserControl.BackStyle = m_BackStyle
    UserControl.ClipControls = m_ClipControls
    UserControl.Enabled = m_Enabled
    UserControl.MousePointer = m_MousePointer
    UserControl.OLEDropMode = m_OLEDropMode
    UserControl.Height = m_def_Height
    UserControl.Width = m_def_Width
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_BorderHighlightColor = PropBag.ReadProperty("BorderHighLightColor", m_def_BorderHighLightColor)
    m_BorderShadowColor = PropBag.ReadProperty("BorderShadowColor", m_def_BorderShadowColor)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_ClipControls = PropBag.ReadProperty("ClipControls", m_def_ClipControls)
    m_CornerRadius = PropBag.ReadProperty("CornerRadius", m_def_CornerRadius)
    m_CaptionStyle = PropBag.ReadProperty("CaptionStyle", m_def_CaptionStyle)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_Font3D = PropBag.ReadProperty("Font3D", m_def_Font3D)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_MousePointer = PropBag.ReadProperty("MousePointer", m_def_MousePointer)
    m_OLEDropMode = PropBag.ReadProperty("OLEDropMode", m_def_OLEDropMode)
    Set picFrame.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set UserControl.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    picFrame.ForeColor = m_ForeColor
    UserControl.BackColor = m_BackColor
    UserControl.BackStyle = m_BackStyle
    UserControl.ClipControls = m_ClipControls
    UserControl.Enabled = m_Enabled
    UserControl.MousePointer = m_MousePointer
    UserControl.OLEDropMode = m_OLEDropMode
    DrawControl
End Sub

Private Sub UserControl_Resize()
    picFrame.Height = ScaleHeight
    picFrame.Width = ScaleWidth
    DrawControl
    RaiseEvent Resize
End Sub

Private Sub UserControl_Show()
    DrawControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("BorderHighLightColor", m_BorderHighlightColor, m_def_BorderHighLightColor)
    Call PropBag.WriteProperty("BorderShadowColor", m_BorderShadowColor, m_def_BorderShadowColor)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("CaptionStyle", m_CaptionStyle, m_def_CaptionStyle)
    Call PropBag.WriteProperty("ClipControls", m_ClipControls, m_def_ClipControls)
    Call PropBag.WriteProperty("CornerRadius", m_CornerRadius, m_def_CornerRadius)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", picFrame.Font, Ambient.Font)
    Call PropBag.WriteProperty("Font3D", m_Font3D, m_def_Font3D)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", m_MousePointer, m_def_MousePointer)
    Call PropBag.WriteProperty("OLEDropMode", m_OLEDropMode, m_def_OLEDropMode)
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "This method forces a complete repaint of a control."
Attribute Refresh.VB_UserMemId = -550
    DrawControl
End Sub

Private Sub DrawControl()

Dim rcTemp As RECT

'On Error Resume Next

m_siCaption.cx = 0
m_siCaption.cy = 0

' Get the width and height needed for the caption area.
If Len(m_Caption) > 0 Then
    rcTemp = GetCaptionHeight(picFrame, m_Caption, m_CaptionStyle)
    m_siCaption.cy = Int(rcTemp.Bottom - rcTemp.Top)
    m_siCaption.cx = Int(rcTemp.Right - rcTemp.Left)
End If

' Set the caption area.
SetCaptionBox m_siCaption

' Draw the frame.
Set picFrame.Picture = Nothing
picFrame.Cls
DoFrame m_siCaption.cy

' Draw the caption.
Set picFrame.Picture = Nothing
picFrame.Cls
m_rcCaption.Left = m_rcCaption.Left + 2
m_rcCaption.Right = m_rcCaption.Right - 2
If m_Enabled = False And UserControl.Ambient.UserMode = True Then
    DrawCaption UserControl.hDC, picFrame, m_Caption, m_rcCaption, False, , m_Alignment
Else
    DrawCaption UserControl.hDC, picFrame, m_Caption, m_rcCaption, True, m_Font3D, m_Alignment
End If

With picFrame
    Set .Picture = Nothing
    .Cls
    BitBlt .hDC, 0, 0, .ScaleX(.ScaleWidth, .ScaleMode, vbPixels), .ScaleY(.ScaleHeight, .ScaleMode, vbPixels), UserControl.hDC, 0, 0, vbSrcCopy
End With

' Create a transparent background if needed.
CreateTransparentMask

' Refresh the control.
UserControl.Refresh

End Sub

Private Sub SetCaptionBox(ByRef siCaption As Size)

Dim lOffset As Long

lOffset = Int(m_CornerRadius / 2)

' Position the blank caption area in the correct place.
With m_rcCaption
    Select Case m_Alignment
        Case ufAlignment.ufTopLeft
            .Top = ScaleTop
            .Left = ScaleLeft + lOffset + 5
            .Right = .Left + m_siCaption.cx + 4
            .Bottom = .Top + m_siCaption.cy
            
        Case ufAlignment.ufTopCenter
            .Top = ScaleTop
            .Left = (Int(picFrame.ScaleX(picFrame.ScaleWidth, picFrame.ScaleMode, vbPixels) - m_siCaption.cx - 5) / 2)
            .Right = .Left + m_siCaption.cx + 4
            .Bottom = .Top + m_siCaption.cy
            
        Case ufAlignment.ufTopRight
            .Top = ScaleTop
            .Left = (picFrame.ScaleX(picFrame.ScaleWidth, picFrame.ScaleMode, vbPixels) - m_siCaption.cx) - (lOffset + 9)
            .Right = .Left + m_siCaption.cx + 4
            .Bottom = .Top + m_siCaption.cy
            
        Case ufAlignment.ufBottomLeft
            .Top = picFrame.ScaleY(picFrame.ScaleHeight, picFrame.ScaleMode, vbPixels) - m_siCaption.cy
            .Left = ScaleLeft + lOffset + 5
            .Right = .Left + m_siCaption.cx + 4
            .Bottom = .Top + m_siCaption.cy
            
        Case ufAlignment.ufBottomCenter
            .Top = picFrame.ScaleY(picFrame.ScaleHeight, picFrame.ScaleMode, vbPixels) - m_siCaption.cy
            .Left = (Int(picFrame.ScaleX(picFrame.ScaleWidth, picFrame.ScaleMode, vbPixels) - m_siCaption.cx - 5) / 2)
            .Right = .Left + m_siCaption.cx + 4
            .Bottom = .Top + m_siCaption.cy
            
        Case ufAlignment.ufBottomRight
            .Top = picFrame.ScaleY(picFrame.ScaleHeight, picFrame.ScaleMode, vbPixels) - m_siCaption.cy
            .Left = (picFrame.ScaleX(picFrame.ScaleWidth, picFrame.ScaleMode, vbPixels) - m_siCaption.cx) - (lOffset + 9)
            .Right = .Left + m_siCaption.cx + 4
            .Bottom = .Top + m_siCaption.cy
    End Select
End With

End Sub

Private Sub DoFrame(ByVal lTextHeight As Long)

With picFrame
    ' Set the width and height of the frame.
    Select Case m_Alignment
        Case ufAlignment.ufTopCenter, ufAlignment.ufTopLeft, ufAlignment.ufTopRight
            m_rcFrame.Top = ScaleTop + Int(lTextHeight / 2)
            m_rcFrame.Left = ScaleLeft
            m_rcFrame.Right = .ScaleX(.ScaleWidth, .ScaleMode, vbPixels)
            m_rcFrame.Bottom = .ScaleY(.ScaleHeight, .ScaleMode, vbPixels)
        Case ufAlignment.ufBottomCenter, ufAlignment.ufBottomLeft, ufAlignment.ufBottomRight
            m_rcFrame.Top = ScaleTop
            m_rcFrame.Left = ScaleLeft
            m_rcFrame.Right = .ScaleX(.ScaleWidth, .ScaleMode, vbPixels)
            m_rcFrame.Bottom = .ScaleY(.ScaleHeight, .ScaleMode, vbPixels) - Int(lTextHeight / 2)
    End Select
End With

' Draw the frame based upon the selected appearance.
Select Case m_Appearance
    Case ufAppearance.ufFlat
        Call DrawFlatEdge(UserControl.hDC, picFrame, m_rcFrame, m_rcCaption, m_BorderColor, m_BorderStyle, m_CornerRadius)

    Case ufAppearance.ufEtched
        Call Draw3DEdge(UserControl.hDC, picFrame, m_rcFrame, m_rcCaption, m_BorderHighlightColor, m_BorderShadowColor, m_CornerRadius)

    Case ufAppearance.ufBump
        Call Draw3DEdge(UserControl.hDC, picFrame, m_rcFrame, m_rcCaption, m_BorderShadowColor, m_BorderHighlightColor, m_CornerRadius)
End Select

End Sub

Private Sub CreateTransparentMask()

' Thanks to Stephen Kent for this routine.

Dim ctl As Control
Dim NonTransColor As Long

If (m_BackStyle = ufTransparent) And (Ambient.UserMode) Then
    picFrame.Picture = picFrame.Image
    
    If TranslateColor(picFrame.BackColor) = vbBlack Then
        NonTransColor = vbWhite
    Else
        NonTransColor = vbBlack
    End If
    
    UserControl.MaskColor = picFrame.BackColor

    On Error Resume Next
    
    For Each ctl In ContainedControls
        picFrame.Line (ScaleX(ctl.Left, ScaleMode, vbTwips), ScaleY(ctl.Top, ScaleMode, vbTwips))-Step(ScaleX(ctl.Width, ScaleMode, vbTwips) - ScaleX(1, vbPixels, vbTwips), ScaleY(ctl.Height, ScaleMode, vbTwips) - ScaleY(1, vbPixels, vbTwips)), NonTransColor, BF
    Next

    UserControl.MaskPicture = picFrame.Image
    UserControl.BackStyle = 0
    
    picFrame.Cls
    picFrame.Refresh
Else
    UserControl.BackStyle = 1

End If

End Sub

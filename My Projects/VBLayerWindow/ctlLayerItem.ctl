VERSION 5.00
Begin VB.UserControl ctlLayerItem 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   217
   Begin VB.PictureBox picTopDrop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   2505
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.PictureBox picBottomDrop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   15
      ScaleHeight     =   60
      ScaleWidth      =   2505
      TabIndex        =   1
      Top             =   435
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Image imgLocked 
      Height          =   240
      Left            =   615
      Picture         =   "ctlLayerItem.ctx":0000
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgVisible 
      Height          =   240
      Left            =   150
      Picture         =   "ctlLayerItem.ctx":0102
      Top             =   120
      Width           =   240
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "LayerName"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1050
      TabIndex        =   0
      Top             =   135
      Width           =   2040
   End
   Begin VB.Shape shpName 
      BackStyle       =   1  'Opaque
      Height          =   360
      Left            =   1005
      Top             =   60
      Width           =   2160
   End
   Begin VB.Shape shpLocked 
      BackStyle       =   1  'Opaque
      Height          =   360
      Left            =   525
      Top             =   60
      Width           =   375
   End
   Begin VB.Shape shpVisible 
      BackStyle       =   1  'Opaque
      Height          =   360
      Left            =   60
      Top             =   60
      Width           =   375
   End
End
Attribute VB_Name = "ctlLayerItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************************************************'
'*'
'*' Module    : ctlLayerItem
'*'
'*'
'*' Author    : Joseph M. Ferris <josephmferris@cox.net>
'*'
'*' Date      : 10.17.2002
'*'
'*' Depends   : Visual Basic 6.  Service Pack 5, or higher is recommended.
'*'             mscomct2.ocx
'*'
'*' Purpose   : Provide a single layer item for use in a layer dialog control.  Example:  Adobe Photoshop.
'*'
'*' Notes     : This component is fairly straightforward.  There is minimal functionality in here, except for providing
'*'             the toggling of the custom "check boxes".  Most of the rest of this control acts as an event broker
'*'             for raising events to the LayerWindow.  Private in Scope.
'*'
'**********************************************************************************************************************'

Option Explicit

'**********************************************************************************************************************'
'*'
'*' API Declarations (Private) - USER32
'*'
'*' 1.  LockWindow Update
'*'
'**********************************************************************************************************************'

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private m_bolVisibleChecked     As Boolean
Private m_bolLockedChecked      As Boolean
Private m_bolSelected           As Boolean
Private m_bolCheckImages        As Boolean
Private m_bolTopIndicator       As Boolean
Private m_bolBottomIndicator    As Boolean

'**********************************************************************************************************************'
'*'
'*' Private Constant Declarations
'*'
'*' 1.  DEFAULT_BORDERCOL       - Default border color.
'*'
'**********************************************************************************************************************'

Private Const DEFAULT_BORDERCOL As Long = vbWindowFrame
Private Const DEFAULT_TRACKCOL  As Long = vbHighlight

'**********************************************************************************************************************'
'*'
'*' Private Member Variables for Props
'*'
'**********************************************************************************************************************'

Private m_OLCColor              As OLE_COLOR
Private m_BorderColor           As OLE_COLOR
Private m_TrackingColor         As OLE_COLOR

'**********************************************************************************************************************'
'*'
'*' Event Declarations
'*'
'**********************************************************************************************************************'

Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event LockedToggle(Value As Boolean)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Resize()
Event Show()
Event WriteProperties(PropBag As PropertyBag)
Event ReadProperties(PropBag As PropertyBag)
Event InitProperties()
Event VisibilityToggle(Value As Boolean)

'**********************************************************************************************************************'
'*'
'*' Property Declarations
'*'
'*' 1.  ActiveControl (Get)
'*' 2.  BorderColor (Get, Let)
'*' 3.  BottomIndicator (Get, Let)
'*' 4.  Caption (Get, Let)
'*' 5.  Enabled (Get, Let)
'*' 6.  Forecolor (Get, Let)
'*' 7.  hwnd (Get)
'*' 8.  LockedChecked (Get, Let)
'*' 9.  MouseIcon (Get, Set)
'*' 10. MousePointer (Get, Let)
'*' 11. Selected (Get, Let)
'*' 12. TopIndicator (Get, Let)
'*' 13. TrackingColor (Get, Let)
'*' 14. VisibleChecked (Get, Let)
'*'
'**********************************************************************************************************************'

Public Property Get ActiveControl() As Object
Attribute ActiveControl.VB_Description = "Returns the control that has focus."
    Set ActiveControl = UserControl.ActiveControl
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
    BorderColor = m_BorderColor
End Property
Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    shpLocked.BorderColor = New_BorderColor
    shpName.BorderColor = New_BorderColor
    shpVisible.BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
End Property

Public Property Get BottomIndicator() As Boolean
    BottomIndicator = m_bolBottomIndicator
End Property
Public Property Let BottomIndicator(Value As Boolean)
    m_bolBottomIndicator = Value
    picBottomDrop.Visible = m_bolBottomIndicator
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblName.Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    lblName.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_OLCColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblName.ForeColor() = New_ForeColor
    m_OLCColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

Public Property Get LockedChecked() As Boolean
    LockedChecked = m_bolLockedChecked
End Property
Public Property Let LockedChecked(Value As Boolean)
    m_bolLockedChecked = Value
    If m_bolLockedChecked = True Then
        If m_bolCheckImages = True Then
            imgLocked.Visible = True
        Else
            shpLocked.BackColor = vbBlue
        End If
    Else
        If m_bolCheckImages = True Then
            imgLocked.Visible = False
        Else
            shpLocked.BackColor = vbWhite
        End If
    End If
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As Integer)
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Selected() As Boolean
    Selected = m_bolSelected
End Property
Public Property Let Selected(Value As Boolean)
    m_bolSelected = Value
    If m_bolSelected = True Then
        shpName.BackColor = vbHighlight
        lblName.ForeColor = vbHighlightText
    Else
        shpName.BackColor = vbWindowBackground
        lblName.ForeColor = vbWindowText
    End If
End Property

Public Property Get TopIndicator() As Boolean
    TopIndicator = m_bolTopIndicator
End Property
Public Property Let TopIndicator(Value As Boolean)
    m_bolTopIndicator = Value
    picTopDrop.Visible = m_bolTopIndicator
End Property

Public Property Get TrackingColor() As OLE_COLOR
Attribute TrackingColor.VB_Description = "Color of tracking bar on drag operations."
    TrackingColor = m_TrackingColor
End Property
Public Property Let TrackingColor(ByVal New_TrackingColor As OLE_COLOR)
    m_TrackingColor = New_TrackingColor
    PropertyChanged "TrackingColor"
End Property

Public Property Get VisibleChecked() As Boolean
    VisibleChecked = m_bolVisibleChecked
End Property
Public Property Let VisibleChecked(Value As Boolean)
    m_bolVisibleChecked = Value
    If m_bolVisibleChecked = True Then
        If m_bolCheckImages = True Then
            imgVisible.Visible = True
        Else
            shpVisible.BackColor = vbBlue
        End If
    Else
        If m_bolCheckImages = True Then
            imgVisible.Visible = False
        Else
            shpVisible.BackColor = vbWhite
        End If
    End If
End Property

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_Click
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Event Broker.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_DblClick
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Event Broker.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_InitProperties
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Initialize properties for Defaults.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_InitProperties()
    RaiseEvent InitProperties
    m_BorderColor = DEFAULT_BORDERCOL
    m_TrackingColor = DEFAULT_TRACKCOL
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_KeyDown
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Event Broker.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None
'*'
'**********************************************************************************************************************'
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_KeyPress
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Event Broker.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_KeyUp
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Event Broker.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_MouseDown
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Event Broker.  Also has the ability to check if either of the two check boxes have been clicked.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Public Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lngClickRegion              As Long             '*' Region Selected (0=Vis, 1=Lock, 2=Anywhere Else)

    '*' Get the clicked region based upon the mouse position.
    '*'
    lngClickRegion = GetClickRegion(X)

    '*' Process click region.
    '*'
    Select Case lngClickRegion
    
        Case 0                                      '*' Visibility.
        
            '*' Toggle the visiblity state.
            '*'
            ToggleVisible
            
            '*' Raise the event.
            '*'
            RaiseEvent VisibilityToggle(m_bolVisibleChecked)
            Exit Sub
            
        Case 1                                      '*' Locked.
        
            '*' Toggle the locked state.
            '*'
            ToggleLocked
            
            '*' Raise the event.
            '*'
            RaiseEvent LockedToggle(m_bolLockedChecked)
            Exit Sub
                   
        Case 2                                      '*' Other
        
            '*' Just raise a click.
            '*'
            RaiseEvent Click
            
    End Select
    
    '*' Pass a mousedown, for good measure.
    '*'
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_MouseMove
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Event Broker.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_MouseUp
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Event Broker.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_ReadProperties
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Read any properties that have been defaulted or set by the user in the IDE.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    RaiseEvent ReadProperties(PropBag)
    m_BorderColor = PropBag.ReadProperty("BorderColor", DEFAULT_BORDERCOL)
    lblName.Caption = PropBag.ReadProperty("Caption", "LayerName")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    lblName.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 1)
    m_TrackingColor = PropBag.ReadProperty("TrackingColor", DEFAULT_TRACKCOL)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_Resize
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Reposition and resize components based upon the size of the control.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'

Private Sub UserControl_Resize()
    
'*' Fail through an any errors.
'*'
On Error Resume Next
    
    '*' Lock the display for cleaner updates.
    '*'
    Call LockWindowUpdate(UserControl.hwnd)

    '*' Move the three shape boxes to their proper places.
    '*'
    shpVisible.Move 0, 0, UserControl.ScaleHeight, UserControl.ScaleHeight
    shpLocked.Move UserControl.ScaleHeight - 1, 0, UserControl.ScaleHeight, UserControl.ScaleHeight
    shpName.Move ((UserControl.ScaleHeight - 1) * 2), 0, UserControl.ScaleWidth - ((UserControl.ScaleHeight - 1) * 2), _
                 UserControl.ScaleHeight
        

    '*' Adjust the label, so it fits into the shpName shape.
    '*'
    AdjustLabel
    
    '*' Reposition the visible and locked images.
    '*'
    imgVisible.Left = ((shpVisible.Width - imgVisible.Width) / 2) + shpVisible.Left
    imgVisible.Top = ((shpVisible.Height - imgVisible.Height) / 2) + shpVisible.Top
        
    imgLocked.Left = ((shpLocked.Width - imgLocked.Width) / 2) + shpLocked.Left
    imgLocked.Top = ((shpLocked.Height - imgLocked.Height) / 2) + shpLocked.Top
        
    '*' Determine whether to use the images or colors, based on the size of the control.
    '*'
    If shpVisible.Height > 16 Then
            
        imgVisible.Visible = (m_bolVisibleChecked = True)
        imgLocked.Visible = (m_bolLockedChecked = True)
        
        m_bolCheckImages = True
        
    Else
        
        imgVisible.Visible = False
        imgLocked.Visible = False
        
        m_bolCheckImages = False
    
    End If
    
    '*' Move the indicator bars.
    '*'
    picBottomDrop.Move 1, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 1, 2
    picTopDrop.Move 1, 0, UserControl.ScaleWidth - 1, 2
    
    '*' Allow Windows to redraw.
    '*'
    Call LockWindowUpdate(0)
    
    '*' Raise the event.
    '*'
    RaiseEvent Resize

End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_Show
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Event Broker.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_Show()
    RaiseEvent Show
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_WriteProperties
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Commit properties to the local property bag.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    RaiseEvent WriteProperties(PropBag)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, DEFAULT_BORDERCOL)
    Call PropBag.WriteProperty("Caption", lblName.Caption, "LayerName")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ForeColor", lblName.ForeColor, &H80000008)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 1)
    Call PropBag.WriteProperty("TrackingColor", m_TrackingColor, DEFAULT_TRACKCOL)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : imgLocked_Click
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Event Broker.  State Change.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub imgLocked_Click()

    '*' Toggle state and raise event.
    '*'
    ToggleLocked
    RaiseEvent LockedToggle(m_bolLockedChecked)
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : imgVisible_Click
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Event Broker and State Change.
'*'
'*' Input     : None
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub imgVisible_Click()

    '*' Toggle state and raise event.
    '*'
    ToggleVisible
    RaiseEvent VisibilityToggle(m_bolVisibleChecked)
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : lblName_MouseDown
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Event Broker and Input Redirector
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub lblName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '*' Pass to the usercontrol, so it can handle the MouseDown()
    '*'
    Call UserControl_MouseDown(Button, Shift, X, Y)
    
    '*' Raise the event.
    '*'
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : lblName_MouseUp
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Event Broker.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub lblName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : AdjustLabel
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   :  Adjust the placement of the label on a resize.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub AdjustLabel()

'*' Fail through on errors.
'*'
On Error Resume Next

    '*' Place the label.
    '*'
    lblName.Move shpName.Left + 5, _
                 (UserControl.ScaleHeight - lblName.Height) / 2, _
                 shpName.Width - 10
                 
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : GetClickRegion
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Determine where a point is on the control.
'*'
'*' Input     : X               - Vertical Coordinate.
'*'
'*' Output    : GetClickRegion  - 0=Vis, 1=Lock, 2= Everywhere else.
'*'
'**********************************************************************************************************************'
Private Function GetClickRegion(X As Single) As Long

    '*' Just compare the input to the known boundaries.
    '*'
    If X > 0 And X < shpVisible.Width Then
        GetClickRegion = 0
    ElseIf X > shpVisible.Width And X < shpName.Left Then
        GetClickRegion = 1
    Else
        GetClickRegion = 2
    End If
    
End Function

'**********************************************************************************************************************'
'*'
'*' Procedure : Refresh
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Refresh the display by redrawing it.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
    Call UserControl_Resize
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : ToggleVisible
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Boolean Toggle
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub ToggleVisible()
    Me.VisibleChecked = Not (m_bolVisibleChecked)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : ToggleLocked
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Boolean Toggle
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub ToggleLocked()
    Me.LockedChecked = Not (m_bolLockedChecked)
End Sub

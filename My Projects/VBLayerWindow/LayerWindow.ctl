VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl LayerWindow 
   BackColor       =   &H80000005&
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3570
   KeyPreview      =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   99
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   238
   Begin VB.PictureBox picContainer 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   0
      ScaleHeight     =   99
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   238
      TabIndex        =   0
      Top             =   0
      Width           =   3570
      Begin VB.PictureBox picContainerItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1440
         Left            =   240
         ScaleHeight     =   96
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   209
         TabIndex        =   1
         Top             =   30
         Width           =   3135
         Begin VBLayerWindow.ctlLayerItem ctlLayerItem 
            Height          =   330
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   582
         End
      End
      Begin MSComCtl2.FlatScrollBar fsbMain 
         Height          =   1500
         Left            =   3360
         TabIndex        =   2
         Top             =   -15
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   2646
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         LargeChange     =   20
         Orientation     =   8323072
         SmallChange     =   5
      End
   End
End
Attribute VB_Name = "LayerWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************************'
'*'
'*' Module    : LayerWindow
'*'
'*'
'*' Author    : Joseph M. Ferris <josephmferris@cox.net>
'*'
'*' Date      : 10.17.2002
'*'
'*' Depends   : Visual Basic 6.0 or Higher, SP5
'*'
'*' Purpose   : Provide a Photoshop-esque layer control for those wishing to have a way to manage layers in their
'*'             applications.
'*'
'*' Notes     : 1.  IMPORTANT!  For those who want more functionality than is provided with the interface of the
'*'                 control, a snapshot of the internal data structure can be retrieved as a property value, modified,
'*'                 and placed back live.  The reason for doing this is that this control will server many different
'*'                 needs for many different applications, and to allow the most use of the data by the developer is
'*'                 a desirable feature in this situation.
'*'
'*' Issues    : None
'*'
'*' ToDo      : 1.  Sync layer item top to top of container.  (Minor - Display Quality)
'*'
'**********************************************************************************************************************'

Option Explicit

'**********************************************************************************************************************'
'*'
'*' API Declarations (Private) - KERNEL32
'*'
'*' 1.  Sleep
'*'
'**********************************************************************************************************************'

Private Declare Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long)

'**********************************************************************************************************************'
'*'
'*' API Declarations (Private) - USER32
'*'
'*' 1.  ClipCursor
'*' 2.  GetCursorPos
'*' 3.  GetKeyState
'*' 4.  GetWindowRect
'*' 5.  LockWindowUpdate
'*' 6.  ReleaseCapture
'*' 7.  ScreenToClient
'*' 8.  SendMessage
'*'
'**********************************************************************************************************************'

Private Declare Function ClipCursor Lib "user32" ( _
              lpRect As Any) As Long

Private Declare Function GetCursorPos Lib "user32" ( _
              lpPoint As POINTAPI) As Long

Private Declare Function GetKeyState Lib "user32" ( _
        ByVal nVirtKey As Long) As Integer

Private Declare Function GetWindowRect Lib "user32" ( _
        ByVal hwnd As Long, _
              pRect As RECT) As Long

Private Declare Function LockWindowUpdate Lib "user32" ( _
        ByVal hwndLock As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" ( _
        ) As Long

Private Declare Function ScreenToClient Lib "user32" ( _
        ByVal hwnd As Long, _
              lpPoint As Any) As Long

Private Declare Function SendMessage Lib "user32" _
        Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
              lParam As Any) As Long

Private Declare Function SetCapture Lib "user32" ( _
        ByVal hwnd As Long) As Long

Private Declare Function SetCursorPos Lib "user32" ( _
        ByVal X As Long, _
        ByVal Y As Long) As Long

Private Declare Sub SetWindowPos Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long)

'**********************************************************************************************************************'
'*'
'*' User Defined Types
'*'
'*' 1. LAYER_CTL_LOCATION   : Reference of where ctlLayerItems exist when they are created
'*' 2. LAYER_DRAG_STATES    : Last known mouse state prior to clipping the cursor
'*' 3. LAYER_ITEM           : Information Layer Type
'*' 4. POINTAPI             : Required by GetCursorPos()
'*' 5. RECT                 : Required by ClipCursor(), GetWindowRect()
'*'
'**********************************************************************************************************************'

Private Type LAYER_CTL_LOCATION
    Left                        As Long
    Top                         As Long
End Type

Private Type LAYER_DRAG_STATES
    Index                       As Integer
    Button                      As Integer
    Shift                       As Integer
    X                           As Single
    Y                           As Single
End Type

Public Type LAYER_ITEM                              '*' Exported Type for Snapshot
    LayerName                   As String
    Visible                     As Boolean
    Locked                      As Boolean
    Selected                    As Boolean
End Type

Private Type POINTAPI
    X                           As Long
    Y                           As Long
End Type

Private Type RECT
    Left                        As Long
    Top                         As Long
    Right                       As Long
    Bottom                      As Long
End Type

'**********************************************************************************************************************'
'*'
'*' Private Constant Declarations
'*'
'*' 1.  HTCAPTION           : Treat as a captioned window.  Used by SendMessage()
'*' 2.  WM_NCLBUTTONDOWN    : Windows Message for Button Down.  Used by SendMessage()
'*' 3-7 Assorted            : Position arguments for SetWindowPos()
'*'
'**********************************************************************************************************************'

Private Const WM_NCLBUTTONDOWN  As Long = &HA1
Private Const HTCAPTION         As Long = 2
Private Const HWND_TOPMOST      As Long = -1
Private Const HWND_NOTOPMOST    As Long = -2
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_SHOWWINDOW    As Long = &H40
Private Const SWP_NOZORDER      As Long = &H4

'**********************************************************************************************************************'
'*'
'*' Private Member Variables
'*'
'**********************************************************************************************************************'

Private m_bolMoving             As Boolean              '*' Flag.  Indicates whether a drag operation is occurring.
Private m_lngLeftClip           As Long                 '*' Clipping Position to use on drag operations.
Private m_litLayerItems()       As LAYER_ITEM
Private m_lclControlPos()       As LAYER_CTL_LOCATION
Private m_ldsDragState          As LAYER_DRAG_STATES
Private m_olcTrackingColor      As OLE_COLOR            '*' Color for trackers on ctlLayerItem.
Private m_XOff                  As Long                 '*' Position offset while dragging.
Private m_YOff                  As Long                 '*' Position offset while dragging.

'**********************************************************************************************************************'
'*'
'*' Event Declarations
'*'
'**********************************************************************************************************************'

Event ChangedLocked(LayerName As String, Value As Boolean)
Event ChangedVisibility(LayerName As String, Value As Boolean)
Event Click()
Event DblClick()
Event InitProperties()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Resize()
Event Reordered()

'**********************************************************************************************************************'
'*'
'*' Error Constant Declarations
'*'
'**********************************************************************************************************************'

Const ERR_MOVE_FAILED_DESC      As String = "Moving the Layer has failed."
Const ERR_DUPLICATE_DESC        As String = "Duplicate Layer name."
Const ERR_LAYER_NOT_FOUND_DESC  As String = "Layer not Found."

Const ERR_MOVE_FAILED_NUM       As Long = vbObjectError + 1301
Const ERR_DUPLICATE_NUM         As Long = vbObjectError + 1302
Const ERR_LAYER_NOT_FOUND_NUM   As Long = vbObjectError + 1303

'**********************************************************************************************************************'
'*'
'*' Property Declarations
'*'
'*' 1.  ActiveControl (Get)
'*' 2.  BorderStyle (Get, Let)
'*' 3.  Enabled (Get, Let)
'*' 4.  hWnd (Get)
'*' 5.  MousePointer (Get, Let)
'*' 6.  SnapShot (Get, Let)
'*'
'**********************************************************************************************************************'
Public Property Get ActiveControl() As Object
Attribute ActiveControl.VB_Description = "Returns the control that has focus."
    Set ActiveControl = UserControl.ActiveControl
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Snapshot() As LAYER_ITEM()
    Snapshot = m_litLayerItems
End Property

Public Property Let Snapshot(New_Snapshot() As LAYER_ITEM)
    m_litLayerItems = New_Snapshot
    BuildContainer
End Property

Public Property Get TrackingColor() As OLE_COLOR
    TrackingColor = m_olcTrackingColor
End Property
Public Property Let TrackingColor(New_Tracking As OLE_COLOR)
    m_olcTracking = New_Tracking
    PropertyChanged "TrackingColor"
End Property

'**********************************************************************************************************************'
'*'
'*' Procedure : AddLayer
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Add a new layer to the LayerWindow
'*'
'*' Input     : LayerName   - Name of the layer.
'*'             Visible     - Whether the control is visible.
'*'             Locked      - Whether the control is locked.
'*'
'*' Output    : None
'*'
'**********************************************************************************************************************'
Public Sub AddLayer(LayerName As String, Visible As Boolean, Locked As Boolean)

'*' Handle Errors Locally.
'*'
On Error GoTo LocalHandler

Dim lngCounter                  As Long             '*' Iterative Counter
Dim lngIndex                    As Long             '*' Index of the Item

    '*' See if there is a layer by that name.
    '*'
    lngIndex = GetIndexByName(LayerName)
    
    '*' Raise an error if it is a dupe.
    '*'
    If Not (lngIndex = -1) Then
    
            '*' Raise an error for a duplicate layer.
            '*'
            Err.Raise ERR_DUPLICATE_NUM, "AddLayer()", ERR_DUPLICATE_DESC
            
    End If

    '*' Get the current max index for the item.
    '*'
    lngIndex = UBound(m_litLayerItems)

    '*' Check to see if there are any items in the array.
    '*'
    If lngIndex = 0 Then

        '*' Check to see if it needs to be offset to accomodate the zero based array.
        '*'
        If Not (m_litLayerItems(0).LayerName = vbNullString) Then
            lngIndex = 1
        End If

    Else
        
        '*' Increment to the new index.
        '*'
        lngIndex = lngIndex + 1
    End If

    '*' Resize the array.
    '*'
    ReDim Preserve m_litLayerItems(lngIndex)

    '*' Add the information.
    '*'
    m_litLayerItems(lngIndex).LayerName = LayerName
    m_litLayerItems(lngIndex).Visible = Visible
    m_litLayerItems(lngIndex).Locked = Locked

    '*' Rebuild the display.
    '*'
    BuildContainer

Exit Sub

LocalHandler:

    '*' Raise the error to the calling parent.
    '*'
    Err.Raise Err.Number, "AddLayer()", Err.Description
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : GetLockedState
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Determine if an item is locked, by its name.
'*'
'*' Input     : LayerName       - Name of layer to check.
'*'
'*' Output    : GetLockedState  - Flag.  Locked?
'*'
'**********************************************************************************************************************'
Public Function GetLockedState(LayerName As String) As Boolean

'*' Handle Errors Locally.
'*'
On Error GoTo LocalHandler

Dim lngCounter                  As Long             '*' Iterative loop counter.
Dim lngIndex                    As Long             '*' Index of Item

    '*' Find the layer.
    '*'
    lngIndex = GetIndexByName(LayerName)
    
    '*' Check to see if it was found.
    '*'
    If lngIndex = -1 Then
        Err.Raise ERR_LAYER_NOT_FOUND_NUM, "GetLockedState()", ERR_LAYER_NOT_FOUND_DESC
    End If
    
    '*' Return the state.
    '*'
    GetLockedState = m_litLayerItems(lngIndex).Locked
    
Exit Function

LocalHandler:

    '*' Raise to calling function.
    '*'
    Err.Raise Err.Number, "GetLockedState()", Err.Description
    
End Function

'**********************************************************************************************************************'
'*'
'*' Procedure : GetVisibilityState
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Determine if an item is locked, by its name.
'*'
'*' Input     : LayerName           - Name of layer to check.
'*'
'*' Output    : GetVisibilityState  - Flag.  Visible?
'*'
'**********************************************************************************************************************'
Public Function GetVisibilityState(LayerName As String) As Boolean

'*' Handle Errors Locally.
'*'
On Error GoTo LocalHandler

Dim lngCounter                  As Long             '*' Iterative loop counter.
Dim lngIndex                    As Long             '*' Index of Item

    '*' Find the layer.
    '*'
    lngIndex = GetIndexByName(LayerName)
    
    '*' Check to see if it was found.
    '*'
    If lngIndex = -1 Then
        Err.Raise ERR_LAYER_NOT_FOUND_NUM, "GetVisibilityState()", ERR_LAYER_NOT_FOUND_DESC
    End If
    
    '*' Return the state.
    '*'
    GetVisibilityState = m_litLayerItems(lngIndex).Visible
    
Exit Function

LocalHandler:

    '*' Raise to calling function.
    '*'
    Err.Raise Err.Number, "GetVisibilityState()", Err.Description
    
End Function

'**********************************************************************************************************************'
'*'
'*' Procedure : MoveLayer
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Move a layer to the position of an existing layer.
'*'
'*' Input     : CurrentLayer    - Name of the layer to move.
'*'             Placement       - Name of the layer to place at.
'*'
'*' Output    : None
'*'
'**********************************************************************************************************************'
Public Sub MoveLayer(CurrentLayer As String, Placement As String, PlaceAfter As Boolean)

'*' Handle Errors Locally.
'*'
On Error GoTo LocalHandler

Dim lngCounter                  As Long             '*' Iterative loop counter.
Dim lngCurrentIndex             As Long             '*' Current Index of item.
Dim lngNewIndex                 As Long             '*' New Index of Item.

    '*' Find the layers that were passed in.
    '*'
    lngCurrentIndex = GetIndexByName(CurrentLayer)
    lngNewIndex = GetIndexByName(Placement) + 1
    
    '*' Check to see if the items were found.
    '*'
    If lngCurrentIndex = -1 Then
        Err.Raise ERR_LAYER_NOT_FOUND_NUM, "MoveLayer()", ERR_LAYER_NOT_FOUND_DESC & ": " & CurrentLayer
    End If
    If lngNewIndex = -1 Then
        Err.Raise ERR_LAYER_NOT_FOUND_NUM, "MoveLayer()", ERR_LAYER_NOT_FOUND_DESC & ": " & Placement
    End If
       
    '*' Adjust if it is going to be placed after or before that item.
    '*'
    If PlaceAfter = True And lngNewIndex < UBound(m_litLayerItems) Then
        lngNewIndex = lngNewIndex + 1
    ElseIf PlaceAfter = False And lngNewIndex > 0 Then
        lngNewIndex = lngNewIndex - 1
    End If
    
    '*' Pop the layer into place.
    '*'
    Call PopLayer(lngCurrentIndex, lngNewIndex)
    
    '*' Rebuild display.
    '*'
    BuildContainer
    
    RaiseEvent Reordered
    
Exit Sub

LocalHandler:

    '*' Raise to calling.
    '*'
    Err.Raise Err.Number, "MoveLayer()", Err.Description
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : RemoveLayer
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Remove a layer from the LayerWindow
'*'
'*' Input     : LayerName       - Name of the layer to remove, or if no layer is named, the currently selected layer.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Public Sub RemoveLayer(Optional LayerName As String = vbNullString)

'*' Handle Errors Locally.
'*'
On Error GoTo LocalHandler

Dim bolItemFound                As Long             '*' Flag.  Whether item was found.
Dim lngCounter                  As Long             '*' Iteratvie Loop Counter.
Dim lngIndex                    As Long             '*' Item index.

    '*' Check to see if a layer exists, or see if one is currently selected.
    '*'
    If LayerName = vbNullString Then
        LayerName = Me.SelectedLayer
    End If
    
    '*' Check to see if a layer was found to be selected, or dump out.
    '*'
    If LayerName = vbNullString Then
        Err.Raise ERR_LAYER_NOT_FOUND_NUM, "RemoveLayer()", ERR_LAYER_NOT_FOUND_DESC
    End If
    
    '*' Find the layer.
    '*'
    lngIndex = GetIndexByName(LayerName)
    
    '*' See if the layer was found.
    '*'
    If lngIndex = -1 Then
        Err.Raise ERR_LAYER_NOT_FOUND_NUM, "RemoveLayer()", ERR_LAYER_NOT_FOUND_DESC
    End If
    
    '*' Physically set the name to nothing.  This will allow for it to be compressed out.
    '*'
    m_litLayerItems(lngIndex).LayerName = vbNullString
    
    '*' Compress the data UDT.
    '*'
    CompressLayer
    
    '*' Rebuild the display.
    '*'
    BuildContainer
    
    '*' Check to see if the scrollbar is enabled.
    '*'
    If fsbMain.Enabled = True Then
    
        '*' Check if removing an item changed the position.
        '*'
        If -picContainerItem.Top > (picContainerItem.ScaleHeight - picContainer.Height) Then
            picContainerItem.Top = -picContainerItem.ScaleHeight + picContainer.Height
            fsbMain.Value = fsbMain.Max
        End If
        
    End If
    
Exit Sub

LocalHandler:

    '*' Raise to the calling parent.
    '*'
    Err.Raise Err.Number, "RemoveLayer()", Err.Description
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : RenameLayer
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Rename a layer to a newly named layer.
'*'
'*' Input     : LayerName       - Original Name
'*'             NewLayerName    - New Name
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Public Sub RenameLayer(LayerName As String, NewLayerName As String)

'*' Handle Errors Locally.
'*'
On Error GoTo LocalHandler

Dim lngCounter                  As Long
Dim lngIndex                    As Long
Dim lngIndexDupe                As Long
    
    '*' Find the layers.
    '*'
    lngIndex = GetIndexByName(LayerName)
    lngIndexDupe = GetIndexByName(NewLayerName)
    
    '*' Don't let the user change something that isn't there.
    '*'
    If UBound(m_litLayerItems) = 0 And m_litLayerItems(0).LayerName = vbNullString Then
        Err.Raise ERR_LAYER_NOT_FOUND_NUM, "RenameLayer()", ERR_LAYER_NOT_FOUND_DESC
    End If
    
    '*' Make sure that the layer was found.
    '*'
    If lngIndex = -1 Then
        Err.Raise ERR_LAYER_NOT_FOUND_NUM, "RenameLayer()", ERR_LAYER_NOT_FOUND_DESC
    End If
    
    '*' Make sure that the new name is not a dupe.
    '*'
    If Not (lngIndexDupe = -1) Then
        Err.Raise ERR_DUPLICATE_NUM, "RenameLayer()", ERR_DUPLICATE_DESC
    End If
    
    '*' Assign the new name.
    '*'
    m_litLayerItems(lngIndex).LayerName = NewLayerName
    
    '*' Redraw the control.
    '*'
    BuildContainer
    
Exit Sub

LocalHandler:

    '*' Raise to the calling parent.
    '*'
    Err.Raise Err.Number, "RenameLayer()", Err.Description
        
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : SelectedLayer
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Get the name of the currently selected layer.
'*'
'*' Input     : None.
'*'
'*' Output    : SelectedLayer   - Name of the layer that is currently selected.
'*'
'**********************************************************************************************************************'
Public Function SelectedLayer() As String

'*' Fail through on error.
'*'
On Error Resume Next

Dim lngCounter                  As Long             '*' Iterative Loop Counter.

    '*' Iterate through all layer data.
    '*'
    For lngCounter = 0 To UBound(m_litLayerItems)
    
        '*' Check to see if the item is selected.  If found, return the name and dump out.
        '*'
        If m_litLayerItems(lngCounter).Selected = True Then
            SelectedLayer = ctlLayerItem(lngCounter).Caption
            Exit Function
        End If
        
    Next lngCounter
    
    '*' Nothing is apparently selected.  Aptly, return nothing.
    '*'
    SelectedLayer = vbNullString
    
End Function

'**********************************************************************************************************************'
'*'
'*' Procedure : ctlLayerItem_Click
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Provides Click() Event for Control.  Selects individually clicked items.
'*'
'*' Input     : Index   - Index of the ctlLayerItem that was clicked.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub ctlLayerItem_Click(Index As Integer)

'*' Fail through on error.
'*'
On Error Resume Next

Dim lngCounter                  As Long             '*' Iterative Loop Counter
Dim pntMousePos                 As POINTAPI         '*' Position holder during event.

    '*' Iterate through all of the layer items (from a data aspect) that are known to exist.
    '*'
    For lngCounter = 0 To UBound(m_litLayerItems)

        '*' Select the item, if it is the current one.  Deselect it if it isn't.
        '*'
        ctlLayerItem(lngCounter).Selected = (lngCounter = Index)
        m_litLayerItems(lngCounter).Selected = (lngCounter = Index)
        
    Next lngCounter

    If m_bolMoving Then
    
        '*' Get mouse position before raising the event.
        '*'
        Call GetCursorPos(pntMousePos)
        
        '*' Disable the trap during the event.
        '*'
        DisableTrap
        
    End If
    
    '*' Raise the event after the selection is made.
    '*'
    RaiseEvent Click
    
    If m_bolMoving Then
    
        '*' Set the mouse back to where it was.
        '*'
        Call SetCursorPos(pntMousePos.X, pntMousePos.Y)
        
        '*' Enable the trap again.
        '*'
        EnableTrap

    End If
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : ctlLayerItem_LockedToggle
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Raised Event from ctlLayerItem.  Dictates the change in state of the Locked Check Item.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub ctlLayerItem_LockedToggle(Index As Integer, Value As Boolean)

    '*' Reflect change locally.
    '*'
    m_litLayerItems(Index).Locked = Value
    
    '*' Raise the event.
    '*'
    RaiseEvent ChangedLocked(m_litLayerItems(Index).LayerName, Value)
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : ctlLayerItem_MouseDown
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Begin the dragging operation of layers.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub ctlLayerItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        
'*' Fail through on error.
'*'
On Error Resume Next

Dim lngCounter                  As Long             '*' Iterative Loop Counter
Dim pntPointStruct              As POINTAPI         '*' POINTAPI Structure for GetCursorPos()
Dim pntClip                     As POINTAPI
Dim rctWindow                   As RECT

    '*' Iterate through all of the layer items (from a data aspect) that are known to exist.
    '*'
    For lngCounter = 0 To UBound(m_litLayerItems)

        '*' Select the item, if it is the current one.  Deselect it if it isn't.
        '*'
        ctlLayerItem(lngCounter).Selected = (lngCounter = Index)
        m_litLayerItems(lngCounter).Selected = (lngCounter = Index)
        
    Next lngCounter

    '*' Flag the fact that the layer is going to begin moving.
    '*'
    m_bolMoving = True
        
    '*' Get the cursor position.
    '*'
    Call GetCursorPos(pntClip)
        
    '*' Set the point of clipping.
    '*'
    m_lngLeftClip = pntClip.X
    
    '*' Store the offset from the current mouse position.
    '*'
    m_XOff = X
    m_YOff = Y
    
    '*' Release capture from the OS.
    '*'
    ReleaseCapture
    
    '*' Set mouse capture to the ctlLayerItem.
    '*'
    SetCapture (ctlLayerItem(Index).hwnd)

    '*' Turn on the clipping.
    '*'
    EnableTrap
        
End Sub

Private Sub ctlLayerItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'*' Fail through on error.
'*'
On Error Resume Next

Dim lngCounter                  As Long             '*' Iterative loop counter.
Dim lngDropIndex                As Long             '*' Index that the user is over.
Dim pntPointStruct              As POINTAPI         '*' POINTAPI Structure for GetCursorPos
Dim rctWindowBounds             As RECT             '*' RECT Structure for GetWindowRect

Static LastX                    As Long             '*' Track last known.
Static LastY                    As Long             '*' Track last known.

                     
    '*' Check to see if the user is in a drag state.
    '*'
    If m_bolMoving = True Then
        
        '*' Get the current position of the cursor.
        '*'
        Call GetCursorPos(pntPointStruct)
        
        '*' Assign current values and leave if the static is blank.
        '*'
        If LastY = 0 And LastX = 0 Then
            LastX = pntPointStruct.X
            LastY = pntPointStruct.Y
            Exit Sub
        End If
        
        '*' Check to see if there are statics and they are the same as last time.  Leave if they are.
        '*'
        If LastX = pntPointStruct.X And LastY = pntPointStruct.Y Then
            Exit Sub
        End If
            
        '*' Get the bounding region from the displayed picturebox that bounds the interal one.
        '*'
        Call GetWindowRect(picContainer.hwnd, rctWindowBounds)
        
        '*' Physically set the position, minding the offset also.
        '*'
        SetWindowPos ctlLayerItem(Index).hwnd, _
                     HWND_TOPMOST, _
                     pntPointStruct.X - rctWindowBounds.Left - m_XOff, _
                     pntPointStruct.Y - rctWindowBounds.Top - m_YOff, _
                     0, _
                     0, _
                     SWP_NOSIZE
                                     
        '*' Check to see if the user is in the five pixel hotspot at the top of the control.
        '*'
        If pntPointStruct.Y - 5 < rctWindowBounds.Top Then
        
            '*' Attempt to move the scrollbar.
            '*'
            If fsbMain.Value >= fsbMain.SmallChange Then
            
                '*' Implement a small change up.
                '*'
                fsbMain.Value = fsbMain.Value - fsbMain.SmallChange
            
            Else
            
                '*' Set to the minimum.
                '*'
                fsbMain.Value = 0
                
            End If
            
        '*' Check to see if the user is in the five pixel hotspot at the bottom of the control.
        '*'
        ElseIf pntPointStruct.Y + 5 > rctWindowBounds.Bottom Then
                    
            '*' Attempt to move the scrollbar.
            '*'
            If fsbMain.Value <= fsbMain.Max - fsbMain.SmallChange Then
            
                '*' Implement a small change down.
                '*'
                fsbMain.Value = fsbMain.Value + fsbMain.SmallChange
            
            Else
            
                '*' Set to the max.
                '*'
                fsbMain.Value = fsbMain.Max
                                
            End If
            
        End If
        
        '*' Get the item that the user is hovering over.
        '*'
        lngDropIndex = GetDropIndex(CLng(Index))
                                                                                
        '*' Iterate through all possible tracking positions.
        '*'
        For lngCounter = -1 To UBound(m_litLayerItems)
                                
            '*' Make sure that it is a workable value.
            '*'
            If lngDropIndex > -2 Then
            
                '*' Look for a match to turn on, otherwise, make sure it is off.
                '*'
                If lngCounter = lngDropIndex And lngCounter < UBound(m_litLayerItems) Then
                    ctlLayerItem(lngCounter + 1).TopIndicator = True
                Else
                    ctlLayerItem(lngCounter + 1).TopIndicator = False
                End If
                
            End If
            
            '*' Make sure that the value is within bounds.
            '*'
            If lngDropIndex <= UBound(m_litLayerItems) Then
                            
                '*' Turn on all matches, watch for the bottom item, and turn everything else off.
                '*'
                If lngCounter = UBound(m_litLayerItems) And lngCounter = lngDropIndex Then
                    ctlLayerItem(lngCounter).BottomIndicator = True
                ElseIf lngCounter = lngDropIndex And lngCounter < UBound(m_litLayerItems) Then
                    ctlLayerItem(lngCounter).BottomIndicator = True
                Else
                    ctlLayerItem(lngCounter).BottomIndicator = False
                    
                End If
                
            End If
                            
        Next lngCounter
                  
    End If
        
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : ctlLayerItem_MouseUp
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Provide the release aspect of the dragging operation.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub ctlLayerItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'*' Fail through on error.
'*'
On Error Resume Next

Dim lngCounter                  As Long
Dim lngNewIndex                 As Long             '*' Index of drop destination.

    For lngCounter = 0 To UBound(m_litLayerItems)
        ctlLayerItem(lngCounter).TopIndicator = False
        ctlLayerItem(lngCounter).BottomIndicator = False
    Next lngCounter
    
    '*' Flag that there is not further dragging operation occurring.
    '*'
    m_bolMoving = False

    Call ClipCursor(ByVal 0&)
           
    '*' Hide the layer so that it doesn't look out of place.
    '*'
    ctlLayerItem(Index).Visible = False
                
    '*' Sync the VB .Top and .Left property of the control with its new position.
    '*'
    Call ForceUpdatePos(ctlLayerItem(Index))
    
    '*' Determine what the index of the new item will be, based upon where it was dropped.
    '*'
    lngNewIndex = GetDropIndex(CLng(Index))
    
    '*' Make sure that the new index and the item index are both valid indices.
    '*'
    If lngNewIndex > -2 And Index > -1 Then
            
        '*' Item moving from down to up, but not to front.
        '*'
        If (lngNewIndex < Index) And lngNewIndex > 0 Then
        
            '*' Offset by one for theoretical offset.
            '*'
            Call PopLayer(CLng(Index), lngNewIndex + 1)
            
        '*' Moving from down to up to the front.
        '*'
        ElseIf (lngNewIndex < Index) And lngNewIndex = 0 Then
                
            '*' Swap layer indices.
            '*'
            Call PopLayer(CLng(Index), lngNewIndex)
            
        ElseIf lngNewIndex = -1 Then
        
            Call PopLayer(CLng(Index), lngNewIndex + 1)
            
        Else
        
            '*' Swap layer indices.
            '*'
            Call PopLayer(CLng(Index), lngNewIndex)
        
        End If
            
    End If
                
    '*' Rebuild the display of the control.
    '*'
    BuildContainer
        
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : ctlLayerItem_VisibilityToggle
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Raised Event from ctlLayerItem.  Provides feedback to the toggle of the Visible Check Item.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub ctlLayerItem_VisibilityToggle(Index As Integer, Value As Boolean)

    '*' Toggle the Data value for the toggle.
    '*'
    m_litLayerItems(Index).Visible = Value
    
    '*' Raise the event.
    '*'
    RaiseEvent ChangedVisibility(m_litLayerItems(Index).LayerName, Value)
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : fsbMain_Change
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Move the Container Item picturebox to give the "illusion" of scrolling.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub fsbMain_Change()

    '*' Offset the internal picturebox to reflect the position of the scrollbar.
    '*'
    picContainerItem.Move 0, -(fsbMain.Value), picContainerItem.ScaleWidth, picContainerItem.ScaleHeight
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picContainer_Resize
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Event raised for Resize from picContainer.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picContainer_Resize()

'*' Fail through on error.
'*'
On Error Resume Next

    '*' Move the scrollbar to the right edge of the container.
    '*'
    fsbMain.Move picContainer.ScaleWidth - fsbMain.Width, 0, fsbMain.Width, picContainer.ScaleHeight
    
    '*' Check to see if the scrollbar is enabled.
    '*'
    If fsbMain.Enabled Then
    
        '*' Resize the container child.
        '*'
        picContainerItem.Move 0, 0, picContainer.Width - fsbMain.Width + 1, picContainerItem.Height
        
    Else
    
        picContainerItem.Move 0, 0, picContainer.Width - fsbMain.Width + 1, picContainer.Height
        
    End If
    
    '*' Resize all of the layer items.
    '*'
    ResizeLayerItems
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_Click
'*'
'*'
'*' Date      : 10.17.2002
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
'*' Date      : 10.17.2002
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
'*' Procedure : UserControl_Initialize
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Array and property initialization (for scrollbar).
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_Initialize()

'*' Fail through on error.
'*'
On Error Resume Next

    '*' Initialize the array for the data of the ctlLayerItems and the data that they represent.
    '*'
    ReDim m_litLayerItems(0)
    ReDim m_lclControlPos(0)
        
    '*' Set the small change of the scrollbar to be the height of a single ctlLayerItem.
    '*'
    fsbMain.SmallChange = ctlLayerItem(0).Height - 1
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_InitProperties
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Event Broker.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_InitProperties()
    RaiseEvent InitProperties
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_KeyDown
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Event Broker.  Watch keyboard for scrolling events.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
'*' Fail through on error.
'*'
On Error Resume Next

Dim bolMoveSelUp                As Boolean          '*' Flag.  Going Up?
Dim bolMoveSelDown              As Boolean          '*' Flag.  Going Down?
Dim lngCounter                  As Long             '*' Iterative Loop Counter
Dim lngIndex                    As Long             '*' Item Index

    '*' Check for up or down arrow press.
    '*'
    Select Case KeyCode
        Case 38                                     '*' Up Arrow
            bolMoveSelUp = True
        Case 40                                     '*' Down Arrow
            bolMoveSelDown = True
    End Select
    
    '*' Check if one of the two keys were pressed.
    '*'
    If Not (bolMoveSelDown = bolMoveSelUp) Then
    
        '*' Initialize the index.
        '*'
        lngIndex = -1
        
        '*' Iterate through the data items.
        '*'
        For lngCounter = 0 To UBound(m_litLayerItems)
            If m_litLayerItems(lngCounter).Selected Then
                lngIndex = lngCounter
                Exit For
            End If
        Next lngCounter
        
        '*' Check to see if an item was found.
        '*'
        If lngIndex = -1 Then
            GoTo EventHandler
        End If
        
        '*' Check if the user pressed down.
        '*'
        If bolMoveSelDown Then
        
            '*' Move the selection down, if it is not at the end.
            '*'
            If lngIndex < UBound(m_litLayerItems) Then
                m_litLayerItems(lngIndex).Selected = False
                ctlLayerItem(lngIndex).Selected = False
                m_litLayerItems(lngIndex + 1).Selected = True
                ctlLayerItem(lngIndex + 1).Selected = True
            End If
                        
            '*' Check for scrollbar adjustments.
            '*'
            If ctlLayerItem(lngIndex).Top + ctlLayerItem(lngIndex).Height >= _
               picContainer.Height + (Abs(picContainerItem.Top)) Then
            
                If fsbMain.Value + fsbMain.SmallChange < fsbMain.Max Then
                    fsbMain.Value = fsbMain.Value + fsbMain.SmallChange
                Else
                    fsbMain.Value = fsbMain.Max
                End If
                
            End If
            
        Else
        
            '*' Move the selection up, if it is not at the front.
            '*'
            If lngIndex > 0 Then
                m_litLayerItems(lngIndex).Selected = False
                ctlLayerItem(lngIndex).Selected = False
                m_litLayerItems(lngIndex - 1).Selected = True
                ctlLayerItem(lngIndex - 1).Selected = True
            End If
            
            '*' Check for scrollbar adjustments.
            '*'
            If ctlLayerItem(lngIndex).Top < Abs(picContainerItem.Top) Then
            
                If fsbMain.Value - fsbMain.SmallChange > 0 Then
                    fsbMain.Value = fsbMain.Value - fsbMain.SmallChange
                Else
                    fsbMain.Value = 0
                End If
                
            End If
            
        End If
        
        '*' Rebuild the display.
        '*'
        BuildContainer
                
    End If
    
EventHandler:
    RaiseEvent KeyDown(KeyCode, Shift)
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_KeyPress
'*'
'*'
'*' Date      : 10.17.2002
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
'*' Date      : 10.17.2002
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
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Event Broker.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_MouseMove
'*'
'*'
'*' Date      : 10.17.2002
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
'*' Date      : 10.17.2002
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
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Read the properties from the prop bag to persist values from IDE.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_olcTrackingColor = PropBag.ReadProperty("TrackingColor", 0)
    
    '*' Draw it for the first time.
    '*'
    BuildContainer
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_Resize
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Resize the outter container of the control.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_Resize()
     
'*' Fail through on any errors.
'*'
On Error Resume Next

    '*' Resize the container and change the largechange height.
    '*'
    picContainer.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    fsbMain.LargeChange = picContainer.ScaleHeight
    
    '*' Raise the Resize() event.
    '*'
    RaiseEvent Resize
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_WriteProperties
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Write properties to prop bag to persist property values.
'*'
'*' Input     : Predefined.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("TrackingColor", m_olcTrackingColor, 0)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : BuildContainer
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Rebuild the display of the control.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'

Private Sub BuildContainer()

'*' Fail through on errors.
'*'
On Error Resume Next

Dim lngCounter                  As Long             '*' Iterative Loop Counter
Dim lngIndex                    As Long             '*' Item Index
Dim lngLayerControlCount        As Long             '*' Number of ctlLayer Controls.
Dim lngLayerItemCount           As Long             '*' Number of data items.

    '*' Lock windows for fast and clean graphic update.
    '*'
    Call LockWindowUpdate(UserControl.hwnd)

    '*' Get the number of data items.
    '*'
    lngLayerItemCount = UBound(m_litLayerItems)
    
    '*' Check to see if there is any.
    '*'
    If lngLayerItemCount = 0 Then
    
        '*' Check to see if there is an item at 0.
        '*'
        If m_litLayerItems(0).LayerName = vbNullString Then
            lngLayerItemCount = -1
        End If
        
    End If
        
    '*' Get the target number of controls to use.
    '*'
    lngLayerControlCount = ctlLayerItem.Count - 1
    
    '*' Check to see if any controls need to be created.
    '*'
    If lngLayerItemCount > lngLayerControlCount Then
    
        '*' Loop until there is the correct amount of controls.
        '*'
        Do Until lngLayerControlCount = lngLayerItemCount
        
            '*' Assign new index and load it.
            '*'
            lngIndex = ctlLayerItem.UBound + 1
            Load ctlLayerItem(lngIndex)
            
            '*' Assign default properties.
            '*'
            ctlLayerItem(lngIndex).Visible = True
            ctlLayerItem(lngIndex).Top = (ctlLayerItem(0).Height - 1) * lngIndex
            lngLayerControlCount = ctlLayerItem.Count - 1
            ctlLayerItem(lngIndex).Left = 0
            
            '*' Store the position it has.
            '*'
            ReDim Preserve m_lclControlPos(lngIndex)
            m_lclControlPos(lngIndex).Top = ctlLayerItem(lngIndex).Top
                                    
        Loop
        
    End If
    
    '*' Match the height to the default.
    '*'
    picContainerItem.Height = (((ctlLayerItem(0).Height - 1) * (UBound(m_litLayerItems) + 1))) + 1
    
    '*' Iterate through all the controls.
    '*'
    For lngCounter = 0 To ctlLayerItem.UBound
        
        '*' Set the tracking color.
        '*'
        ctlLayerItem(lngCounter).TrackingColor = m_olcTrackingColor
        
        '*' Check if it should be visible.
        '*'
        If lngCounter > lngLayerItemCount Then
        
            '*' Turn it off.
            '*'
            ctlLayerItem(lngCounter).Visible = False
            
        Else
            
            '*' Position it correctly.
            '*'
            If ctlLayerItem(lngCounter).Left <> m_lclControlPos(lngCounter).Left Then
                ctlLayerItem(lngCounter).Left = m_lclControlPos(lngCounter).Left
            End If
            
            If ctlLayerItem(lngCounter).Top <> m_lclControlPos(lngCounter).Top Then
                ctlLayerItem(lngCounter).Top = m_lclControlPos(lngCounter).Top
            End If
            
            '*' Assign properties from the data item.
            '*'
            ctlLayerItem(lngCounter).Caption = m_litLayerItems(lngCounter).LayerName
            ctlLayerItem(lngCounter).VisibleChecked = m_litLayerItems(lngCounter).Visible
            ctlLayerItem(lngCounter).LockedChecked = m_litLayerItems(lngCounter).Locked
            ctlLayerItem(lngCounter).Selected = m_litLayerItems(lngCounter).Selected

            '*' Make it visible.
            '*'
            ctlLayerItem(lngCounter).Visible = True
                        
        End If
        
    Next lngCounter
    
    '*' Make any scrollbar adjustments, as needed.
    '*'
    If picContainerItem.ScaleHeight > picContainer.ScaleHeight Then
        fsbMain.Max = picContainerItem.ScaleHeight - picContainer.ScaleHeight + 1
        fsbMain.Enabled = True
    Else
        fsbMain.Enabled = False
        picContainerItem.Top = 0
    End If
    
    '*' Remove the lock on the control.
    '*'
    Call LockWindowUpdate(0)
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : CompressLayer
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Remove any items that have a vbNullString as its caption.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub CompressLayer()

'*' Fail through on error.
'*'
On Error Resume Next

Dim litLocal()                  As LAYER_ITEM       '*' Temp Data Store
Dim lngCounter                  As Long             '*' Iterative Loop Counter
Dim lngIndex                    As Long             '*' Item Index

    '*' Initialize array.
    '*'
    ReDim litLocal(0)
    
    '*' Default Index.
    '*'
    lngIndex = 0
    
    '*' Iterate through the data.
    '*'
    For lngCounter = 0 To UBound(m_litLayerItems)
        
        '*' Check to see if it is valid, then add it to the temp data store if it is.
        '*'
        If Not (m_litLayerItems(lngCounter).LayerName = vbNullString) Then
                
            ReDim Preserve litLocal(lngIndex)
            litLocal(lngIndex).Locked = m_litLayerItems(lngCounter).Locked
            litLocal(lngIndex).Visible = m_litLayerItems(lngCounter).Visible
            litLocal(lngIndex).LayerName = m_litLayerItems(lngCounter).LayerName
            litLocal(lngIndex).Selected = m_litLayerItems(lngCounter).Selected
            
            lngIndex = lngIndex + 1
            
        End If
        
    Next lngCounter
    
    '*' Assign the temp to the member.
    '*'
    m_litLayerItems = litLocal
        
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : DisableTrap
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Set clipping area for mouse to entire screen area.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub DisableTrap()

Dim lngResult                   As Long
Dim rctClipping                 As RECT

    '*' Set the clipping area to be the full screen.
    '*'
    With rctClipping
      .Left = 0&
      .Top = 0&
      .Right = Screen.Width / Screen.TwipsPerPixelX
      .Bottom = Screen.Height / Screen.TwipsPerPixelY
    End With
    
    '*' Clip it.
    '*'
    lngResult& = ClipCursor(rctClipping)

End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : EnableTrap
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Set clipping area for drag operations.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub EnableTrap()

'*' Fail through on error.
'*'
On Error Resume Next

Dim lngResult                   As Long             '*' Function Return
Dim rctClipping                 As RECT             '*' RECT for ClipCursor
Dim rctResult                   As RECT             '*' RECT for GetWindowRect()
    
    '*' Get the rectangle of the clipping area.
    '*'
    Call GetWindowRect(picContainer.hwnd, rctResult)
    
    '*' Use the result RECT and the current mouse position (m_lngLeftClip) to set the rectangle to clip.
    '*'
    With rctClipping
      .Left = m_lngLeftClip
      .Top = rctResult.Top
      .Right = m_lngLeftClip + 1
      .Bottom = rctResult.Bottom
    End With
    
    '*' Clip it.
    '*'
    lngResult& = ClipCursor(rctClipping)
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : ForceUpdatePos
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Provides a fix for a known VB Behavioral issue.  When use ReleaseCapture() and SendMessage() to move
'*'             a control, VB does not update .Left and .Top properties of the control.
'*'
'*' Input     : ctlUnknown  - Any object of having a Type of Control
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub ForceUpdatePos(ctlUnknown As Control)

'*' Fail through on error.
'*'
On Error Resume Next

Dim rctControlBounds            As RECT             '*' Control Boundaries
    
    '*' Obtain the control's position in relation to its container.
    '*'
    Call GetWindowRect(ctlUnknown.hwnd, rctControlBounds)
    Call ScreenToClient(ctlUnknown.Container.hwnd, rctControlBounds.Left)
    
    '*' Call .Move() to physically "move" the control to its current location.  Fools VB into updating the property.
    '*'
    ctlUnknown.Move rctControlBounds.Left, rctControlBounds.Top
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : GetDropIndex
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Get the index of a completed drag operation's target.
'*'
'*' Input     : Index           - Index of control that is being dropped.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Function GetDropIndex(Index As Long) As Long

'*' Fail through on errors.
'*'
On Error Resume Next

Dim lngCounter                  As Long             '*' Iterative Loop Counter
Dim lngX                        As Long             '*' X Pos for Mouse
Dim lngY                        As Long             '*' Y Pos for Mouse
Dim pntCursorPos                As POINTAPI         '*' POINTAPI for GetCursorPos
Dim rctContainer                As RECT             '*' RECT for GetWindowRect

    '*' Get the mouse position and window rectangles.
    '*'
    Call GetWindowRect(picContainer.hwnd, rctContainer)
    Call GetCursorPos(pntCursorPos)
    
    '*' Check that the cursor is inside the horizontal constraints.
    '*'
    If rctContainer.Right > pntCursorPos.X And rctContainer.Left < pntCursorPos.X Then
    
        '*' Check that the cursor is inside the vertical constraints.
        '*'
        If rctContainer.Top <= pntCursorPos.Y And rctContainer.Bottom >= pntCursorPos.Y Then
            
            '*' Get the X, Y coords of the cursor, relative to the control.
            '*'
            lngX = rctContainer.Right - pntCursorPos.X
            lngY = Abs(picContainerItem.Top) + (pntCursorPos.Y - rctContainer.Top)
            
            '*' Iterate through the layer items.
            '*'
            For lngCounter = 0 To UBound(m_litLayerItems)
                
                '*' Check any index that isn't the dragged one.
                '*'
                If lngCounter <> Index Then
                    If ctlLayerItem(lngCounter).Top <= lngY And _
                       ctlLayerItem(lngCounter).Top + ctlLayerItem(lngCounter).Height >= lngY Then
                                                
                        '*' Determine if the user wants to drop before or after this object.
                        '*'
                        If ctlLayerItem(lngCounter).Top <= lngY And _
                           ctlLayerItem(lngCounter).Top + (0.5 * ctlLayerItem(lngCounter).Height) >= lngY Then
                        
                            '*' Top half drop.  Adjust it if it is above terminal (0).
                            '*'
                            GetDropIndex = lngCounter - 1
                            
                        Else
                                                                                
                            '*' Bottom half drop.
                            '*'
                            GetDropIndex = lngCounter
                    
                        End If
                                                
                        '*' Bail out.
                        '*'
                        Exit Function
                        
                    End If
                    
                End If
                
            Next lngCounter
            
       End If
        
    End If
    
    '*' Return a failing index.
    '*'
    GetDropIndex = -2
    
End Function

'**********************************************************************************************************************'
'*'
'*' Procedure : IsButtonDown
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Retrieve state of mouse button.
'*'
'*' Input     : Button          - Button to check for being pressed
'*'
'*' Output    : IsButtonDown    - Flag. Button down?
'*'
'**********************************************************************************************************************'
Private Function IsButtonDown(ByVal Button As MouseButtonConstants) As Boolean
    IsButtonDown = CBool(GetKeyState(Button) And &H8000)
End Function

'**********************************************************************************************************************'
'*'
'*' Procedure : PopLayer
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Pop a layer into a given index.
'*'
'*' Input     : CurrentIndex    - Index of current layer.
'*'             NewIndex        - The Index it should have when it is done.
'*'
'*' Output    : None
'*'
'**********************************************************************************************************************'
Private Function PopLayer(CurrentIndex As Long, NewIndex As Long)

On Error GoTo LocalHandler

Dim litTemp()                   As LAYER_ITEM       '*' Temp Data Store
Dim lngCounter                  As Long             '*' Iterative Counter
Dim lngIndex                    As Long             '*' Item Index
Dim lngIndexSort()              As Long             '*' Sorted Item Index
    
    '*' Initialize the Index
    '*'
    lngIndex = 0
    
    '*' Resize the Sorted Item Index to its final size.  (Minus one for the current element).
    '*'
    ReDim lngIndexSort(UBound(m_litLayerItems) - 1)
    
    '*' Iterate through the data.
    '*'
    For lngCounter = 0 To UBound(m_litLayerItems)
        
        '*' Store it if it isn't the current index.
        '*'
        If Not (CurrentIndex = lngCounter) Then
        
            lngIndexSort(lngIndex) = lngCounter
            lngIndex = lngIndex + 1
                    
        End If
        
    Next lngCounter
    
    '*' Resize the data store.
    '*'
    ReDim litTemp(UBound(m_litLayerItems))
    
    '*' Reset the index.
    '*'
    lngIndex = 0
    
    '*' Iterate through the data.
    '*'
    For lngCounter = 0 To UBound(m_litLayerItems)
    
        '*' Check for moved items.
        '*'
        If Not (lngCounter = NewIndex) Then
            
            '*' Write data to temp data store.
            '*'
            With litTemp(lngCounter)
                .LayerName = m_litLayerItems(lngIndexSort(lngIndex)).LayerName
                .Locked = m_litLayerItems(lngIndexSort(lngIndex)).Locked
                .Visible = m_litLayerItems(lngIndexSort(lngIndex)).Visible
                .Selected = m_litLayerItems(lngIndexSort(lngIndex)).Selected
            End With
        
            '*' Increment index.
            '*'
            lngIndex = lngIndex + 1
            
        Else
        
            '*' Write source index to its new home.
            '*'
            With litTemp(lngCounter)
                .LayerName = m_litLayerItems(CurrentIndex).LayerName
                .Locked = m_litLayerItems(CurrentIndex).Locked
                .Visible = m_litLayerItems(CurrentIndex).Visible
                .Selected = m_litLayerItems(CurrentIndex).Selected
            End With
        
        End If
            
    Next lngCounter
    
    '*' Overwrite member data store.
    '*'
    m_litLayerItems = litTemp
    
Exit Function

LocalHandler:

    '*' Raise the error to the calling function.
    '*'
    Err.Raise ERR_MOVE_FAILED_NUM, "PopLayer()", ERR_MOVE_FAILED_DESC & ": " & Err.Description
    
End Function

'**********************************************************************************************************************'
'*'
'*' Procedure : ResizeLayerItems
'*'
'*'
'*' Date      : 10.17.2002
'*'
'*' Purpose   : Match the width of the layer items to the containing control.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub ResizeLayerItems()

'*' Fail through on any errors.
'*'
On Error Resume Next

Dim lngCounter                  As Long             '*' Iterative Loop Counter.

    '*' Iterate through the cltLayerItems that are present.
    '*'
    For lngCounter = ctlLayerItem.lBound To ctlLayerItem.UBound
    
        '*' Match the width to that of the container.
        '*'
        ctlLayerItem(lngCounter).Width = picContainerItem.ScaleWidth
        
    Next lngCounter

End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : GetIndexByName
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Retrieve the index of an item, based on the name that is passed.
'*'
'*' Input     : strName         - Name of the layer to look for.
'*'
'*' Output    : GetIndexByName  - Long Index of Item, if exists.  -1 on no match.
'*'
'**********************************************************************************************************************'
Private Function GetIndexByName(strName As String) As Long

On Error Resume Next

Dim lngCounter                  As Long
Dim lngIndex                    As Long

    '*' Initialize the index.
    '*'
    lngIndex = -1

    '*' Iterate through all the data items.
    '*'
    For lngCounter = 0 To UBound(m_litLayerItems)
        
        '*' Check for a match.  If it matches, return exit the loop.
        '*'
        If m_litLayerItems(lngCounter).LayerName = strName Then
            lngIndex = lngCounter
            Exit For
        End If
    Next lngCounter
    
    '*' Return current index.
    '*'
    GetIndexByName = lngIndex
    
End Function

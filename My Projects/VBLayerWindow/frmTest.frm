VERSION 5.00
Object = "*\AVBLayers.vbp"
Begin VB.Form frmTest 
   Caption         =   "VBLayers Test Form"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VBLayerWindow.LayerWindow LayerWindow1 
      Height          =   2655
      Left            =   45
      TabIndex        =   7
      Top             =   60
      Width           =   4335
      _extentx        =   7646
      _extenty        =   4683
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Item"
      Height          =   375
      Left            =   4575
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Log"
      Height          =   375
      Left            =   4590
      TabIndex        =   6
      Top             =   2340
      Width           =   1575
   End
   Begin VB.TextBox txtEvents 
      Height          =   1740
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   2805
      Width           =   6150
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "Rename Selected"
      Height          =   390
      Left            =   4590
      TabIndex        =   3
      Top             =   1395
      Width           =   1560
   End
   Begin VB.CommandButton cmdMoveSelected 
      Caption         =   "Move Selected"
      Height          =   405
      Left            =   4590
      TabIndex        =   2
      Top             =   1860
      Width           =   1575
   End
   Begin VB.CommandButton cmdRemoveSelected 
      Caption         =   "Remove Selected"
      Height          =   405
      Left            =   4590
      TabIndex        =   1
      Top             =   915
      Width           =   1560
   End
   Begin VB.CommandButton cmdPopulate 
      Caption         =   "Populate"
      Height          =   375
      Left            =   4575
      TabIndex        =   0
      Top             =   45
      Width           =   1575
   End
   Begin VB.Shape shpGenericBorder 
      BorderColor     =   &H80000000&
      Height          =   2685
      Left            =   30
      Top             =   45
      Width           =   4365
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************************************'
'*'
'*' Module    : frmTest
'*'
'*'
'*' Author    : Joseph M. Ferris <josephmferris@cox.net>
'*'
'*' Date      : 10.18.2002
'*'
'*' Depends   : Visual Basic 6.  Service Pack 5, or higher is recommended.
'*'             mscomct2.ocx
'*'             vblayers.ocx
'*'
'*' Purpose   : Provides a test environment for the VBLayers ActiveX Control
'*'
'*' Notes     : This sample form provides examples of some, but not all of the features supported by this version of
'*'             VBLayers.  Please refer to the included documentation, vblayers.xml, for a breakdown of the complete
'*'             interface for this control.
'*'
'*'             For more access to the data inside of the control, try declaring a variable array of type
'*'             LAYER_ITEM and setting its value to the snapshot property of the control.  Changes will be reflected
'*'             immediately when reassigning the LAYER_ITEM array back to the snapshot property.
'*'
'**********************************************************************************************************************'

Option Explicit

'**********************************************************************************************************************'
'*'
'*' Procedure : cmdAdd_Click
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Provide a means to add a layer in a test environment.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub cmdAdd_Click()

'*' Handle Errors Locally.
'*'
On Error GoTo LocalHandler

Dim strNewName                  As String           '*' User response.

    '*' Get user response for new layer name.
    '*'
    strNewName = InputBox("Name for new Layer:")
    
    '*' Attempt to add the new layer.
    '*'
    Call LayerWindow1.AddLayer(strNewName, True, True)
    
Exit Sub

LocalHandler:

    '*' Log the error.
    '*'
    AppendLog "Error: (" & Err.Number & ") " & Err.Description
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : cmdClear_Click
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Clear event viewer text box.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub cmdClear_Click()

'*' Fail through on error.
'*'
On Error Resume Next

    '*' Clear the text box.
    '*'
    txtEvents.Text = vbNullString
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : cmdPopulate_Click
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Add a quick set of test layers in the test environment.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub cmdPopulate_Click()

'*' Handle Errors Locally.
'*'
On Error GoTo LocalHandler

    '*' Try to add some generic items.
    '*'
    Call LayerWindow1.AddLayer("A", True, True)
    Call LayerWindow1.AddLayer("B", True, True)
    Call LayerWindow1.AddLayer("C", True, True)
    Call LayerWindow1.AddLayer("D", True, True)
    Call LayerWindow1.AddLayer("E", True, True)
    Call LayerWindow1.AddLayer("F", True, True)
    Call LayerWindow1.AddLayer("G", True, True)
    Call LayerWindow1.AddLayer("H", True, True)
    Call LayerWindow1.AddLayer("I", True, True)
    Call LayerWindow1.AddLayer("J", True, True)
    Call LayerWindow1.AddLayer("K", True, True)
    Call LayerWindow1.AddLayer("L", True, True)
    Call LayerWindow1.AddLayer("M", True, True)
    Call LayerWindow1.AddLayer("N", True, True)
    Call LayerWindow1.AddLayer("O", True, True)
    Call LayerWindow1.AddLayer("P", True, True)
    Call LayerWindow1.AddLayer("Q", True, True)
    Call LayerWindow1.AddLayer("R", True, True)
    Call LayerWindow1.AddLayer("S", True, True)
    Call LayerWindow1.AddLayer("T", True, True)
    Call LayerWindow1.AddLayer("U", True, True)
    Call LayerWindow1.AddLayer("V", True, True)
    Call LayerWindow1.AddLayer("W", True, True)
    Call LayerWindow1.AddLayer("X", True, True)
    Call LayerWindow1.AddLayer("Y", True, True)
    Call LayerWindow1.AddLayer("Z", True, True)
    
Exit Sub

LocalHandler:

    '*' Log the current error.
    '*'
    AppendLog "Error: " & Err.Description
    
    '*' Move on to the next item.
    '*'
    Resume Next
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : cmdRemoveSelected_Click
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Remove the currently selected layer.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub cmdRemoveSelected_Click()

'*' Handle Errors Locally.
'*'
On Error GoTo LocalHandler

    '*' Remove the currently selected layer.
    '*'
    Call LayerWindow1.RemoveLayer
    
Exit Sub

LocalHandler:

    '*' Log the error.
    '*'
    AppendLog "Error: (" & Err.Number & ") " & Err.Description

End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : cmdMoveSelected_Click
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Move the currently selected item to a new position.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub cmdMoveSelected_Click()

Dim strMoveBefore               As String           '*' User response.

    '*' Get name of the layer to place before.
    '*'
    strMoveBefore = InputBox("Place Layer '" & LayerWindow1.SelectedLayer & "' before item named:")
    
    '*' Move the layer.
    '*'
    Call LayerWindow1.MoveLayer(LayerWindow1.SelectedLayer, strMoveBefore, False)
    
Exit Sub

LocalHandler:

    '*' Log the error.
    '*'
    AppendLog "Error: (" & Err.Number & ") " & Err.Description
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : cmdRename_Click
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Allow the user to rename the currently selected layer.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'

Private Sub cmdRename_Click()

'*' Handle Error Locally.
'*'
On Error GoTo LocalHandler

Dim strNewName                  As String           '*' User Response

    '*' Prompt for the new name.
    '*'
    strNewName = InputBox("New Name for Layer '" & LayerWindow1.SelectedLayer & "':")
    
    '*' Attempt to rename it.
    '*'
    Call LayerWindow1.RenameLayer(LayerWindow1.SelectedLayer, strNewName)
    
Exit Sub

LocalHandler:

    '*' Log the error.
    '*'
    AppendLog "Error: (" & Err.Number & ") " & Err.Description
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : AppendLog
'*'
'*'
'*' Date      : 10.18.2002
'*'
'*' Purpose   : Append a string to the end of the log and move the log to the end of display.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub AppendLog(Message As String)
    
'*' Fail through on errors.
'*'
On Error Resume Next

    '*' Log to the text box.
    '*'
    txtEvents.Text = txtEvents.Text & Message & vbCrLf
    
    '*' Move to the end of the text box.
    '*'
    txtEvents.SelStart = Len(txtEvents.Text)
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Message Logging.
'*'
'**********************************************************************************************************************'

Private Sub LayerWindow1_ChangedLocked(LayerName As String, Value As Boolean)
    AppendLog "ChangedLocked(" & LayerName & ", " & Value & ")"
End Sub

Private Sub LayerWindow1_ChangedVisibility(LayerName As String, Value As Boolean)
    AppendLog "ChangedVisibility(" & LayerName & ", " & Value & ")"
End Sub

Private Sub LayerWindow1_Click()
    AppendLog "Click()"
End Sub

Private Sub LayerWindow1_DblClick()
    AppendLog "DblClick()"
End Sub

Private Sub LayerWindow1_GotFocus()
    AppendLog "GotFocus()"
End Sub

Private Sub LayerWindow1_KeyDown(KeyCode As Integer, Shift As Integer)
    AppendLog "KeyDown(" & KeyCode & ", " & Shift & ")"
End Sub

Private Sub LayerWindow1_KeyPress(KeyAscii As Integer)
    AppendLog "KeyPress(" & KeyAscii & ")"
End Sub

Private Sub LayerWindow1_KeyUp(KeyCode As Integer, Shift As Integer)
    AppendLog "KeyUp(" & KeyCode & ", " & Shift & ")"
End Sub

Private Sub LayerWindow1_LostFocus()
    AppendLog "LostFocus()"
End Sub

Private Sub LayerWindow1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    AppendLog "MouseDown(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"
End Sub

Private Sub LayerWindow1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    AppendLog "Mousemove(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"
End Sub

Private Sub LayerWindow1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    AppendLog "MouseUp(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"
End Sub

Private Sub LayerWindow1_Reordered()
    AppendLog "Reordered()"
End Sub

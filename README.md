<div align="center">

## Form Resizer Deluxe


</div>

### Description

Class module for resizing/repositioning controls on a form. See source code for details.
 
### More Info
 
see source code


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Buzzurro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-buzzurro.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-buzzurro-form-resizer-deluxe__1-1107/archive/master.zip)





### Source Code

```
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE DESCRIPTION:
'  Class for scaling/repositioning controls on a form
'
' DATE CREATED:
'  10-22-1998
'
' AUTHOR:
'  John Buzzurro
'
' COPYRIGHT NOTICE:
'  Copyright (c) 1998 by John Buzzurro
'
' NOTES:
' A) To give your form resizing ability:
'
'  1) Create an instance of this class
'  2) Set the SourceForm property of this class = your form
'  3) In your Form_Resize() event handler, call the ScaleControls() method of
'   this class
'  4) Optional - To refine the type of scaling/positioning of a control:
'   Set the .Tag property of the control to a string containing an "@" sign
'   followed by any of the following, separated by commas: T,L,H,W,
'   Where  T = Adjust control's Top position
'        L = Adjust control's Left position
'        H = Adjust control's height
'        W = Adjust control's width
'
'   Example: "@T,L"
'   Note that if the .Tag property does not start with a "@", the resizer
'   assumes "@T,L,H,W"; If the .Tag property is set only to "@", the
'   resizer will not attempt to reposition or resize the control.
'
' B) If you Add or Remove controls at runtime, OR you adjust the height or
'  width of the form programmatically at runtime, you MUST call the
'  ReInitialize() method of this class.
'
' C) For Image controls, you need to set the Stretch property to True for the
'  control to properly resize.
'
' EXAMPLE FORM MODULE CODE:
'  Option Explicit
'
'  Dim mcFormResize As New clsFormResize
'
'  Private Sub Form_Load()
'    mcFormResize.SourceForm = Me
'  End Sub
'
'  Private Sub Form_Resize()
'    mcFormResize.ScaleControls
'  End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
' Information we store about a control
Private Type tControlPosition
  cControl As Control   ' Reference to the control instance
  nLeft As Long      ' Original Left pos
  nTop As Long      ' Original Top pos
  nWidth As Long     ' Original Width
  nHeight As Long     ' Original Height
End Type
' Module-scope storage
Private mfSourceForm As Form        ' The form we are resizing
Private mnLastWidth As Long         ' Original form width
Private mnLastHeight As Long        ' Original form height
Private matControlPos() As tControlPosition ' Array for storing control info
Private mbIsFirstTime As Boolean      ' Flag indicating first time scale
'*****************************************************************************
' Property: SourceForm (get)
'      Returns the form object to which this CFormMetric instance belongs
'*****************************************************************************
Public Property Get SourceForm() As Form
  Set SourceForm = mfSourceForm
End Property
'*****************************************************************************
' Property: SourceForm (put)
'      Sets the form object to which this CFormMetric instance belongs
'*****************************************************************************
Public Property Let SourceForm(ByVal vNewValue As Form)
  Set mfSourceForm = vNewValue
End Property
'*****************************************************************************
' Method:  ScaleControls()
'      Adjusts the size and position of the form's controls relative to
'      the current form size
'*****************************************************************************
Public Sub ScaleControls()
  Dim sFlags As String, _
    sTemp As String
  Dim nDeltaLeft As Long, _
    nDeltaTop As Long, _
    nDeltaWidth As Long, _
    nDeltaHeight As Long, _
    nTextHeight As Long
  Dim iControl As Integer
  Dim nWidthChange As Double, _
    nHeightChange As Double
  Dim bIsLineControl As Boolean
  Dim cControl As Control
  If (mbIsFirstTime) Then
    Call SaveInitialState
    Exit Sub
  End If
  ' If the form is minimized, there's nothing to do
  If (mfSourceForm.WindowState = vbMinimized) Then Exit Sub
  ' Calculate the change in form size
  nDeltaWidth = mfSourceForm.ScaleWidth - mnLastWidth
  nDeltaHeight = mfSourceForm.ScaleHeight - mnLastHeight
  nHeightChange = mfSourceForm.ScaleHeight / mnLastHeight
  nWidthChange = mfSourceForm.ScaleWidth / mnLastWidth
  For iControl = LBound(matControlPos) To UBound(matControlPos)
    Set cControl = matControlPos(iControl).cControl
    With cControl
      ' Test whether this is a line control; If it is,
      ' we need to set its X1, X2, Y1, Y2 properties instead of the
      ' usual .Top, .Left, .Height, .Width properties
      If (TypeOf cControl Is VB.Line) Then
        bIsLineControl = True
      Else
        ' Not a line control
        bIsLineControl = False
      End If
      On Error GoTo errScaleControls
      ' See if the control has specified which attributes can be changed
      sFlags = UCase(.Tag)
      ' If none specified, assume all
      If (sFlags = "") Then sFlags = "@T,H,L,W"
      ' If Tag property is used for something else, assume all
      If (Left$(sFlags, 1) <> "@") Then sFlags = "@T,H,L,W"
      ' Resize/Reposition the control
      If (bIsLineControl) Then
        ' Line control
        If (InStr(sFlags, "T")) Then .Y1 = (matControlPos(iControl).nTop * nHeightChange)
        If (InStr(sFlags, "H")) Then .Y2 = (matControlPos(iControl).nHeight * nHeightChange)
        If (InStr(sFlags, "L")) Then .X1 = (matControlPos(iControl).nLeft * nWidthChange)
        If (InStr(sFlags, "W")) Then .X2 = (matControlPos(iControl).nWidth * nWidthChange)
      Else
        ' All other controls
        If (InStr(sFlags, "T")) Then .Top = (matControlPos(iControl).nTop * nHeightChange)
        If (InStr(sFlags, "H")) Then .Height = (matControlPos(iControl).nHeight * nHeightChange)
        If (InStr(sFlags, "L")) Then .Left = (matControlPos(iControl).nLeft * nWidthChange)
        If (InStr(sFlags, "W")) Then .Width = (matControlPos(iControl).nWidth * nWidthChange)
      End If
'      nTextHeight = 0
'      nTextHeight = mfSourceForm.TextHeight(.Caption)
'      If Not nTextHeight Then nTextHeight = mfSourceForm.TextHeight(.Text)
'      If (nTextHeight > .Height) Then
'        .Height = mfSourceForm.TextHeight(.Caption) * 1.2
'        .Height = mfSourceForm.TextHeight(.Text) * 1.2
'      End If
    End With
skipControl:
  Next iControl
Exit Sub
errScaleControls:
  ' If the Left, Top, Height or Width property is read-only, skip to next line;
  ' Otherwise, skip the control entirely
  If (Err.Number = 383 Or Err.Number = 387 Or Err.Number = 393 Or Err.Number = 438) Then Resume Next
  Resume skipControl
End Sub
'*****************************************************************************
' Method:  SizeToScreen()
'      Size the form relative to the current screen resolution
'
' Params:  Percentage of total screen size to use for the form size
'*****************************************************************************
Public Sub SizeFormToScreen(nPercent As Integer)
  Dim w As Long, _
    h As Long
  w = Int(Screen.Width * (nPercent / 100))
  h = Int(Screen.Height * (nPercent / 100))
  mfSourceForm.Width = w
  mfSourceForm.Height = h
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method:  ReInitialize()
'  ReInitialize Method; This method should be called if:
'  a) You programmatically change the form size at runtime;
'  b) You add or remove controls to/from the form at runtime
'
' MODIFIES:
'  Recreates the matControlPos() array and saves the current form
'  information
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReInitialize()
  Call SaveInitialState
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESCRIPTION:
'  Class instance initialization; Initialize module-scope variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
  mbIsFirstTime = True
  mnLastWidth = 0
  mnLastHeight = 0
  Set mfSourceForm = Nothing
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESCRIPTION:
'  Save the initial state of the form and controls attached to this class
'  instance
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveInitialState()
  Call SaveFormInfo
  Call SaveControlInfo
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESCRIPTION:
'  Save form width and height
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveFormInfo()
  ' Take a snapshot of the form's initial position and size
  With mfSourceForm
    If (TypeOf mfSourceForm Is MDIForm) Then
      mnLastWidth = .Width
      mnLastHeight = .Height
    Else
      mnLastWidth = .ScaleWidth
      mnLastHeight = .ScaleHeight
    End If
  End With
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESCRIPTION:
'  Save state information for each control on the form
'
' NOTES:
'  We only save info for controls that have a Visible property
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveControlInfo()
  Dim cControl As Control
  Dim bCanSetLeft As Boolean, _
    bCanSetTop As Boolean, _
    bCanSetWidth As Boolean, _
    bCanSetHeight As Boolean, _
    bHasVisibleProp As Boolean, _
    bHasCaptionProp As Boolean, _
    bHasTextProp As Boolean, _
    bTemp As Boolean
  Dim i As Integer
  Erase matControlPos
  ''
  ' Loop through each control on the form...
  For Each cControl In mfSourceForm.Controls
    bCanSetLeft = True
    bCanSetTop = True
    bCanSetWidth = True
    bCanSetHeight = True
    bHasVisibleProp = True
    bHasCaptionProp = True
    bHasTextProp = True
    With cControl
      ' Test whether control has a Visible property
      On Error GoTo errNoVisibleProp
      bTemp = .Visible
      On Error GoTo 0
      ' If control has visible property, save its info in an array
      If (bHasVisibleProp) Then
        i = i + 1
        ReDim Preserve matControlPos(1 To i)
        Set matControlPos(i).cControl = cControl
        ' If this is a Line control...
        If (TypeOf cControl Is VB.Line) Then
          ' ... then this is a special case 'cause its position
          '   is specified by different properties than normal
          matControlPos(i).nLeft = .X1
          matControlPos(i).nTop = .Y1
          matControlPos(i).nWidth = .X2
          matControlPos(i).nHeight = .Y2
        Else
          ' This is not a Line control
          On Error Resume Next
          matControlPos(i).nLeft = .Left
          matControlPos(i).nTop = .Top
          matControlPos(i).nWidth = .Width
          matControlPos(i).nHeight = .Height
          On Error GoTo 0
        End If
      End If
    End With
  Next cControl
  mbIsFirstTime = False
Exit Sub
errNoVisibleProp:
  bHasVisibleProp = False
  Resume Next
End Sub
```


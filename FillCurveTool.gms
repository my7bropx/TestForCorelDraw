' ==============================================================================
' CorelDRAW Fill Curve Tool with Selected Elements
' Version: 2.0
' Compatible with: CorelDRAW 2018-2026
' ==============================================================================
' Description:
'   Fills open or closed curves with selected elements while preserving:
'   - Original element size and angle
'   - Element colors and properties
'   - Curve original shape
'
' Features:
'   - User-friendly dialog interface
'   - Adjustable horizontal and vertical spacing
'   - Multiple fill patterns (Grid, Hexagonal)
'   - Works with both open and closed curves
'   - Preserves all element attributes
' ==============================================================================

Sub FillCurveWithElements()
    On Error GoTo ErrorHandler

    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange

    ' Validate selection
    If sr.Count < 2 Then
        MsgBox "Please select at least TWO objects:" & vbCrLf & _
               "1. One or more fill elements (objects to repeat)" & vbCrLf & _
               "2. One container curve (open or closed)", _
               vbExclamation, "Selection Required"
        Exit Sub
    End If

    ' Identify container and fill elements
    Dim container As Shape
    Dim fillElements As New Collection
    Dim i As Integer
    Dim largestArea As Double
    Dim tempArea As Double

    largestArea = 0

    ' Find the largest shape as container
    For i = 1 To sr.Count
        tempArea = sr(i).SizeWidth * sr(i).SizeHeight
        If tempArea > largestArea Then
            largestArea = tempArea
            Set container = sr(i)
        End If
    Next i

    ' Collect fill elements (all except container)
    For i = 1 To sr.Count
        If Not sr(i) Is container Then
            fillElements.Add sr(i)
        End If
    Next i

    If fillElements.Count = 0 Then
        MsgBox "Could not identify fill elements. Please ensure you have selected:" & vbCrLf & _
               "- At least one small element to repeat" & vbCrLf & _
               "- One larger container curve", _
               vbExclamation, "Invalid Selection"
        Exit Sub
    End If

    ' Convert container to curves if needed
    If container.Type <> cdrCurveShape Then
        container.ConvertToCurves
    End If

    ' Show options dialog
    Dim patternType As String
    Dim hSpacing As Double
    Dim vSpacing As Double
    Dim fillMode As String

    If Not ShowOptionsDialog(patternType, hSpacing, vSpacing, fillMode) Then
        Exit Sub ' User cancelled
    End If

    ' Perform the fill operation
    Call PerformFill(container, fillElements, patternType, hSpacing, vSpacing, fillMode)

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
End Sub


' ==============================================================================
' Show Options Dialog
' ==============================================================================
Function ShowOptionsDialog(ByRef patternType As String, _
                           ByRef hSpacing As Double, _
                           ByRef vSpacing As Double, _
                           ByRef fillMode As String) As Boolean

    On Error GoTo DialogError

    ' Create dialog
    Dim dlg As Object
    Set dlg = Application.CreateDialog

    ' Dialog dimensions
    dlg.SetSize 420, 280
    dlg.Title = "Fill Curve with Elements - Options"

    ' Pattern Type
    dlg.AddLabel 20, 20, 150, 20, "Fill Pattern:"
    Dim cboPattern As Object
    Set cboPattern = dlg.AddComboBox(180, 20, 200, 100, "", 0)
    cboPattern.AddItem "Grid Pattern"
    cboPattern.AddItem "Hexagonal Pattern"
    cboPattern.Text = "Grid Pattern"

    ' Horizontal Spacing
    dlg.AddLabel 20, 60, 150, 20, "Horizontal Spacing (%):"
    Dim txtHSpacing As Object
    Set txtHSpacing = dlg.AddTextBox(180, 60, 200, 20, "0")

    ' Vertical Spacing
    dlg.AddLabel 20, 95, 150, 20, "Vertical Spacing (%):"
    Dim txtVSpacing As Object
    Set txtVSpacing = dlg.AddTextBox(180, 95, 200, 20, "0")

    ' Fill Mode
    dlg.AddLabel 20, 135, 150, 20, "Fill Mode:"
    Dim cboFillMode As Object
    Set cboFillMode = dlg.AddComboBox(180, 135, 200, 80, "", 0)
    cboFillMode.AddItem "Fill Inside Only"
    cboFillMode.AddItem "Fill on Boundary"
    cboFillMode.AddItem "Fill Inside & Boundary"
    cboFillMode.Text = "Fill Inside Only"

    ' Help text
    dlg.AddLabel 20, 175, 380, 40, _
        "Spacing: 0% = elements touch, 100% = one element width/height apart" & vbCrLf & _
        "Negative values = overlap, Positive = gap"

    ' Buttons
    dlg.AddButton 140, 230, 120, 30, "OK", 1
    dlg.AddButton 270, 230, 120, 30, "Cancel", 2

    ' Show dialog
    If dlg.Show = 1 Then
        patternType = cboPattern.Text
        hSpacing = Val(txtHSpacing.Text) / 100
        vSpacing = Val(txtVSpacing.Text) / 100
        fillMode = cboFillMode.Text
        ShowOptionsDialog = True
    Else
        ShowOptionsDialog = False
    End If

    Exit Function

DialogError:
    ' Fallback to input boxes if dialog creation fails
    patternType = "Grid Pattern"

    Dim userInput As String
    userInput = InputBox("Enter horizontal spacing percentage (0 = touching, 100 = one width apart):", _
                        "Horizontal Spacing", "0")
    If userInput = "" Then
        ShowOptionsDialog = False
        Exit Function
    End If
    hSpacing = Val(userInput) / 100

    userInput = InputBox("Enter vertical spacing percentage (0 = touching, 100 = one height apart):", _
                        "Vertical Spacing", "0")
    If userInput = "" Then
        ShowOptionsDialog = False
        Exit Function
    End If
    vSpacing = Val(userInput) / 100

    fillMode = "Fill Inside Only"
    ShowOptionsDialog = True
End Function


' ==============================================================================
' Perform Fill Operation
' ==============================================================================
Sub PerformFill(container As Shape, _
                fillElements As Collection, _
                patternType As String, _
                hSpacing As Double, _
                vSpacing As Double, _
                fillMode As String)

    On Error Resume Next

    ' Get container boundaries
    Dim leftX As Double, rightX As Double
    Dim topY As Double, bottomY As Double

    leftX = container.LeftX
    rightX = container.RightX
    topY = container.TopY
    bottomY = container.BottomY

    ' Get average dimensions of fill elements
    Dim avgWidth As Double, avgHeight As Double
    Dim elem As Variant
    Dim totalW As Double, totalH As Double
    Dim elemCount As Integer

    totalW = 0: totalH = 0: elemCount = 0
    For Each elem In fillElements
        totalW = totalW + elem.SizeWidth
        totalH = totalH + elem.SizeHeight
        elemCount = elemCount + 1
    Next elem

    avgWidth = totalW / elemCount
    avgHeight = totalH / elemCount

    ' Calculate spacing
    Dim dx As Double, dy As Double

    If patternType = "Hexagonal Pattern" Then
        dx = avgWidth * (1 + hSpacing)
        dy = avgHeight * 0.866 * (1 + vSpacing) ' Hex vertical spacing
    Else ' Grid Pattern
        dx = avgWidth * (1 + hSpacing)
        dy = avgHeight * (1 + vSpacing)
    End If

    ' Start fill operation
    Application.Optimization = True
    ActiveDocument.BeginCommandGroup "Fill Curve with Elements"

    Dim x As Double, y As Double
    Dim cx As Double, cy As Double
    Dim row As Integer
    Dim count As Integer
    Dim elemIndex As Integer

    count = 0
    row = 0
    elemIndex = 1

    ' Fill loop
    y = bottomY + avgHeight / 2
    Do While y <= topY - avgHeight / 2
        x = leftX + avgWidth / 2

        ' Offset for hexagonal pattern
        If patternType = "Hexagonal Pattern" And row Mod 2 = 1 Then
            x = x + dx / 2
        End If

        Do While x <= rightX - avgWidth / 2
            cx = x
            cy = y

            ' Check if point should be filled based on fill mode
            Dim shouldFill As Boolean
            shouldFill = False

            If fillMode = "Fill Inside Only" Then
                If container.Curve.IsPointInside(cx, cy) Then
                    shouldFill = True
                End If
            ElseIf fillMode = "Fill on Boundary" Then
                ' Check if point is near the boundary
                If Not container.Curve.IsPointInside(cx, cy) Then
                    ' Check if any nearby point is inside
                    If container.Curve.IsPointInside(cx + avgWidth/4, cy) Or _
                       container.Curve.IsPointInside(cx - avgWidth/4, cy) Or _
                       container.Curve.IsPointInside(cx, cy + avgHeight/4) Or _
                       container.Curve.IsPointInside(cx, cy - avgHeight/4) Then
                        shouldFill = True
                    End If
                End If
            Else ' Fill Inside & Boundary
                ' More lenient check
                If container.Curve.IsPointInside(cx, cy) Or _
                   container.Curve.IsPointInside(cx + avgWidth/3, cy) Or _
                   container.Curve.IsPointInside(cx - avgWidth/3, cy) Or _
                   container.Curve.IsPointInside(cx, cy + avgHeight/3) Or _
                   container.Curve.IsPointInside(cx, cy - avgHeight/3) Then
                    shouldFill = True
                End If
            End If

            If shouldFill Then
                ' Cycle through fill elements
                Dim sourceElem As Shape
                Set sourceElem = fillElements(elemIndex)

                ' Duplicate and position
                Dim newShape As Shape
                Set newShape = sourceElem.Duplicate
                newShape.CenterX = cx
                newShape.CenterY = cy
                ' Preserve original rotation, size, and all properties

                count = count + 1

                ' Cycle to next element
                elemIndex = elemIndex + 1
                If elemIndex > fillElements.Count Then
                    elemIndex = 1
                End If
            End If

            x = x + dx
        Loop

        row = row + 1
        y = y + dy
    Loop

    ' Finish
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveWindow.Refresh

    MsgBox "Successfully placed " & count & " elements inside the curve!", _
           vbInformation, "Fill Complete"
End Sub

' ==============================================================================
' CorelDRAW Fill Curve Tool - Simple Version
' Version: 2.0 Simple
' Compatible with: CorelDRAW 2018-2026 (Maximum Compatibility)
' ==============================================================================
' This simplified version uses InputBox for settings to ensure compatibility
' across all CorelDRAW versions from 2018 to 2026
' ==============================================================================

Sub FillCurveSimple()
    On Error GoTo ErrorHandler

    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange

    ' Validate selection
    If sr.Count < 2 Then
        MsgBox "Selection Required:" & vbCrLf & vbCrLf & _
               "Please select at least 2 objects:" & vbCrLf & _
               "  - One or more small elements (to repeat)" & vbCrLf & _
               "  - One larger container curve", _
               vbExclamation + vbOKOnly, "Fill Curve Tool"
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

    ' Collect fill elements
    For i = 1 To sr.Count
        If Not sr(i) Is container Then
            fillElements.Add sr(i)
        End If
    Next i

    If fillElements.Count = 0 Then
        MsgBox "Error: Could not identify fill elements." & vbCrLf & vbCrLf & _
               "Make sure the container curve is larger than the fill elements.", _
               vbExclamation, "Invalid Selection"
        Exit Sub
    End If

    ' Convert container to curves if needed
    If container.Type <> cdrCurveShape Then
        container.ConvertToCurves
    End If

    ' Get user options
    Dim hSpacing As String
    Dim vSpacing As String
    Dim patternChoice As String

    ' Pattern selection
    patternChoice = InputBox( _
        "Select Fill Pattern:" & vbCrLf & vbCrLf & _
        "1 = Grid Pattern (regular rows and columns)" & vbCrLf & _
        "2 = Hexagonal Pattern (honeycomb style)" & vbCrLf & vbCrLf & _
        "Enter 1 or 2:", _
        "Pattern Selection", "1")

    If patternChoice = "" Then Exit Sub ' User cancelled

    ' Horizontal spacing
    hSpacing = InputBox( _
        "Horizontal Spacing:" & vbCrLf & vbCrLf & _
        "   0  = Elements touch each other" & vbCrLf & _
        " 50  = Half element width gap" & vbCrLf & _
        "100 = Full element width gap" & vbCrLf & _
        "-50 = Elements overlap 50%" & vbCrLf & vbCrLf & _
        "Enter percentage (-100 to 500):", _
        "Horizontal Spacing", "0")

    If hSpacing = "" Then Exit Sub ' User cancelled

    ' Vertical spacing
    vSpacing = InputBox( _
        "Vertical Spacing:" & vbCrLf & vbCrLf & _
        "   0  = Elements touch each other" & vbCrLf & _
        " 50  = Half element height gap" & vbCrLf & _
        "100 = Full element height gap" & vbCrLf & _
        "-50 = Elements overlap 50%" & vbCrLf & vbCrLf & _
        "Enter percentage (-100 to 500):", _
        "Vertical Spacing", "0")

    If vSpacing = "" Then Exit Sub ' User cancelled

    ' Convert to decimal
    Dim hSpace As Double, vSpace As Double
    hSpace = Val(hSpacing) / 100
    vSpace = Val(vSpacing) / 100

    ' Determine pattern type
    Dim isHexPattern As Boolean
    isHexPattern = (Val(patternChoice) = 2)

    ' Perform the fill
    Call DoFillOperation(container, fillElements, hSpace, vSpace, isHexPattern)

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Fill Curve Tool Error"
End Sub


' ==============================================================================
' Main Fill Operation
' ==============================================================================
Sub DoFillOperation(container As Shape, _
                    fillElements As Collection, _
                    hSpacing As Double, _
                    vSpacing As Double, _
                    isHexPattern As Boolean)

    On Error Resume Next

    ' Get container boundaries
    Dim leftX As Double, rightX As Double
    Dim topY As Double, bottomY As Double

    leftX = container.LeftX
    rightX = container.RightX
    topY = container.TopY
    bottomY = container.BottomY

    ' Calculate average element dimensions
    Dim elem As Variant
    Dim avgW As Double, avgH As Double
    Dim totalW As Double, totalH As Double
    Dim elemCount As Integer

    totalW = 0
    totalH = 0
    elemCount = 0

    For Each elem In fillElements
        totalW = totalW + elem.SizeWidth
        totalH = totalH + elem.SizeHeight
        elemCount = elemCount + 1
    Next elem

    avgW = totalW / elemCount
    avgH = totalH / elemCount

    ' Calculate step sizes
    Dim dx As Double, dy As Double

    If isHexPattern Then
        ' Hexagonal pattern
        dx = avgW * (1 + hSpacing)
        dy = avgH * 0.866025 * (1 + vSpacing) ' sqrt(3)/2 for hex
    Else
        ' Grid pattern
        dx = avgW * (1 + hSpacing)
        dy = avgH * (1 + vSpacing)
    End If

    ' Progress indicator
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

    ' Main fill loop - start from bottom
    y = bottomY + avgH / 2

    Do While y <= topY - avgH / 2
        x = leftX + avgW / 2

        ' Hexagonal offset for odd rows
        If isHexPattern And (row Mod 2 = 1) Then
            x = x + dx / 2
        End If

        Do While x <= rightX - avgW / 2
            cx = x
            cy = y

            ' Check if center point is inside the curve
            ' Works for both open and closed curves
            If container.Curve.IsPointInside(cx, cy) Then
                ' Get source element (cycle through if multiple)
                Dim sourceElem As Shape
                Set sourceElem = fillElements(elemIndex)

                ' Create duplicate
                Dim newShape As Shape
                Set newShape = sourceElem.Duplicate

                ' Position the duplicate
                newShape.CenterX = cx
                newShape.CenterY = cy

                ' The duplicate preserves:
                ' - Original size (SizeWidth, SizeHeight)
                ' - Original rotation angle
                ' - Original colors and fill
                ' - All other properties

                count = count + 1

                ' Move to next element
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

    ' Complete operation
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveWindow.Refresh

    ' Show result
    MsgBox "Fill Complete!" & vbCrLf & vbCrLf & _
           "Total elements placed: " & count, _
           vbInformation, "Success"
End Sub


' ==============================================================================
' Quick Fill with Default Settings (No prompts)
' ==============================================================================
Sub QuickFill()
    On Error GoTo ErrorHandler

    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange

    If sr.Count < 2 Then
        MsgBox "Please select at least 2 objects: fill element(s) and container curve", _
               vbExclamation, "Selection Required"
        Exit Sub
    End If

    ' Find container
    Dim container As Shape
    Dim fillElements As New Collection
    Dim i As Integer
    Dim largestArea As Double, tempArea As Double

    largestArea = 0
    For i = 1 To sr.Count
        tempArea = sr(i).SizeWidth * sr(i).SizeHeight
        If tempArea > largestArea Then
            largestArea = tempArea
            Set container = sr(i)
        End If
    Next i

    For i = 1 To sr.Count
        If Not sr(i) Is container Then
            fillElements.Add sr(i)
        End If
    Next i

    If container.Type <> cdrCurveShape Then
        container.ConvertToCurves
    End If

    ' Use default settings: Grid pattern, 0% spacing
    Call DoFillOperation(container, fillElements, 0, 0, False)

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Quick Fill Error"
End Sub

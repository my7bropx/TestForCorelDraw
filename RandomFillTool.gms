' ==============================================================================
' CorelDRAW Random Fill Tool
' Version: 3.0
' Compatible with: CorelDRAW 2018-2026
' ==============================================================================
' Description:
'   Fills open or closed curves with RANDOMLY placed and rotated elements
'   to achieve maximum coverage with density and spacing control
' ==============================================================================

Option Explicit

' Main entry point
Sub RandomFillCurve()
    On Error GoTo ErrorHandler

    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange

    ' Validate selection
    If sr.Count < 2 Then
        MsgBox "Please select at least 2 objects:" & vbCrLf & vbCrLf & _
               "1. One or more fill elements (to scatter randomly)" & vbCrLf & _
               "2. One container curve (open or closed)", _
               vbExclamation, "Random Fill Tool"
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
        MsgBox "Error: Could not identify fill elements." & vbCrLf & _
               "Make sure the container is larger than fill elements.", _
               vbExclamation, "Invalid Selection"
        Exit Sub
    End If

    ' Convert container to curves if needed
    If container.Type <> cdrCurveShape Then
        container.ConvertToCurves
    End If

    ' Get user settings
    Dim density As Integer
    Dim minSpacing As Double
    Dim allowRotation As Boolean
    Dim allowOverlap As Boolean

    If Not GetUserSettings(density, minSpacing, allowRotation, allowOverlap) Then
        Exit Sub ' User cancelled
    End If

    ' Perform random fill
    Call PerformRandomFill(container, fillElements, density, minSpacing, allowRotation, allowOverlap)

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Random Fill Error"
End Sub


' ==============================================================================
' Get user settings via InputBox
' ==============================================================================
Function GetUserSettings(ByRef density As Integer, _
                         ByRef minSpacing As Double, _
                         ByRef allowRotation As Boolean, _
                         ByRef allowOverlap As Boolean) As Boolean

    Dim userInput As String

    ' Density setting
    userInput = InputBox( _
        "FILL DENSITY:" & vbCrLf & vbCrLf & _
        "How many elements to place?" & vbCrLf & vbCrLf & _
        "  Low density:    50-100" & vbCrLf & _
        "  Medium density: 200-500" & vbCrLf & _
        "  High density:   1000-2000" & vbCrLf & _
        "  Very high:      3000+" & vbCrLf & vbCrLf & _
        "Enter number of attempts (50-5000):", _
        "Density Control", "500")

    If userInput = "" Then
        GetUserSettings = False
        Exit Function
    End If

    density = Val(userInput)
    If density < 50 Then density = 50
    If density > 5000 Then density = 5000

    ' Spacing setting
    userInput = InputBox( _
        "MINIMUM SPACING:" & vbCrLf & vbCrLf & _
        "Minimum distance between elements:" & vbCrLf & vbCrLf & _
        "  0%   = Elements can touch/overlap" & vbCrLf & _
        "  10%  = Small gap" & vbCrLf & _
        "  25%  = Medium gap" & vbCrLf & _
        "  50%  = Large gap" & vbCrLf & _
        "  -50% = Allow 50% overlap" & vbCrLf & vbCrLf & _
        "Enter percentage (-100 to 200):", _
        "Spacing Control", "0")

    If userInput = "" Then
        GetUserSettings = False
        Exit Function
    End If

    minSpacing = Val(userInput) / 100

    ' Rotation setting
    userInput = InputBox( _
        "ALLOW RANDOM ROTATION?" & vbCrLf & vbCrLf & _
        "Rotating elements helps fill space better:" & vbCrLf & vbCrLf & _
        "  1 = YES - Rotate elements randomly (0-360°)" & vbCrLf & _
        "  2 = NO  - Keep original angle" & vbCrLf & vbCrLf & _
        "Enter 1 or 2:", _
        "Rotation Control", "1")

    If userInput = "" Then
        GetUserSettings = False
        Exit Function
    End If

    allowRotation = (Val(userInput) = 1)

    ' Overlap setting
    userInput = InputBox( _
        "OVERLAP MODE:" & vbCrLf & vbCrLf & _
        "How to handle overlapping elements:" & vbCrLf & vbCrLf & _
        "  1 = ALLOW - Elements can overlap (faster, denser)" & vbCrLf & _
        "  2 = PREVENT - Check collisions (slower, cleaner)" & vbCrLf & vbCrLf & _
        "Enter 1 or 2:", _
        "Overlap Control", "1")

    If userInput = "" Then
        GetUserSettings = False
        Exit Function
    End If

    allowOverlap = (Val(userInput) = 1)

    GetUserSettings = True
End Function


' ==============================================================================
' Perform Random Fill Operation
' ==============================================================================
Sub PerformRandomFill(container As Shape, _
                      fillElements As Collection, _
                      density As Integer, _
                      minSpacing As Double, _
                      allowRotation As Boolean, _
                      allowOverlap As Boolean)

    On Error Resume Next

    ' Initialize random number generator
    Randomize Timer

    ' Get container boundaries
    Dim leftX As Double, rightX As Double
    Dim topY As Double, bottomY As Double

    leftX = container.LeftX
    rightX = container.RightX
    topY = container.TopY
    bottomY = container.BottomY

    ' Get average dimensions of fill elements
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

    ' Calculate minimum spacing distance
    Dim minDist As Double
    minDist = ((avgW + avgH) / 2) * minSpacing

    ' Collection to store placed element positions (for collision detection)
    Dim placedPositions As New Collection

    ' Start fill operation
    Application.Optimization = True
    ActiveDocument.BeginCommandGroup "Random Fill Curve"

    Dim attempt As Integer
    Dim placed As Integer
    Dim cx As Double, cy As Double
    Dim rotation As Double
    Dim elemIndex As Integer
    Dim canPlace As Boolean

    placed = 0
    elemIndex = 1

    ' Progress message
    MsgBox "Starting random fill with " & density & " attempts..." & vbCrLf & _
           "This may take a moment. Please wait.", vbInformation, "Processing"

    ' Random placement loop
    For attempt = 1 To density
        ' Generate random position within container bounds
        cx = leftX + Rnd() * (rightX - leftX)
        cy = bottomY + Rnd() * (topY - bottomY)

        ' Check if point is inside container
        If container.Curve.IsPointInside(cx, cy) Then

            ' Generate random rotation if allowed
            If allowRotation Then
                rotation = Rnd() * 360
            Else
                rotation = 0
            End If

            ' Check for collision if overlap not allowed
            canPlace = True

            If Not allowOverlap Then
                canPlace = CheckNoCollision(cx, cy, avgW, avgH, placedPositions, minDist)
            End If

            ' Place element if conditions met
            If canPlace Then
                ' Select source element (cycle through if multiple)
                Dim sourceElem As Shape
                Set sourceElem = fillElements(elemIndex)

                ' Create duplicate
                Dim newShape As Shape
                Set newShape = sourceElem.Duplicate

                ' Position and rotate
                newShape.CenterX = cx
                newShape.CenterY = cy

                If allowRotation Then
                    newShape.Rotate rotation
                End If

                ' Store position for collision detection
                If Not allowOverlap Then
                    Call AddPosition(placedPositions, cx, cy, avgW, avgH)
                End If

                placed = placed + 1

                ' Cycle to next element
                elemIndex = elemIndex + 1
                If elemIndex > fillElements.Count Then
                    elemIndex = 1
                End If
            End If
        End If
    Next attempt

    ' Complete operation
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveWindow.Refresh

    ' Show result
    MsgBox "Random Fill Complete!" & vbCrLf & vbCrLf & _
           "Attempts: " & density & vbCrLf & _
           "Successfully placed: " & placed & " elements" & vbCrLf & _
           "Success rate: " & Round((placed / density) * 100, 1) & "%", _
           vbInformation, "Success"
End Sub


' ==============================================================================
' Check if position has no collision with existing elements
' ==============================================================================
Function CheckNoCollision(cx As Double, cy As Double, _
                          w As Double, h As Double, _
                          placedPositions As Collection, _
                          minDist As Double) As Boolean

    On Error Resume Next

    Dim i As Integer
    Dim pos As Variant
    Dim dx As Double, dy As Double
    Dim dist As Double

    ' If no elements placed yet, no collision
    If placedPositions.Count = 0 Then
        CheckNoCollision = True
        Exit Function
    End If

    ' Check distance to all placed elements
    For i = 1 To placedPositions.Count
        pos = placedPositions(i)

        dx = cx - pos(0)
        dy = cy - pos(1)
        dist = Sqr(dx * dx + dy * dy)

        ' If too close, collision detected
        If dist < (minDist + (w + h) / 2) Then
            CheckNoCollision = False
            Exit Function
        End If
    Next i

    ' No collision
    CheckNoCollision = True
End Function


' ==============================================================================
' Add position to collection for collision tracking
' ==============================================================================
Sub AddPosition(placedPositions As Collection, _
                cx As Double, cy As Double, _
                w As Double, h As Double)

    On Error Resume Next

    Dim pos(3) As Double
    pos(0) = cx
    pos(1) = cy
    pos(2) = w
    pos(3) = h

    placedPositions.Add pos
End Sub


' ==============================================================================
' Quick Random Fill (Preset: High density, rotation, overlap allowed)
' ==============================================================================
Sub QuickRandomFill()
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

    ' Use preset: 1000 attempts, 0% spacing, rotation ON, overlap ON
    Call PerformRandomFill(container, fillElements, 1000, 0, True, True)

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Quick Random Fill Error"
End Sub

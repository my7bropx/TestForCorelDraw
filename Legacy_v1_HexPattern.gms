Sub FillWithObjects_HexPattern()
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    
    If sr.Count <> 2 Then
        MsgBox "⚠️ Please select exactly TWO objects: one to repeat, one as the container shape."
        Exit Sub
    End If
    
    Dim container As Shape, fillObject As Shape
    
    ' Determine which is the container (larger one)
    If sr(1).SizeWidth * sr(1).SizeHeight > sr(2).SizeWidth * sr(2).SizeHeight Then
        Set container = sr(1)
        Set fillObject = sr(2)
    Else
        Set container = sr(2)
        Set fillObject = sr(1)
    End If
    
    ' Convert to curves if needed and handle the reference properly
    If container.Type <> cdrCurveShape Then
        container.ConvertToCurves
        ' Refresh selection to get the converted shape
        sr.RemoveAll
        container.AddToSelection
        Set container = ActiveSelectionRange(1)
    End If
    
    ' Make sure container is a valid closed shape
    If container.Curve.Nodes.Count < 3 Then
        MsgBox "⚠️ Container shape seems invalid. Use a closed shape like a circle or outline."
        Exit Sub
    End If
    
    ' Get bounding box - Fixed coordinate system
    Dim leftX As Double: leftX = container.LeftX
    Dim rightX As Double: rightX = container.RightX
    Dim topY As Double: topY = container.TopY
    Dim bottomY As Double: bottomY = container.BottomY
    
    ' Get dimensions of the fill object
    Dim objW As Double: objW = fillObject.SizeWidth
    Dim objH As Double: objH = fillObject.SizeHeight
    
    ' Hexagonal pattern spacing
    Dim dx As Double: dx = objW * 0.75  ' Horizontal spacing for hex pattern
    Dim dy As Double: dy = objH * 0.866 ' Vertical spacing (√3/2)
    
    Dim x As Double, y As Double, cx As Double, cy As Double
    Dim row As Integer: row = 0
    Dim count As Integer: count = 0
    
    ' Start command group for undo
    ActiveDocument.BeginCommandGroup "Fill Container with Hex Pattern"
    Application.Optimization = True
    
    ' Fixed loop direction for CorelDRAW coordinate system
    y = bottomY + objH / 2
    Do While y <= topY - objH / 2
        x = leftX + objW / 2
        
        ' Offset every other row for hexagonal pattern
        If row Mod 2 = 1 Then
            x = x + dx / 2
        End If
        
        Do While x <= rightX - objW / 2
            cx = x
            cy = y
            
            ' Check if the center point is inside the container
            ' and also check if the object bounds would fit reasonably
            If container.Curve.IsPointInside(cx, cy) Then
                ' Additional check: ensure object doesn't extend too far outside
                Dim margin As Double: margin = objW * 0.3 ' Allow some overlap
                If cx - objW/2 >= leftX - margin And cx + objW/2 <= rightX + margin And _
                   cy - objH/2 >= bottomY - margin And cy + objH/2 <= topY + margin Then
                    
                    Dim newShape As Shape
                    Set newShape = fillObject.Duplicate
                    newShape.CenterX = cx
                    newShape.CenterY = cy
                    count = count + 1
                End If
            End If
            
            x = x + dx
        Loop
        
        row = row + 1
        y = y + dy
    Loop
    
    ' Clean up
    Application.Optimization = False
    ActiveWindow.Refresh
    ActiveDocument.EndCommandGroup
    
    MsgBox "✅ Done! Placed " & count & " objects in hexagonal pattern inside the shape."
End Sub

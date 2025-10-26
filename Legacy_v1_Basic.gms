Sub HexFill2024()

    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange

    If sr.Count <> 2 Then
        MsgBox "Select exactly TWO shapes: one fill object (like a red circle) and one container shape (closed)."
        Exit Sub
    End If

    Dim container As Shape, fillShape As Shape
    If sr(1).SizeWidth * sr(1).SizeHeight > sr(2).SizeWidth * sr(2).SizeHeight Then
        Set container = sr(1)
        Set fillShape = sr(2)
    Else
        Set container = sr(2)
        Set fillShape = sr(1)
    End If

    ' Convert container to curves if needed
    If container.Type <> cdrCurveShape Then
        Set container = container.ConvertToCurves
    End If

    If Not container.Curve.Closed Then
        MsgBox "Container shape must be a CLOSED curve."
        Exit Sub
    End If

    ' Dimensions of fill shape
    Dim w As Double: w = fillShape.SizeWidth
    Dim h As Double: h = fillShape.SizeHeight
    Dim dx As Double: dx = w
    Dim dy As Double: dy = h * 0.866  ' vertical spacing for hex pattern

    ' Use direct edge properties (no BoundingBox!)
    Dim leftX As Double: leftX = container.LeftX
    Dim rightX As Double: rightX = container.RightX
    Dim topY As Double: topY = container.TopY
    Dim bottomY As Double: bottomY = container.BottomY

    Dim x As Double, y As Double, cx As Double, cy As Double
    Dim row As Integer, count As Integer: count = 0

    Application.Optimization = True
    ActiveDocument.BeginCommandGroup "Hex Fill"

    row = 0
    For y = topY + h / 2 To bottomY - h / 2 Step dy
        For x = leftX + w / 2 To rightX - w / 2 Step dx
            cx = x
            If row Mod 2 = 1 Then cx = cx + dx / 2
            cy = y

            ' Test if center is inside the curve
            If container.Curve.IsPointInside(cx, cy) Then
                Dim newObj As Shape
                Set newObj = fillShape.Duplicate
                newObj.CenterX = cx
                newObj.CenterY = cy
                count = count + 1
            End If
        Next x
        row = row + 1
    Next y

    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveWindow.Refresh

    MsgBox "✅ Done! " & count & " objects filled inside the shape."

End Sub

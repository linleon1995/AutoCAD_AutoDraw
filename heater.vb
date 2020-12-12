




Public tube_head As Double
Public tube_diameter As Double
Public tube_num_row As Double
Public tube_num_stick As Double
Public efficient_dist As Double
Public fan_diameter As Double
Public motor_frame_length As Double
Public motor_frame_width As Double
Public thickness As Double
Public material As String
Public partition_material As String
Public num_motor As Double
Public inner_dist As Double
Public connect_width As Double
Public num_screw As Double
Public screw_dist As Double
Public num_part_screw As Double
Public is_expand As String
Public in_sb_d As Double
Public out_sb_in_d As Double ' 外端板穿面孔徑
Public out_sb_out_d As Double ' 外端板焊面孔徑

Public text_height As Double
Public text_dist As Double
Public arrow_size1 As Double
Public arrow_size2 As Double
Public dim_text_height1 As Double
Public dim_text_height2 As Double
Public dim_dist1 As Double
Public dim_dist2 As Double
Public dim_dist3 As Double
Public fan_type As Double

Public fin_length As Double
Public fin_width As Double
Public row_dist As Double
Public stick_dist As Double

Public Sub AddRectCircles(ByVal start, ByVal radius, ByVal length, ByVal width)
    Dim p(2)  As Double
    Dim cir_obj As AcadCircle
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(start, radius)
    p(0) = start(0) + length: p(1) = start(1)
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(p, radius)
    p(0) = start(0): p(1) = start(1) + width
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(p, radius)
    p(0) = start(0) + length: p(1) = start(1) + width
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(p, radius)
End Sub


Public Sub AddLinedCircles(ByVal start, ByVal radius, ByVal distance, ByVal num_circle, ByVal direction, _
                           Optional text_dir As Double = -1, _
                           Optional text_dist As Double = 100, _
                           Optional text_height As Double = 20, _
                           Optional arrow_size1 As Double = 20)
    Dim c(2)  As Double, c2(2)  As Double
    Dim cir_obj As AcadCircle
    Dim text_loc(2)  As Double
    c(0) = start(0): c(1) = start(1)
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(c, radius)

    cur_layer = ThisDrawing.ActiveLayer.Name
    c2(0) = c(0): c2(1) = c(1)
    For i = 1 To num_circle - 1
        SelectActiveLayer cur_layer
        If direction = 0 Then
            c(0) = c(0) + distance: c(1) = c(1)
        ElseIf direction = 1 Then
            c(0) = c(0): c(1) = c(1) + distance
        ElseIf direction = 2 Then
            c(0) = c(0) - distance: c(1) = c(1)
        ElseIf direction = 3 Then
            c(0) = c(0): c(1) = c(1) - distance
        End If
        Set cir_obj = ThisDrawing.ModelSpace.AddCircle(c, radius)

        If text_dir > 0 Then
            SelectActiveLayer "其他尺寸"
            If text_dir = 0 Then
                text_loc(0) = c(0) + text_dist: text_loc(1) = (c(1)+c2(1))/2
            ElseIf text_dir = 1 Then
                text_loc(0) = (c(0)+c2(0))/2: text_loc(1) = c(1) + text_dist
            ElseIf text_dir = 2 Then
                text_loc(0) = c(0) - text_dist: text_loc(1) = (c(1)+c2(1))/2
            ElseIf text_dir = 3 Then
                text_loc(0) = (c(0)+c2(0))/2: text_loc(1) = c(1) - text_dist
            End If
            ' text_loc(0) = (c(0)+c2(0))/2
            Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(c, c2, text_loc)
            c2(0) = c(0): c2(1) = c(1)
            AcadDimAligned.TextHeight = text_height
            AcadDimAligned.ArrowheadSize = arrow_size1
        End If
    Next
    SelectActiveLayer cur_layer
End Sub


Public Sub AddFinCircles(ByVal start, ByVal num_row, ByVal num_stick, ByVal row_dist, _
                         ByVal stick_dist, ByVal diameter)
    Dim start2(2)  As Double
    start2(0) = start(0) + row_dist: start2(1) = start(1) - stick_dist/2


    For i = 1 To num_row
        If i Mod 2 = 1 Then
            AddLinedCircles start, diameter/2, stick_dist, num_stick, 1
            start(0) = start(0) + 2*row_dist
        Else:
            AddLinedCircles start2, diameter/2, stick_dist, num_stick, 1
            start2(0) = start2(0) + 2*row_dist
        End If
    Next
End Sub


Public Sub SelectActiveLayer(ByVal layer_name)
    For Each lay0 In ThisDrawing.Layers ' 在所有的圖層中進行循環
        If lay0.Name = layer_name Then ' 如果找到圖層名
            ThisDrawing.ActiveLayer = lay0 ' 設定圖層為當前圖層
            Exit For ' 結束尋找
        End If
    Next lay0
End Sub


Public Sub AddRect(ByVal start, ByVal length, ByVal width)
    Dim line_obj As AcadLine
    Dim p1(2)  As Double
    Dim p2(2)  As Double
    Dim ends(2)  As Double
    
    p1(0) = start(0) + length: p1(1) = start(1)
    p2(0) = start(0): p2(1) = start(1) + width
    ends(0) = start(0) + length: ends(1) = start(1) + width
    
    Set line_obj = ThisDrawing.ModelSpace.AddLine(start, p1)
    Set line_obj = ThisDrawing.ModelSpace.AddLine(start, p2)
    Set line_obj = ThisDrawing.ModelSpace.AddLine(ends, p1)
    Set line_obj = ThisDrawing.ModelSpace.AddLine(ends, p2)
End Sub


Public Sub AddConnect(ByVal start1, ByVal end1, ByVal l1, ByVal l2, ByVal l3, ByVal w1, ByVal w2)
    Dim start2(2)  As Double
    If end1(0) > start1(0) And end1(1) > start1(1) Then
        AddRect start1, l1 + l2 + l3, w1
        start2(0) = start1(0) + l1: start2(1) = start1(1) + w1
        AddRect start2, l1 + l2, w1 + w2
    ElseIf end1(0) < start1(0) And end1(1) > start1(1) Then
        AddRect start1, -1 * w1, l1 + l2 + l3
        start2(0) = start1(0) - w1: start2(1) = start1(1) + l1
        AddRect start2, -1 * (w1 + w2), l1 + l2
    ElseIf end1(0) < start1(0) And end1(1) < start1(1) Then
        AddRect start1, -1 * (l1 + l2 + l3), -1 * w1
        start2(0) = start1(0) - l1: start2(1) = start1(1) - w1
        AddRect start2, -1 * (l1 + l2), -1 * (w1 + w2)
    ElseIf end1(0) > start1(0) And end1(1) < start1(1) Then
        AddRect start1, w1, -1 * (l1 + l2 + l3)
        start2(0) = start1(0) + w1: start2(1) = start1(1) - l1
        AddRect start2, w1 + w2, -1 * (l1 + l2)
    End If
End Sub


Public Sub AddCross(ByVal center, ByVal l1, ByVal l2, ByVal w1, ByVal w2)
    Dim p1(2) As Double
    Dim p2(2) As Double
    p1(0) = center(0) - l1: p1(1) = ccenter(1)
    p2(0) = center(0) + l2: p2(1) = ccenter(1)
    Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
    p1(0) = center(0): p1(1) = ccenter(1) + w2
    p2(0) = center(0): p2(1) = ccenter(1) - w1
    Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
End Sub


Public Sub AddHill(ByVal start, ByVal l1, ByVal l2, ByVal l3, ByVal w1, ByVal w2, ByVal direction, _
                   Optional add_dim, Optional text_height, Optional arrow_size)
    Dim p1(2) As Double
    Dim p2(2) As Double
    Dim p3(2) As Double
    Dim p4(2) As Double
    Dim p5(2) As Double
    Dim p6(2) As Double
    Dim t(2) As Double
    Dim line_obj As AcadLine
    
    ' TODO: Independent var for dimension
    If direction = "h" Then
        p1(0) = start(0): p1(1) = start(1)
        p2(0) = start(0) + l1: p2(1) = start(1)
        p3(0) = start(0) + l1: p3(1) = start(1) + w1
        p4(0) = start(0) + l1 + l2: p4(1) = start(1) + w1
        p5(0) = start(0) + l1 + l2: p5(1) = start(1) + w1 - w2
        p6(0) = start(0) + l1 + l2 + l3: p6(1) = start(1) + w1 - w2
        
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p2, p3)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p3, p4)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p4, p5)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p5, p6)

        For i = 0 to 4
            If add_dim(i) = 1 Then
                t(0) = (p1(0) + p2(0)) / 2: t(1) = p1(0) + dim_dist1
                Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, t)
            ElseIf add_dim(i) = 2 Then
                t(0) = p2(0) - dim_dist1: t(1) = (p2(1) + p3(1)) / 2
                Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p2, p3, t)
            ElseIf add_dim(i) = 3 Then
                t(0) = (p3(0) + p4(0)) / 2: t(1) = p3(1) + dim_dist1
                Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p3, p4, t)
            ElseIf add_dim(i) = 4 Then
                t(0) = p4(0) + dim_dist1: t(1) = (p4(1) + p5(1)) / 2
                Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p4, p5, t)
            ElseIf add_dim(i) = 5 Then
                t(0) = (p5(0) + p6(0)) / 2: t(1) = p5(1) + dim_dist1
                Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p5, p6, t)
            End If
            AcadDimAligned.TextHeight = dim_text_height2
            AcadDimAligned.ArrowheadSize = arrow_size2
        Next

    ElseIf direction = "h_flip" Then
        p1(0) = start(0): p1(1) = start(1)
        p2(0) = start(0) + l1: p2(1) = start(1)
        p3(0) = start(0) + l1: p3(1) = start(1) - w1
        p4(0) = start(0) + l1 + l2: p4(1) = start(1) - w1
        p5(0) = start(0) + l1 + l2: p5(1) = start(1) - w1 + w2
        p6(0) = start(0) + l1 + l2 + l3: p6(1) = start(1) - w1 + w2
        
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p2, p3)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p3, p4)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p4, p5)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p5, p6)

        SelectActiveLayer "其他尺寸"
        t(0) = p3(0) - dim_dist1: t(1) = (p2(1) + p3(1)) / 2
        Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p2, p3, t)
        AcadDimAligned.TextHeight = text_height
        AcadDimAligned.ArrowheadSize = arrow_size

        ' For i = 0 To 4
        '     If add_dim(i) = 1 Then
        '         t(0) = (p1(0) + p2(0)) / 2: t(1) = p1(0) - dim_dist1
        '         Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, t)
        '     ElseIf add_dim(i) = 2 Then
        '         t(0) = p2(0) - dim_dist1: t(1) = (p2(1) + p3(1)) / 2
        '         Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p2, p3, t)
        '     ElseIf add_dim(i) = 3 Then
        '         t(0) = (p3(0) + p4(0)) / 2: t(1) = p3(1) - dim_dist1
        '         Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p3, p4, t)
        '     ElseIf add_dim(i) = 4 Then
        '         t(0) = p4(0) + dim_dist1: t(1) = (p4(1) + p5(1)) / 2
        '         Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p4, p5, t)
        '     ElseIf add_dim(i) = 5 Then
        '         t(0) = (p5(0) + p6(0)) / 2: t(1) = p5(1) - dim_dist1
        '         Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p5, p6, t)
        '     End If
        '     AcadDimAligned.TextHeight = dim_text_height2
        '     AcadDimAligned.ArrowheadSize = arrow_size2
        ' Next

        ' p1(0) = start(0) + l1: p1(1) = start(1)
        ' Set line_obj = ThisDrawing.ModelSpace.AddLine(start, p1)
        ' p2(0) = p1(0): p2(1) = p1(1) - w1
        ' Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        ' p1(0) = p2(0) + l2: p1(1) = p2(1)
        ' Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        ' p2(0) = p1(0): p2(1) = p1(1) + w2
        ' Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        ' p1(0) = p2(0) + l3: p1(1) = p2(1)
        ' Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
    ElseIf direction = "v" Then
        p1(0) = start(0): p1(1) = start(1) + l1
        Set line_obj = ThisDrawing.ModelSpace.AddLine(start, p1)

        p2(0) = p1(0) + w1: p2(1) = p1(1)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        p1(0) = p2(0): p1(1) = p2(1) + l2
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        p2(0) = p1(0) - w2: p2(1) = p1(1)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        p1(0) = p2(0): p1(1) = p2(1) + l3
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
    ElseIf direction = "v_flip" Then
        p1(0) = start(0): p1(1) = start(1) + l1
        Set line_obj = ThisDrawing.ModelSpace.AddLine(start, p1)
        p2(0) = p1(0) - w1: p2(1) = p1(1)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        p1(0) = p2(0): p1(1) = p2(1) + l2
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        p2(0) = p1(0) + w2: p2(1) = p1(1)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        p1(0) = p2(0): p1(1) = p2(1) + l3
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
    End If

    SelectActiveLayer "鈑金"
End Sub


Public Sub AddArcwithLines(ByVal center, ByVal radius, ByVal max_dist)
	Dim arcObj As AcadArc
	Dim startAngleInDegree As Double
	Dim endAngleInDegree As Double
    Dim p1(2) As Double, p2(2) As Double
    x = (max_dist - radius*2) / 2

    ' Add side lines
    p1(0) = center(0) - radius: p1(1) = center(1) - x
    p2(0) = center(0) - radius: p2(1) = center(1) + x
    Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
    p1(0) = center(0) + radius: p1(1) = center(1) - x
    p2(0) = center(0) + radius: p2(1) = center(1) + x
    Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)

    ' Add arc
	' Define the circle
	startAngleInDegree = 180
	endAngleInDegree = 0

	' Convert the angles in degrees to angles in radians
	startAngleInRadian = startAngleInDegree * 3.141592 / 180
	endAngleInRadian = endAngleInDegree * 3.141592 / 180

	' Create the arc object in model space
    p1(0) = center(0): p1(1) = center(1) - x
	Set arcObj = ThisDrawing.ModelSpace.AddArc(p1, radius, startAngleInRadian, endAngleInRadian)
    p1(0) = center(0): p1(1) = center(1) + x
	Set arcObj = ThisDrawing.ModelSpace.AddArc(p1, radius, endAngleInRadian, startAngleInRadian)
End Sub


Public Sub AddMainDims(ByVal h1, ByVal h2, ByVal h_dist, ByVal h_dir, _
                       ByVal v1, ByVal v2, ByVal v_dist, ByVal v_dir, _
                       ByVal text_height, ByVal arrow_size1)
    Dim h_text(2)  As Double, v_text(2)  As Double
    If h_dir = "up" Then
        h_text(0) = h1(0) + Abs(h1(0) - h2(0)) / 2: h_text(1) = h1(1) + h_dist
    ElseIf h_dir = "down" Then
        h_text(0) = h1(0) + Abs(h1(0) - h2(0)) / 2: h_text(1) = h1(1) - h_dist
    End If
    
    If v_dir = "right" Then
        v_text(0) = v1(0) + v_dist: v_text(1) = v1(1) + Abs(v1(1) - v2(1)) / 2
    ElseIf v_dir = "left" Then
        v_text(0) = v1(0) - v_dist: v_text(1) = v1(1) + Abs(v1(1) - v2(1)) / 2
    End If

    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(h1, h2, h_text)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size1
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(v1, v2, v_text)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size1
End Sub


Public Sub AddTwoCrossRects(ByVal start, ByVal length, ByVal width, ByVal l1, ByVal l2, _
                            ByVal w1, ByVal w2, Optional text_height, Optional arrow_size)
    
    Dim start_r1(2)  As Double, start_r2(2)  As Double
    Dim a(2)  As Double, b(2)  As Double, c(2)  As Double, d(2)  As Double, _
        e(2)  As Double, f(2)  As Double, g(2)  As Double, h(2)  As Double, i(2)  As Double
    Dim t1(2)  As Double, t2(2)  As Double, t3(2)  As Double, t4(2)  As Double, _
        t5(2)  As Double, t6(2)  As Double, t7(2)  As Double, t8(2)  As Double
    start_r1(0) = start(0): start_r1(1) = start(1) + w1
    start_r2(0) = start(0) + l1: start_r2(1) = start(1)
    in_length = length - l1 - l2
    in_width = width - w1 - w2
    
    ' TODO: 參數拉出去
    ' TODO: dimension in right position
    ' text_height = 25
    ' arrow_size1 = 20
    ' dim_dist1 = 100
    a(0) = start_r2(0) + in_length: a(1) = start_r2(1)
    b(0) = start_r2(0) + in_length: b(1) = start_r1(1)
    c(0) = start_r2(0) + in_length + l2: c(1) = start_r1(1)
    d(0) = start_r1(0) + length: d(1) = start_r1(1) + in_width
    e(0) = start_r2(0) + in_length: e(1) = start_r1(1) + in_width
    f(0) = start_r2(0) + in_length: f(1) = start_r2(1) + width
    g(0) = start_r1(0) + l1: g(1) = start_r2(1) + width
    h(0) = start_r1(0) + l1: h(1) = start_r1(1) + in_width
    i(0) = start_r1(0): i(1) = start_r1(1) + in_width

    ' t1(0) = start_r1(0) + 0.5 * length: t1(1) = start_r1(1) - dim_dist1 - w1
    ' t2(0) = start_r2(0) - dim_dist1 - l1: t2(1) = start_r2(1) + 0.5 * width
    t3(0) = start_r1(0) + length + dim_dist1: t3(1) = start_r2(1) + 0.5 * w1
    t4(0) = start_r1(0) + length + dim_dist1: t4(1) = start_r2(1) + w1 + 0.5 * in_width
    t5(0) = start_r1(0) + length + dim_dist1: t5(1) = start_r2(1) + w1 + in_width + 0.5 * w2
    t6(0) = start_r1(0) + 0.5 * l1 + in_length + 0.5 * l2: t6(1) = start_r2(1) + width + dim_dist1
    t7(0) = start_r1(0) + l1 + 0.5 * in_length: t7(1) = start_r2(1) + width + dim_dist1
    t8(0) = start_r1(0) + 0.5 * l1 - 50: t8(1) = start_r2(1) + width + dim_dist1

    SelectActiveLayer "鈑金"
    AddRect start_r1, length, in_width
    AddRect start_r2, in_length, width

    ' SelectActiveLayer "主尺寸"
    SelectActiveLayer "其他尺寸"
    AddMainDims start_r1, c, dim_dist1, "down", start_r2, g, dim_dist1, "left", text_height, arrow_size1

    ' SelectActiveLayer "其他尺寸"
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(a, b, t3)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size1
    TextOutsideAlign = 1
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(b, e, t4)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size1
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(e, f, t5)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size1
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(e, d, t6)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size1
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(e, h, t7)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size1
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(h, i, t8)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size1
    SelectActiveLayer "鈑金"
End Sub


Public Sub fans_board(ByVal start, ByVal d1, ByVal o_v2, ByVal o_v3, ByVal o_v4, ByVal o_v5, _
                      ByVal o_v6, ByVal comp_width, ByVal d3)
    Dim line_obj As AcadLine
    Dim p1(2) As Double
    Dim p2(2) As Double
    Dim p3(2) As Double
    Dim p4(2) As Double
    Dim p5(2) As Double
    Dim add_dim(4) As Double
    Dim text_loc(2) As Double
    Dim chordPoint(2) As Double, FarchordPoint(2) As Double
    comp_length = efficient_dist + 2*o_v4

    ' 鈑金
    ' TODO: complete AddHill dimension function
    ' add_dim(0) = 1: add_dim(1) = 2: add_dim(2) = 3: add_dim(3) = 4: add_dim(4) = 5
    AddHill start, 0, comp_length, 0, o_v4, o_v4, "h_flip", text_height:=dim_text_height1, arrow_size:=arrow_size1
    p1(0) = start(0): p1(1) = start(1) + o_v2
    AddHill p1, 0, comp_length, 0, o_v2, o_v2, "h_flip", text_height:=dim_text_height1, arrow_size:=arrow_size1
    p1(0) = p1(0): p1(1) = p1(1) + o_v3
    AddHill p1, 0, comp_length, 0, o_v3, o_v3, "h_flip", text_height:=dim_text_height1, arrow_size:=arrow_size1
    p1(0) = p1(0): p1(1) = p1(1) + o_v2
    AddHill p1, 0, comp_length, 0, o_v2, o_v2, "h_flip", text_height:=dim_text_height1, arrow_size:=arrow_size1
    p1(0) = p1(0): p1(1) = p1(1) + o_v4
    AddHill p1, 0, comp_length, 0, o_v4, o_v4, "h_flip", text_height:=dim_text_height1, arrow_size:=arrow_size1
    p2(0) = p1(0) + comp_length: p2(1) = p1(1)
    Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)

    SelectActiveLayer "其他尺寸"
    '總長寬標註
    text_loc(0) = (p1(0) + p2(0)) / 2: text_loc(1) = p1(1) + dim_dist1
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, text_loc)
    AcadDimAligned.TextHeight = dim_text_height1
    AcadDimAligned.ArrowheadSize = arrow_size1

    p1(0) = p2(0): p1(1) = p2(1) - comp_width
    text_loc(0) = p1(0) + dim_dist1: text_loc(1) = (p1(1) + p2(1)) / 2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, text_loc)
    AcadDimAligned.TextHeight = dim_text_height1
    AcadDimAligned.ArrowheadSize = arrow_size1
    
    ' 上下折面及連接處螺絲
    SelectActiveLayer "鈑金"
    p1(0) = start(0) + o_v4/2: p1(1) = start(1) - o_v5
    AddRectCircles p1, d1/2, comp_length-o_v4, comp_width-2*o_v5
    p1(0) = start(0) + o_v4/2: p1(1) = start(1) + o_v2 - thickness - o_v6
    AddRectCircles p1, d1/2, comp_length-o_v4, o_v3+2*(thickness+o_v6)
    x = o_v6 + thickness + (o_v3 - screw_dist) / 2

    ' 上下折面及連接處螺絲 標註
    SelectActiveLayer "其他尺寸"
    p2(0) = p1(0): p2(1) = p1(1) + o_v6 + thickness
    text_loc(0) = p1(0) - dim_dist2: text_loc(1) = (p1(1) + p2(1)) / 2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2

    p1(1) = p1(1) + o_v6 + thickness + o_v3: p2(1) = p2(1) + o_v6 + thickness + o_v3
    text_loc(0) = p1(0) - dim_dist2: text_loc(1) = (p1(1) + p2(1)) / 2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2

    chordPoint(0) = p2(0) + d1/2/Sqr(2): chordPoint(1) = p2(1) + d1/2/Sqr(2)
    FarchordPoint(0) = p2(0) - d1/2/Sqr(2): FarchordPoint(1) = p2(1) - d1/2/Sqr(2)
    Set dim_obj = ThisDrawing.ModelSpace.AddDimDiametric(chordPoint, FarchordPoint, LeaderLength:=5)
    dim_obj.TextHeight = dim_text_height2
    dim_obj.ArrowheadSize = arrow_size2

    ' 風斗側邊螺絲孔
    SelectActiveLayer "鈑金"
    p1(0) = start(0) + o_v4/2: p1(1) = start(1) + o_v2 - thickness - o_v6 + x
    AddLinedCircles p1, d1/2, screw_dist/(num_screw-1), num_screw, 1, _
                    text_dist:=dim_dist2, text_dir:=2, text_height:=dim_text_height2, arrow_size1:=arrow_size1

    ' 風斗側邊螺絲孔 標註
    SelectActiveLayer "其他尺寸"
    p2(0) = p1(0) - (connect_width+thickness)/2: p2(1) = p1(1)
    text_loc(0) = (p1(0) + p2(0)) / 2: text_loc(1) = p1(1) - dim_dist3
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2
    SelectActiveLayer "鈑金"
    p1(0) = p1(0) + comp_length - connect_width - thickness
    AddLinedCircles p1, d1/2, screw_dist/(num_screw-1), num_screw, 1

    ' 風斗與馬達架螺絲孔
    comp = 4 - (Int(efficient_dist/num_motor) Mod 4)
    dist = Int(efficient_dist/num_motor) + comp
    p1(0) = start(0) + comp_length/2: p1(1) = start(1) + comp_width/2
    p1(0) = p1(0) - (0.5*num_motor-0.5)*dist
    AddLinedCircles p1, fan_diameter/2, dist, num_motor, 0

    p2(0) = p1(0) - motor_frame_length/2: p2(1) = p1(1) - motor_frame_width/2
    p3(0) = p1(0) - 70: p3(1) = p1(1) - 70
    p5(0) = p1(0) - 70: p5(1) = p1(1)
    For i = 1 to num_motor
        ' 馬達架螺絲孔
        SelectActiveLayer "螺絲"
        ' TODO: 
        AddRectCircles p2, 11.5/2, motor_frame_length, motor_frame_width
        ' AddCross p2, 11.5, 11.5, 11.5, 11.5
        ' 風扇文字說明
        SelectActiveLayer "鈑金"
        ' TODO: error message
        If is_expand = "y" Then
            t = "抽唇"
        ElseIf is_expand = "n" Then
            t = "不抽唇"
        End If
        ' 50 text height
        ThisDrawing.ModelSpace.AddText fan_type & "''", p5, 50
        ThisDrawing.ModelSpace.AddText t, p3, 50

        ' 側板螺絲孔
        If i < num_motor Then
            SelectActiveLayer "鈑金"
            p4(0) = p1(0) + dist/2: p4(1) = p1(1) - screw_dist/2
            If i = 1 Then
                AddLinedCircles p4, d1/2, screw_dist/(num_part_screw-1), num_part_screw, 1, _
                                text_dist:=10, text_dir:=2, text_height:=dim_text_height2, arrow_size1:=arrow_size2
            Else 
                AddLinedCircles p4, d1/2, screw_dist/(num_part_screw-1), num_part_screw, 1
            End If   
        End If

        p1(0) = p1(0) + dist
        p2(0) = p2(0) + dist
        p3(0) = p3(0) + dist
        p5(0) = p5(0) + dist
    Next

    SelectActiveLayer "鈑金"
End Sub


Public Sub inner_side_board(ByVal start, ByVal d)
    Dim p(2) As Double, p2(2) As Double
    Dim comp_length As Double
    Dim comp_width As Double
    Dim text_loc(2) As Double
    Dim chordPoint(2) As Double, FarchordPoint(2) As Double

    SelectActiveLayer "螺絲"
    p(0) = start(0): p(1) = start(1) + 1.5*stick_dist
    AddFinCircles p, tube_num_row, 1, row_dist, stick_dist, d
    p(0) = start(0): p(1) = start(1) + (tube_num_stick-1-1.5)*stick_dist
    AddFinCircles p, tube_num_row, 1, row_dist, stick_dist, d

    ' If tube_num_stick > 10
    SelectActiveLayer "鈑金"
    ' 銅管孔
    AddFinCircles start, tube_num_row, tube_num_stick, row_dist, stick_dist, in_sb_d
    p(0) = start(0) - 0.5*row_dist: p(1) = start(1) - 0.75*stick_dist
    comp_length = tube_num_row*row_dist
    comp_width = tube_num_stick*stick_dist
    ' 外框
    AddRect p, comp_length, comp_width

    ' 外框尺寸標註
    SelectActiveLayer "其他尺寸"
    p2(0) = p(0) + comp_length: p2(1) = p(1)
    text_loc(0) = p(0) + 0.5*comp_length: text_loc(1) = p(1) - dim_dist1
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p, p2, text_loc)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size1
    p2(0) = p(0): p2(1) = p(1) + comp_width
    text_loc(0) = p(0) - dim_dist1: text_loc(1) = p(1) + 0.5*comp_width
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p, p2, text_loc)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size1

    ' 銅管孔尺寸標註
    ' p2(0) = start(0): p2(1) = start(1) + (tube_num_stick-1-1.5)*stick_dist
    If tube_num_row Mod 2 = 0 Then
        comp = 0.5
    Else
        comp = 0
    End If
        
    p2(0) = start(0) + (tube_num_row-1)*row_dist
    p2(1) = start(1) + Int(tube_num_stick/2)*stick_dist - comp*stick_dist
    chordPoint(0) = p2(0) + in_sb_d/2/Sqr(2): chordPoint(1) = p2(1) - in_sb_d/2/Sqr(2)
    FarchordPoint(0) = p2(0) - in_sb_d/2/Sqr(2): FarchordPoint(1) = p2(1) + in_sb_d/2/Sqr(2)
    ' TODO: LeaderLength:=5
    Set dim_obj = ThisDrawing.ModelSpace.AddDimDiametric(chordPoint, FarchordPoint, LeaderLength:=5)
    dim_obj.TextHeight = dim_text_height2
    dim_obj.ArrowheadSize = arrow_size2

    p2(0) = start(0) + (tube_num_row-1)*row_dist
    p2(1) = start(1) + (tube_num_stick-1-1.5)*stick_dist - comp*stick_dist
    chordPoint(0) = p2(0) + d/2/Sqr(2): chordPoint(1) = p2(1) - d/2/Sqr(2)
    FarchordPoint(0) = p2(0) - d/2/Sqr(2): FarchordPoint(1) = p2(1) + d/2/Sqr(2)
    Set dim_obj = ThisDrawing.ModelSpace.AddDimDiametric(chordPoint, FarchordPoint, LeaderLength:=5)
    dim_obj.TextHeight = dim_text_height2
    dim_obj.ArrowheadSize = arrow_size2

    SelectActiveLayer "鈑金"
End Sub


Public Sub partition(ByVal start, ByVal v1, ByVal v2, ByVal v3, ByVal d1, ByVal dist1)
    Dim comp_length As Double
    Dim comp_width As Double
    Dim p(2) As Double: Dim p2(2) As Double
    Dim text_loc(2) As Double
	Dim center(2) As Double
	Dim chordPoint(2) As Double, FarchordPoint(2) As Double
	Dim leaderLen As Integer

    ' 鈑金
    comp_length = fin_length + 2*v1
    comp_width = inner_dist - v3 + 2*v1
    ' comp_width = Int(comp_width+1)
    AddTwoCrossRects start, comp_length, comp_width, v1, v1, v1, v1, _
                     text_height:=dim_text_height1, arrow_size:=arrow_size

    ' 螺絲
    p(0) = start(0) + v1 + (fin_length-screw_dist)/2: p(1) = start(1) + comp_width - dist1
    AddLinedCircles p, d1/2, screw_dist/(num_part_screw-1), num_part_screw, 0, _
                    text_dist:=dim_dist2, text_dir:=1, text_height:=dim_text_height2, _
                    arrow_size1:=arrow_size2

    ' 螺絲孔標註
    SelectActiveLayer "其他尺寸"
	FarchordPoint(0) = p(0) - d1/2/Sqr(2): FarchordPoint(1) = p(1) + d1/2/Sqr(2)
	chordPoint(0) = p(0) + d1/2/Sqr(2): chordPoint(1) = p(1) - d1/2/Sqr(2)
    Set dim_obj = ThisDrawing.ModelSpace.AddDimDiametric(chordPoint, FarchordPoint, LeaderLength:=5)
    dim_obj.TextHeight = dim_text_height2
    dim_obj.ArrowheadSize = arrow_size2

    SelectActiveLayer "鈑金"
End Sub


Public Sub outer_side_board(ByVal start, ByVal start2, ByVal v1, ByVal v2, ByVal v3, ByVal v4, _
                            ByVal v5, ByVal v6, ByVal fin_length, ByVal fin_width, ByVal d1)
    Dim p1(2) As Double, p2(2) As Double, p3(2) As Double
    Dim line_obj As AcadLine
    Dim text_loc(2) As Double
    Dim chordPoint(2) As Double, FarchordPoint(2) As Double
    
    comp_length = inner_dist + v3 + (tube_num_row-1)*row_dist + 2*connect_width
    comp_width = fin_length + connect_width + v1
    If comp_width - Int(comp_width) = 0 Then
        tmp = comp_width + 2*v2
    Else    
        tmp =  Int(comp_width + 2*v2 + 1)
    End If
    ext = (tmp-comp_width)/2
    comp_width = tmp

    ' 穿面
    AddTwoCrossRects start, comp_length, comp_width,  connect_width, connect_width, _
                     v1,  connect_width, text_height:=dim_text_height1, arrow_size:=arrow_size

    ' 焊面
    ' 外框
    AddTwoCrossRects start2, comp_length, comp_width,  connect_width,connect_width, _
                     v1,  connect_width, text_height:=dim_text_height1, arrow_size:=arrow_size
    ' 左側連接處螺絲孔
    p1(0) = start2(0) + (connect_width+thickness)/2: p1(1) = start2(1) + v1 + (fin_length+2*v2-screw_dist)/2
    AddLinedCircles p1, d1/2, screw_dist/(num_screw-1), num_screw, 1, _
                    text_dist:=dim_dist2, text_dir:=2, text_height:=dim_text_height2, arrow_size1:=arrow_size2
    
    SelectActiveLayer "其他尺寸"
    chordPoint(0) = p1(0) + d1/2/Sqr(2): chordPoint(1) = p1(1) + d1/2/Sqr(2)
    FarchordPoint(0) = p1(0) - d1/2/Sqr(2): FarchordPoint(1) = p1(1) - d1/2/Sqr(2)
    Set dim_obj = ThisDrawing.ModelSpace.AddDimDiametric(chordPoint, FarchordPoint, LeaderLength:=5)
    dim_obj.TextHeight = dim_text_height2
    dim_obj.ArrowheadSize = arrow_size2


    p1(1) = p1(1) + screw_dist
    p3(0) = start2(0): p3(1) = p1(1)
    text_loc(0) = (p1(0)+p3(0))/2: text_loc(1) = p1(1) + dim_dist2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p3, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2

    ' 右側連接處螺絲孔
    SelectActiveLayer "鈑金"
    p1(0) = start2(0) + comp_length -  connect_width/2: p1(1) = start2(1) + v1 + v5
    AddLinedCircles p1, d1/2, fin_length+2*v2-2*v5, 2, 1

    SelectActiveLayer "其他尺寸"
    p3(0) = p1(0): p3(1) = p1(1) - v5
    text_loc(0) = p1(0) + dim_dist2: text_loc(1) = (p1(1)+p3(1))/2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p3, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2
    
    p1(1) = p1(1) + fin_length+2*v2-2*v5
    p3(0) = p1(0): p3(1) = p1(1) + v5
    text_loc(0) = p1(0) + dim_dist2: text_loc(1) = (p1(1)+p3(1))/2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p3, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2

    ' 上連接處螺絲孔
    SelectActiveLayer "鈑金"
    p1(0) = start2(0) +  connect_width + v6: p1(1) = start2(1) + comp_width - connect_width/2
    AddLinedCircles p1, d1/2, 0, 1, 0, _
                    text_dist:=dim_dist2, text_dir:=2, text_height:=dim_text_height1, arrow_size1:=arrow_size1
    SelectActiveLayer "其他尺寸"
    p3(0) = p1(0) - v6: p3(1) = p1(1)
    text_loc(0) = (p1(0)+p3(0))/2: text_loc(1) = p1(1) + dim_dist2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p3, text_loc)
    AcadDimAligned.TextHeight = dim_text_height1
    AcadDimAligned.ArrowheadSize = arrow_size1

    ' 長螺絲孔
    SelectActiveLayer "鈑金"
    ' TODO: 20 12 8
    max_d = 12
    d2 = 8
    x1 = 20
    x2 = 8
    x3 = 10
    p1(0) = start2(0) +  connect_width + x1: p1(1) = start2(1) + x2
    AddArcwithLines p1, d2/2, max_d
    p3(0) = p1(0) + comp_length - 2*connect_width - x1 - x3: p3(1) = p1(1)
    AddArcwithLines p3, d2/2, max_d

    ' 長螺絲孔 標註

    ' 銅管孔
    p1(0) = start2(0) +  connect_width + inner_dist: p1(1) = start2(1) + v1 + ext + stick_dist*3/4
    ' TODO: 10
    AddFinCircles p1, tube_num_row, tube_num_stick, row_dist, stick_dist, 10

    SelectActiveLayer "螺絲"
    p2(0) = p1(0): p2(1) = p1(1) + 1.5*stick_dist
    AddFinCircles p2, tube_num_row, 1, row_dist, stick_dist, d1
    p2(0) = p1(0): p2(1) = p1(1) + (tube_num_stick-1-1.5)*stick_dist
    AddFinCircles p2, tube_num_row, 1, row_dist, stick_dist, d1
    ' p2(0) = p1(0): p2(1) = p1(1) + (tube_num_stick/2+0.5)*stick_dist # Add Round for odd sticks
    ' AddFinCircles p2, tube_num_row, 1, row_dist, stick_dist, d1
    SelectActiveLayer "鈑金"

End Sub


Public Sub heater() ' 一般熱排
    Set lay1 = ThisDrawing.Layers.Add("鈑金") ' 增加一個名為“鈑金”的圖層
    Set lay2 = ThisDrawing.Layers.Add("主尺寸") ' 增加一個名為“主尺寸”的圖層
    lay2.color = 1 ' 圖層設置為紅色
    Set lay3 = ThisDrawing.Layers.Add("其他尺寸") ' 增加一個名為“其他尺寸”的圖層
    lay3.color = 3 ' 圖層設置為綠色
    Set lay4 = ThisDrawing.Layers.Add("螺絲") ' 增加一個名為“螺絲”的圖層
    lay4.color = 6 ' 圖層設置為洋紅
    ThisDrawing.ActiveLayer = lay1 ' 將“鈑金”設置為當前圖層

    On Error Resume Next ' 如果有錯誤, 不管他
    ' 刪除所有作圖
    For Each oEntity In ThisDrawing.ModelSpace
        oEntity.Delete
    Next

    ' 10 HP 3分5R20T 1454 mm
    Dim start(2)  As Double ' 全圖作圖起始點
    Dim comp_start(2)  As Double ' 配件作圖起始點
    Dim comp_start2(2)  As Double
    start(0) = 1000: start(1) = 1000
    
    ' TODO: 輸入不合法 報錯
    ' TODO: 管徑 彎頭 考慮下拉式選單
    tube_head = 22
    tube_diameter = 3
    tube_num_row = 5
    tube_num_stick = 20
    efficient_dist = 1454
    fan_diameter = 295
    motor_frame_length = 272
    motor_frame_width = 272
    thickness = 1
    material = "錏板"
    partition_material = "鋁板"
    num_motor = 4
    inner_dist = 101
    connect_width = 14
    num_screw = 4
    screw_dist = 405
    num_part_screw = 3
    is_expand = "y"

    ' tube_head = 19.05
    ' tube_diameter = 2.5
    ' tube_num_row = 4
    ' tube_num_stick = 14
    ' efficient_dist = 780
    ' fan_diameter = 295
    ' motor_frame_length = 272.2
    ' motor_frame_width = 272.2
    ' thickness = 1
    ' material = "錏板"
    ' partition_material = "鋁板"
    ' num_motor = 2
    ' inner_dist = 106.9
    ' connect_width = 11
    ' num_screw = 4
    ' screw_dist = 282
    ' num_part_screw = 3
    ' is_expand = "y"

    ' 排支數換算鰭片長寬 & 穿管孔徑
    If tube_diameter = 2.5 Then
        row_dist = 19.05
        stick_dist = 25.4
        in_sb_d = 8.45
        ' out_sb_in_d = 13.2
        ' out_sb_out_d = 14
    ElseIf tube_diameter = 3 Then
        in_sb_d = 10.2
        ' out_sb_in_d = 13.2
        ' out_sb_out_d = 14
        If tube_head = 22 Then
            row_dist = 22
            stick_dist = 25.4
        ElseIf tube_head = 19.05 Then
            row_dist = 19.05
            stick_dist = 25.4
        End If
    ElseIf tube_diameter = 4 Then
        in_sb_d = 13.2
        ' out_sb_in_d = 14
        ' out_sb_out_d = 16.5
        row_dist = 33
        stick_dist = 38.1
    ElseIf tube_diameter = 5 Then
        in_sb_d = 16.5
        ' out_sb_in_d = 20
        ' out_sb_out_d = 20
        row_dist = 33
        stick_dist = 38.1
    ' TODO: 
    ' Else
        ' error message
    End If

    fin_length = tube_num_stick * stick_dist
    fin_width = tube_num_row * row_dist 

    ' 風扇孔徑換算風扇大小
    If is_expand = "y" Then
        If fan_diameter = 228 Then
            fan_type = 9
        ElseIf fan_diameter = 246 Then
            fan_type = 10
        ElseIf fan_diameter = 295 Then
            fan_type = 12
        ElseIf fan_diameter = 329 Then
            fan_type = 14
        ElseIf fan_diameter = 389 Then
            fan_type = 16
        ElseIf fan_diameter = 425 Then
            fan_type = 18
        ElseIf fan_diameter = 485 Then
            fan_type = 20
        ElseIf fan_diameter = 571 Then
            fan_type = 24
        Else
            fan_type = -1
        End If
    ElseIf is_expand = "n" Then
        If fan_diameter = 243 Then
            fan_type = 9
        ElseIf fan_diameter = 264 Then
            fan_type = 10
        ElseIf fan_diameter = 312 Then
            fan_type = 12
        ElseIf fan_diameter = 357 Then
            fan_type = 14
        ElseIf fan_diameter = 422 Then
            fan_type = 16
        ElseIf fan_diameter = 465 Then
            fan_type = 18
        ElseIf fan_diameter = 525 Then
            fan_type = 20
        ElseIf fan_diameter = 611 Then
            fan_type = 24
        Else
            fan_type = -1
        End If
    End If
    d1 = 5
    d2 = 3.2
    d3 = 11.5

    screw1 = 5
    screw2 = 3.2
    text_height = 30
    text_dist = 200
    arrow_size1 = 24
    arrow_size2 = 14
    dim_text_height1 = 26
    dim_text_height2 = 20
    dim_dist1 = 100
    dim_dist2 = 50
    dim_dist3 = 25

    ' 風斗板常數
    f_v1 = 2
    ' 隔板常數
    p_v1 = 11  ' 連接處寬 
    p_v2 = 3
    p_v3 = 20 ' 非鰭片空間扣除預留空間
    p_v4 = 6  ' 螺絲孔與邊緣距離
    ' 外端板常數
    o_v1 = 21
    o_v2 = (tube_num_row-1)*row_dist + inner_dist + o_v1 + 2*thickness ' 風斗板上下板寬度
    o_v3 = fin_length + 2*f_v1 + 2*thickness ' 風斗位置寬度
    o_v4 = connect_width + thickness ' 風斗板連接處寬度
    o_v5 = 8 ' TODO:
    o_v6 = 60 ' TODO:
    o_v7 = 29 ' 常更動 外端板下方連接處
    
    out_v1 = 5

    ' 一般熱排 風斗板
    comp_start(0) = start(0): comp_start(1) = start(1) + 1500
    comp_title = tube_num_row & "R" & tube_num_stick & "T" & "   " & efficient_dist &  "  m / m"
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height

    comp_length = efficient_dist + 2*o_v4
    comp_width = o_v3 + 2*o_v2 + 2*o_v4

    comp_start(0) = start(0) + 3000: comp_start(1) = start(1)
    fans_board comp_start, d1, o_v2, o_v3, o_v4, o_v5, o_v6, comp_width, d3

    comp_start(0) = comp_start(0): comp_start(1) = comp_start(1) - text_dist
    comp_title = "風斗板 " & thickness & "t " & material & "  " & comp_length & " X " & comp_width _ 
                  & " X " & "1只"
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height

    ' 一般熱排 內端板
    comp_start(0) = start(0) + 2000: comp_start(1) = start(1) + 500
    inner_side_board comp_start, d2

    comp_start(0) = comp_start(0): comp_start(1) = comp_start(1) - text_dist
    comp_title = "內端板 " & thickness & "t " & partition_material & "  " & fin_length & " X " & fin_width _ 
                  & " X " & "1只"
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height

    ' 一般熱排 隔板
    comp_length = fin_length + 2*p_v1
    comp_width = inner_dist - p_v3 + 2*p_v1

    comp_start(0) = start(0): comp_start(1) = start(1)
    partition comp_start, p_v1, p_v2, p_v3, d2, p_v4
    
    comp_start(0) = comp_start(0): comp_start(1) = comp_start(1) - text_dist
    comp_title = "隔板 " & thickness & "t " & material & "  " & comp_length & " X " & comp_width _ 
                  & " X " & num_motor-1 & "只"
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height
    
    
    
    ' 一般熱排 外端板
    comp_length = inner_dist + o_v1 + fin_width + 2*connect_width - p_v3
    comp_width = fin_length + 2*f_v1 +  connect_width + o_v7

    comp_start(0) = start(0): comp_start(1) = start(1) + 500
    comp_start2(0) = start(0) + 1000: comp_start2(1) = start(1) + 500
    outer_side_board comp_start, comp_start2, o_v7, f_v1, o_v1, p_v3, out_v1, o_v6, _
                     fin_length, fin_width, d2

    comp_start(0) = comp_start(0): comp_start(1) = comp_start(1) - text_dist
    comp_title = "外端板 " & thickness & "t " & material & "  " & comp_length & " X " & comp_width _ 
                  & " X " & "2只"
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height

    ' MsgBox "製圖完成"


End Sub

' Private Sub CommandButton1_Click()
'     '
' MsgBox "找到圖層:"
'     MsgBox material
' End Sub

Private Sub CommandButton1_Click()
    UserForm1.Hide

    t1 = Val(TextBox1.Text)
    t2 = Val(TextBox2.Text)
    t3 = Val(TextBox3.Text)
    t4 = Val(TextBox4.Text)
    t5 = Val(TextBox5.Text)
    t6 = Val(TextBox6.Text)
    t7 = Val(TextBox7.Text)
    t8 = Val(TextBox8.Text)
    t9 = Val(TextBox9.Text)
    t10 = Val(TextBox10.Text)
    t11 = Val(TextBox11.Text)
    t12 = TextBox12.Text
    t13 = TextBox13.Text
    t14 = Val(TextBox14.Text)
    t15 = Val(TextBox15.Text)
    t16 = Val(TextBox16.Text)
    t17 = Val(TextBox17.Text)
    t18 = TextBox18.Text

    tube_head = t1
    tube_diameter = t2
    tube_num_row = t3
    tube_num_stick = t4
    efficient_dist = t5
    fan_diameter = t6
    motor_frame_length = t9
    motor_frame_width = t10
    thickness = t11
    material = t12
    partition_material = t13
    num_motor = t8
    inner_dist = t7
    connect_width = t14
    num_screw = t15
    screw_dist = t16
    num_part_screw = t17
    is_expand = t18

    heater
End Sub


' Private Sub Label1_Click()

' End Sub

' Private Sub Label2_Click()

' End Sub

' Private Sub Label3_Click()

' End Sub

' Private Sub TextBox1_Change()

' End Sub

' Private Sub TextBox2_Change()

' End Sub

' Private Sub TextBox3_Change()

' End Sub

' Private Sub TextBox4_Change()

' End Sub

' Private Sub TextBox5_Change()

' End Sub

' Private Sub TextBox6_Change()

' End Sub

' Private Sub TextBox7_Change()

' End Sub

' Private Sub TextBox8_Change()

' End Sub

' Private Sub TextBox9_Change()

' End Sub

' Private Sub TextBox10_Change()

' End Sub

' Private Sub TextBox11_Change()

' End Sub

' Private Sub UserForm_Click()

' End Sub













Private Sub TextBox5_Change()

End Sub



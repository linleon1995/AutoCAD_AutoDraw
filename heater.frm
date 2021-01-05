VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Heater"
   ClientHeight    =   9780.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14628
   OleObjectBlob   =   "heater.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Public tube_head As String
Public tube_diameter As Double
Public tube_num_row As Double
Public tube_num_stick As Double
Public efficient_dist As Double
Public fan_diameter As Double
Public motor_frame_length As Double
Public motor_frame_diagonal As Double
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


Public layout_origin As Double
Public layout_range As Double
Public outer_side_board_length As Double
Public outer_side_board_width As Double
Public outer_side_board_thickness As Double
Public inner_side_board_length As Double
Public inner_side_board_width As Double
Public inner_side_board_thickness As Double
Public partition_length As Double
Public partition_width As Double
Public fans_board_length As Double
Public fans_board_width As Double
Public fans_board_thickness As Double
Public tube_hole_type As String
Public outer_side_board_in_length As String
Public num_inner_side_board As Double
Public outer_side_board_up_hole As String

Public tube_num_row_array As Variant

' Function Round_5(num As Integer, d As Integer) As Double
'     Dim new_num  As Double
'     new_num =
'     If d > 0 Then
'         num Mod (10^d)
'     Else
'     End If
' End Function


Public Sub AddRectCircles(ByVal start, ByVal radius, ByVal length, ByVal width, _
                          Optional add_cross As Boolean = False, _
                          Optional cross_length As Double = 7.5)

    Dim p(2)  As Double
    Dim cir_obj As AcadCircle
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(start, radius)
    If add_cross = True Then
        AddCross start, cross_length, cross_length, cross_length, cross_length
    End If

    p(0) = start(0) + length: p(1) = start(1)
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(p, radius)
    If add_cross = True Then
        AddCross p, cross_length, cross_length, cross_length, cross_length
    End If

    p(0) = start(0): p(1) = start(1) + width
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(p, radius)
    If add_cross = True Then
        AddCross p, cross_length, cross_length, cross_length, cross_length
    End If

    p(0) = start(0) + length: p(1) = start(1) + width
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(p, radius)
    If add_cross = True Then
        AddCross p, cross_length, cross_length, cross_length, cross_length
    End If
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
            SelectActiveLayer "尺寸"
            If text_dir = 0 Then
                text_loc(0) = c(0) + text_dist: text_loc(1) = (c(1) + c2(1)) / 2
            ElseIf text_dir = 1 Then
                text_loc(0) = (c(0) + c2(0)) / 2: text_loc(1) = c(1) + text_dist
            ElseIf text_dir = 2 Then
                text_loc(0) = c(0) - text_dist: text_loc(1) = (c(1) + c2(1)) / 2
            ElseIf text_dir = 3 Then
                text_loc(0) = (c(0) + c2(0)) / 2: text_loc(1) = c(1) - text_dist
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
                         ByVal stick_dist, ByVal diameter, Optional direction As String = "right")
    Dim start2(2)  As Double
    start2(0) = start(0) + row_dist: start2(1) = start(1) - stick_dist / 2

    For i = 1 To num_row
        If direction = "left" Then
            shift = -2 * row_dist
        ElseIf direction = "right" Then
            shift = 2 * row_dist
        End If
        If i Mod 2 = 1 Then
            AddLinedCircles start, diameter / 2, stick_dist, num_stick, 1
            start(0) = start(0) + shift
        Else:
            AddLinedCircles start2, diameter / 2, stick_dist, num_stick, 1
            start2(0) = start2(0) + shift
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
    p1(0) = center(0) - l1: p1(1) = center(1)
    p2(0) = center(0) + l2: p2(1) = center(1)
    Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
    p1(0) = center(0): p1(1) = center(1) + w2
    p2(0) = center(0): p2(1) = center(1) - w1
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

        For i = 0 To 4
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

        SelectActiveLayer "尺寸"
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
    x = (max_dist - radius * 2) / 2

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

    ' 標註
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
                            ByVal w1, ByVal w2, Optional dim_dist, Optional text_height, _
                            Optional arrow_size)
    
    Dim start_r1(2)  As Double, start_r2(2)  As Double
    Dim a(2)  As Double, b(2)  As Double, c(2)  As Double, d(2)  As Double, _
        e(2)  As Double, f(2)  As Double, g(2)  As Double, h(2)  As Double, i(2)  As Double
    Dim t1(2)  As Double, t2(2)  As Double, t3(2)  As Double, t4(2)  As Double, _
        t5(2)  As Double, t6(2)  As Double, t7(2)  As Double, t8(2)  As Double
    start_r1(0) = start(0): start_r1(1) = start(1) + w1
    start_r2(0) = start(0) + l1: start_r2(1) = start(1)
    in_length = length - l1 - l2
    in_width = width - w1 - w2
    
    a(0) = start_r2(0) + in_length: a(1) = start_r2(1)
    b(0) = start_r2(0) + in_length: b(1) = start_r1(1)
    c(0) = start_r2(0) + in_length + l2: c(1) = start_r1(1)
    d(0) = start_r1(0) + length: d(1) = start_r1(1) + in_width
    e(0) = start_r2(0) + in_length: e(1) = start_r1(1) + in_width
    f(0) = start_r2(0) + in_length: f(1) = start_r2(1) + width
    g(0) = start_r1(0) + l1: g(1) = start_r2(1) + width
    h(0) = start_r1(0) + l1: h(1) = start_r1(1) + in_width
    i(0) = start_r1(0): i(1) = start_r1(1) + in_width

    ' t1(0) = start_r1(0) + 0.5 * length: t1(1) = start_r1(1) - dim_dist - w1
    ' t2(0) = start_r2(0) - dim_dist - l1: t2(1) = start_r2(1) + 0.5 * width
    t3(0) = start_r1(0) + length + dim_dist: t3(1) = start_r2(1) + 0.5 * w1
    t4(0) = start_r1(0) + length + dim_dist: t4(1) = start_r2(1) + w1 + 0.5 * in_width
    t5(0) = start_r1(0) + length + dim_dist: t5(1) = start_r2(1) + w1 + in_width + 0.5 * w2
    t6(0) = start_r1(0) + 0.5 * l1 + in_length + 0.5 * l2: t6(1) = start_r2(1) + width + dim_dist
    t7(0) = start_r1(0) + l1 + 0.5 * in_length: t7(1) = start_r2(1) + width + dim_dist
    t8(0) = start_r1(0) + 0.5 * l1 - 50: t8(1) = start_r2(1) + width + dim_dist

    SelectActiveLayer "鈑金"
    AddRect start_r1, length, in_width
    AddRect start_r2, in_length, width

    SelectActiveLayer "尺寸"
    AddMainDims start_r1, c, dim_dist, "down", start_r2, g, dim_dist, "left", text_height, arrow_size1
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


Public Sub fans_board(ByVal start, ByVal comp_length, ByVal comp_width, ByVal d1, ByVal o_v2, ByVal o_v3, _
                      ByVal o_v4, ByVal o_v6, ByVal d3, ByVal thickness)
    Dim line_obj As AcadLine
    Dim p1(2) As Double
    Dim p2(2) As Double
    Dim p3(2) As Double
    Dim p4(2) As Double
    Dim p5(2) As Double
    Dim add_dim(4) As Double
    Dim text_loc(2) As Double
    Dim chordPoint(2) As Double, FarchordPoint(2) As Double
    ' comp_length = efficient_dist + 2*o_v4

    ' 鈑金
    ' TODO: complete AddHill dimension function
    ' add_dim(0) = 1: add_dim(1) = 2: add_dim(2) = 3: add_dim(3) = 4: add_dim(4) = 5
    AddHill start, 0, comp_length, 0, o_v4, o_v4, "h_flip", text_height:=dim_text_height1, arrow_size:=arrow_size1
    p1(0) = start(0): p1(1) = start(1) + o_v2
    p3(0) = p1(0): p3(1) = p1(1)
    AddHill p1, 0, comp_length, 0, o_v2, o_v2, "h_flip", text_height:=dim_text_height1, arrow_size:=arrow_size1
    p1(0) = p1(0): p1(1) = p1(1) + o_v3
    AddHill p1, 0, comp_length, 0, o_v3, o_v3, "h_flip", text_height:=dim_text_height1, arrow_size:=arrow_size1
    p1(0) = p1(0): p1(1) = p1(1) + o_v2
    AddHill p1, 0, comp_length, 0, o_v2, o_v2, "h_flip", text_height:=dim_text_height1, arrow_size:=arrow_size1
    p1(0) = p1(0): p1(1) = p1(1) + o_v4
    AddHill p1, 0, comp_length, 0, o_v4, o_v4, "h_flip", text_height:=dim_text_height1, arrow_size:=arrow_size1
    p2(0) = p1(0) + comp_length: p2(1) = p1(1)
    Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)

    p3(0) = p3(0) + dim_dist1 + 100 + comp_length + o_v2: p3(1) = p3(1) + o_v4
    AddHill p3, -1 * o_v4, o_v3, -1 * o_v4, o_v2, o_v2, "v_flip"
    p4(0) = p3(0) - o_v2: p4(1) = p3(1) - o_v4 + (o_v3 - fan_diameter) / 2
    AddHill p4, 0, fan_diameter, 0, 8, 8, "v_flip"

    SelectActiveLayer "尺寸"
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
    p1(0) = start(0) + Val(Format((connect_width + outer_side_board_thickness) / 2, ".0"))
    p1(1) = start(1) - o_v4 / 2
    AddRectCircles p1, d1 / 2, comp_length - o_v4, comp_width - o_v4

    SelectActiveLayer "尺寸"
    p2(0) = p1(0): p2(1) = p1(1) - o_v4 / 2
    text_loc(0) = p1(0) - dim_dist2: text_loc(1) = (p1(1) + p2(1)) / 2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2


    x = o_v6 + thickness + (o_v3 - screw_dist) / 2
    If outer_side_board_up_hole = "是" Then
        SelectActiveLayer "鈑金"
        p1(0) = start(0) + Val(Format((connect_width + outer_side_board_thickness) / 2, ".0"))
        p1(1) = start(1) + o_v2 + o_v6 + thickness + o_v3
        ' AddRectCircles p1, d1/2, comp_length-o_v4, o_v3+2*(thickness+o_v6)

        AddLinedCircles p1, d1 / 2, comp_length - o_v4, 2, 0

        SelectActiveLayer "尺寸"
        ' ' ' 下折面 距離標註
        ' p2(0) = start(0) + Val(Format((connect_width+outer_side_board_thickness)/2, ".0"))
        ' p2(1) = start(1) + o_v2
        ' text_loc(0) = p1(0) - dim_dist2: text_loc(1) = (p1(1) + p2(1)) / 2
        ' Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, text_loc)
        ' AcadDimAligned.TextHeight = dim_text_height2
        ' AcadDimAligned.ArrowheadSize = arrow_size2

        ' 上折面 距離標註
        p2(0) = p1(0)
        p2(1) = p1(1) - thickness - o_v6
        ' p1(1) = p1(1) + o_v6 + thickness + o_v3
        
        text_loc(0) = p1(0) - dim_dist2: text_loc(1) = (p1(1) + p2(1)) / 2
        Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, text_loc)
        AcadDimAligned.TextHeight = dim_text_height2
        AcadDimAligned.ArrowheadSize = arrow_size2

        ' 上折面螺絲孔 孔徑標註
        chordPoint(0) = p2(0) + d1 / 2 / Sqr(2): chordPoint(1) = p2(1) + d1 / 2 / Sqr(2)
        FarchordPoint(0) = p2(0) - d1 / 2 / Sqr(2): FarchordPoint(1) = p2(1) - d1 / 2 / Sqr(2)
        Set dim_obj = ThisDrawing.ModelSpace.AddDimDiametric(chordPoint, FarchordPoint, LeaderLength:=5)
        dim_obj.TextHeight = dim_text_height2
        dim_obj.ArrowheadSize = arrow_size2
    End If

    ' 風斗左側邊螺絲孔
    SelectActiveLayer "鈑金"
    p1(0) = start(0) + Val(Format((connect_width + outer_side_board_thickness) / 2, ".0"))
    p1(1) = start(1) + o_v2 - thickness - o_v6 + x
    AddLinedCircles p1, d1 / 2, screw_dist / (num_screw - 1), num_screw, 1, _
                    text_dist:=dim_dist2, text_dir:=2, text_height:=dim_text_height2, arrow_size1:=arrow_size1

    ' 風斗左側邊螺絲孔 標註
    SelectActiveLayer "尺寸"
    
    p2(0) = p1(0) - Val(Format((connect_width + outer_side_board_thickness) / 2, ".0")): p2(1) = p1(1)
    text_loc(0) = (p1(0) + p2(0)) / 2: text_loc(1) = p1(1) - dim_dist3
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2

    ' 風斗右側邊螺絲孔
    SelectActiveLayer "鈑金"
    p1(0) = p1(0) + comp_length - 2 * Val(Format((connect_width + outer_side_board_thickness) / 2, ".0"))
    AddLinedCircles p1, d1 / 2, screw_dist / (num_screw - 1), num_screw, 1

    ' 風斗與馬達架螺絲孔
    ' TODO: check correctness
    ' comp = 4 - (Int(efficient_dist/num_motor) Mod 4)
    ' dist = Int(efficient_dist/num_motor) + comp
    dist = Int(efficient_dist / num_motor)
    p1(0) = start(0) + comp_length / 2: p1(1) = start(1) + comp_width / 2 - o_v4
    p1(0) = p1(0) - (0.5 * num_motor - 0.5) * dist
    AddLinedCircles p1, fan_diameter / 2, dist, num_motor, 0

    ' 螺絲定位
    p2(0) = p1(0) - motor_frame_length / 2: p2(1) = p1(1) - motor_frame_length / 2
    ' 文字定位
    p3(0) = p1(0) - 70: p3(1) = p1(1) - 70
    p5(0) = p1(0) - 70: p5(1) = p1(1)
    For i = 1 To num_motor
        ' 馬達架螺絲孔
        SelectActiveLayer "螺絲"
        AddRectCircles p2, 11.5 / 2, motor_frame_length, motor_frame_length, add_cross:=True

        ' 風扇文字說明
        SelectActiveLayer "鈑金"
        If fan_type = -1 Then
            ThisDrawing.ModelSpace.AddText "Φ" & fan_diameter, InsertionPoint:=p5, Height:=50
        Else
            ThisDrawing.ModelSpace.AddText fan_type & "''", InsertionPoint:=p5, Height:=50
        End If
        ThisDrawing.ModelSpace.AddText is_expand, InsertionPoint:=p3, Height:=50

        ' 隔板螺絲孔
        If i < num_motor Then
            SelectActiveLayer "鈑金"
            p4(0) = p1(0) + dist / 2: p4(1) = p1(1) - screw_dist / 2
            If i = 1 Then
                AddLinedCircles p4, d1 / 2, screw_dist / (num_part_screw - 1), num_part_screw, 1, _
                                text_dist:=10, text_dir:=2, text_height:=dim_text_height2, arrow_size1:=arrow_size2
            Else
                AddLinedCircles p4, d1 / 2, screw_dist / (num_part_screw - 1), num_part_screw, 1
            End If
        End If

        p1(0) = p1(0) + dist
        p2(0) = p2(0) + dist
        p3(0) = p3(0) + dist
        p5(0) = p5(0) + dist
    Next

    ' 馬達架螺絲孔尺寸
    ' 馬達架螺絲孔對角線尺寸標註
    SelectActiveLayer "尺寸"
    p2(0) = p2(0) - dist
    p1(0) = p2(0) + motor_frame_length: p1(1) = p2(1) + motor_frame_length
    text_loc(0) = (p1(0) + p2(0)) / 2: text_loc(1) = p1(1) + dim_dist2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2

    ' 馬達架螺絲孔長寬尺寸標註
    ' p2(0) = p2(0) - dist
    ' p1(0) = p2(0) + motor_frame_length: p1(1) = p2(1)
    ' text_loc(0) = (p1(0) + p2(0)) / 2: text_loc(1) = p1(1) - dim_dist2
    ' Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, text_loc)
    ' AcadDimAligned.TextHeight = dim_text_height2
    ' AcadDimAligned.ArrowheadSize = arrow_size2
    ' p1(0) = p2(0): p1(1) = p2(1) + motor_frame_length
    ' text_loc(0) = p1(0) - dim_dist2: text_loc(1) = (p1(1) + p2(1)) / 2
    ' Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, text_loc)
    ' AcadDimAligned.TextHeight = dim_text_height2
    ' AcadDimAligned.ArrowheadSize = arrow_size2

    SelectActiveLayer "鈑金"
End Sub


Public Sub inner_side_board(ByVal start, ByVal comp_length, ByVal comp_width, ByVal d)
    Dim p(2) As Double, p2(2) As Double, p3(2) As Double, p4(2) As Double
    Dim text_loc(2) As Double
    Dim chordPoint(2) As Double, FarchordPoint(2) As Double

    SelectActiveLayer "螺絲"
    p(0) = start(0): p(1) = start(1) + 1.5 * stick_dist
    AddFinCircles p, tube_num_row, 1, row_dist, stick_dist, d
    p(0) = start(0): p(1) = start(1) + (tube_num_stick - 1 - 1.5) * stick_dist
    AddFinCircles p, tube_num_row, 1, row_dist, stick_dist, d

    SelectActiveLayer "鈑金"
    ' 銅管孔
    AddFinCircles start, tube_num_row, tube_num_stick, row_dist, stick_dist, in_sb_d
    p(0) = start(0) - 0.5 * row_dist: p(1) = start(1) - 0.75 * stick_dist

    ' 外框
    AddRect p, comp_length, comp_width

    ' 分割線
    num_split = UBound(tube_num_row_array) ' 分割數
    If num_split > 0 Then ' 輸入一個數值以上，代表需要繪製分割線
        p4(0) = p(0): p4(1) = p(1) + inner_side_board_width
        p2(0) = p(0): p2(1) = p(1)
        For i = 0 To num_split - 1
            p2(0) = p2(0) + tube_num_row_array(i) * row_dist
            p3(0) = p2(0): p3(1) = p2(1) + inner_side_board_width
            
            SelectActiveLayer "尺寸"
            text_loc(0) = (p3(0) + p4(0)) / 2: text_loc(1) = p3(1) + dim_dist1
            Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p3, p4, text_loc)
            AcadDimAligned.TextHeight = dim_text_height1
            AcadDimAligned.ArrowheadSize = arrow_size1
            p4(0) = p3(0): p4(1) = p3(1)

            SelectActiveLayer "鈑金"
            Set line_obj = ThisDrawing.ModelSpace.AddLine(p2, p3)
        Next
        SelectActiveLayer "尺寸"

        p4(0) = p3(0) + tube_num_row_array(num_split) * row_dist
        text_loc(0) = (p3(0) + p4(0)) / 2: text_loc(1) = p3(1) + dim_dist1
        Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p3, p4, text_loc)
        AcadDimAligned.TextHeight = dim_text_height1
        AcadDimAligned.ArrowheadSize = arrow_size1
    End If

    ' 外框尺寸標註
    SelectActiveLayer "尺寸"
    p2(0) = p(0) + inner_side_board_length: p2(1) = p(1)
    text_loc(0) = p(0) + 0.5 * inner_side_board_length: text_loc(1) = p(1) - dim_dist1
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p, p2, text_loc)
    AcadDimAligned.TextHeight = dim_text_height1
    AcadDimAligned.ArrowheadSize = arrow_size1
    p2(0) = p(0): p2(1) = p(1) + inner_side_board_width
    text_loc(0) = p(0) - dim_dist1: text_loc(1) = p(1) + 0.5 * inner_side_board_width
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p, p2, text_loc)
    AcadDimAligned.TextHeight = dim_text_height1
    AcadDimAligned.ArrowheadSize = arrow_size1

    ' 銅管孔尺寸標註
    ' p2(0) = start(0): p2(1) = start(1) + (tube_num_stick-1-1.5)*stick_dist
    If tube_num_row Mod 2 = 0 Then
        comp = 0.5
    Else
        comp = 0
    End If
        
    p2(0) = start(0) + (tube_num_row - 1) * row_dist
    p2(1) = start(1) + Int(tube_num_stick / 2) * stick_dist - comp * stick_dist
    chordPoint(0) = p2(0) + in_sb_d / 2 / Sqr(2): chordPoint(1) = p2(1) - in_sb_d / 2 / Sqr(2)
    FarchordPoint(0) = p2(0) - in_sb_d / 2 / Sqr(2): FarchordPoint(1) = p2(1) + in_sb_d / 2 / Sqr(2)
    Set dim_obj = ThisDrawing.ModelSpace.AddDimDiametric(chordPoint, FarchordPoint, LeaderLength:=5)
    dim_obj.TextHeight = dim_text_height2
    dim_obj.ArrowheadSize = arrow_size2

    p2(0) = start(0) + (tube_num_row - 1) * row_dist
    p2(1) = start(1) + (tube_num_stick - 1 - 1.5) * stick_dist - comp * stick_dist
    chordPoint(0) = p2(0) + d / 2 / Sqr(2): chordPoint(1) = p2(1) - d / 2 / Sqr(2)
    FarchordPoint(0) = p2(0) - d / 2 / Sqr(2): FarchordPoint(1) = p2(1) + d / 2 / Sqr(2)
    Set dim_obj = ThisDrawing.ModelSpace.AddDimDiametric(chordPoint, FarchordPoint, LeaderLength:=5)
    dim_obj.TextHeight = dim_text_height2
    dim_obj.ArrowheadSize = arrow_size2

    SelectActiveLayer "鈑金"
End Sub


Public Sub partition(ByVal start, ByVal comp_length, ByVal comp_width, ByVal v1, ByVal v2, _
                     ByVal v3, ByVal d1, ByVal dist1)
    Dim p1(2) As Double: Dim p2(2) As Double
    Dim text_loc(2) As Double
    Dim center(2) As Double
    Dim chordPoint(2) As Double, FarchordPoint(2) As Double
    Dim leaderLen As Integer

    ' 鈑金
    AddTwoCrossRects start, comp_length, comp_width, v1, v1, v1, v1, _
                     dim_dist:=dim_dist1, text_height:=dim_text_height1, arrow_size:=arrow_size

    ' 螺絲
    x = (comp_length - 2 * v1 - screw_dist) / 2
    p1(0) = start(0) + v1 + x: p1(1) = start(1) + comp_width - dist1
    AddLinedCircles p1, d1 / 2, screw_dist / (num_part_screw - 1), num_part_screw, 0, _
                    text_dist:=dim_dist2, text_dir:=1, text_height:=dim_text_height2, arrow_size1:=arrow_size2

    ' 螺絲孔標註
    SelectActiveLayer "尺寸"
    FarchordPoint(0) = p1(0) - d1 / 2 / Sqr(2): FarchordPoint(1) = p1(1) + d1 / 2 / Sqr(2)
    chordPoint(0) = p1(0) + d1 / 2 / Sqr(2): chordPoint(1) = p1(1) - d1 / 2 / Sqr(2)
    Set dim_obj = ThisDrawing.ModelSpace.AddDimDiametric(chordPoint, FarchordPoint, LeaderLength:=5)
    dim_obj.TextHeight = dim_text_height2
    dim_obj.ArrowheadSize = arrow_size2

    ' x = (fin_length - screw_dist) / 2
    p2(0) = p1(0) - x: p2(1) = p1(1)
    text_loc(0) = (p1(0) + p2(0)) / 2: text_loc(1) = p1(1) + dim_dist2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2

    p1(0) = p1(0) + comp_length - 2 * v1 - x
    p2(0) = p2(0) + comp_length - 2 * v1 - x
    text_loc(0) = (p1(0) + p2(0)) / 2: text_loc(1) = p1(1) + dim_dist2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2
    SelectActiveLayer "鈑金"
End Sub


Public Sub outer_side_board(ByVal start2, ByVal comp_length, ByVal comp_width, ByVal v1, _
                            ByVal v2, ByVal v3, ByVal v4, ByVal v5, ByVal v6, ByVal fin_length, _
                            ByVal fin_width, ByVal d1, ByVal thickness, ByVal d2)
    Dim p1(2) As Double, p2(2) As Double, p3(2) As Double
    Dim reflect_point(2) As Double
    Dim line_obj As AcadLine
    Dim text_loc(2) As Double
    Dim chordPoint(2) As Double, FarchordPoint(2) As Double

    ext = (comp_width - v1 - connect_width - fin_length) / 2
    
    
    If tube_hole_type = "橢圓孔" Then
        middle = start2(0) - layout_range / 2 ' 鏡射中線
    End If

    ' 外框
    ' 焊面
    AddTwoCrossRects start2, comp_length, comp_width, connect_width, connect_width, _
                     v1, connect_width, dim_dist:=dim_dist1, text_height:=dim_text_height1, _
                     arrow_size:=arrow_size
 
    ' 穿面
    If tube_hole_type = "橢圓孔" Then
        reflect_point(0) = start2(0) - 2 * (start2(0) - middle) - comp_length: reflect_point(1) = start2(1)
        AddTwoCrossRects reflect_point, comp_length, comp_width, connect_width, connect_width, _
                        v1, connect_width, dim_dist:=dim_dist1, text_height:=dim_text_height1, _
                        arrow_size:=arrow_size
    End If
    

    ' 左側連接處螺絲孔
    ' 焊面
    p1(0) = start2(0) + Val(Format((connect_width + thickness) / 2, ".0"))
    p1(1) = start2(1) + v1 + (comp_width - v1 - connect_width - screw_dist) / 2
    AddLinedCircles p1, d1 / 2, screw_dist / (num_screw - 1), num_screw, 1, _
                    text_dist:=dim_dist2, text_dir:=2, text_height:=dim_text_height2, arrow_size1:=arrow_size2
    
    ' 穿面
    If tube_hole_type = "橢圓孔" Then
        reflect_point(0) = p1(0) - 2 * (p1(0) - middle): reflect_point(1) = p1(1)
        AddLinedCircles reflect_point, d1 / 2, screw_dist / (num_screw - 1), num_screw, 1
    End If

    SelectActiveLayer "尺寸"
    chordPoint(0) = p1(0) + d1 / 2 / Sqr(2): chordPoint(1) = p1(1) + d1 / 2 / Sqr(2)
    FarchordPoint(0) = p1(0) - d1 / 2 / Sqr(2): FarchordPoint(1) = p1(1) - d1 / 2 / Sqr(2)
    Set dim_obj = ThisDrawing.ModelSpace.AddDimDiametric(chordPoint, FarchordPoint, LeaderLength:=5)
    dim_obj.TextHeight = dim_text_height2
    dim_obj.ArrowheadSize = arrow_size2

    p1(1) = p1(1)
    p3(0) = p1(0): p3(1) = p1(1) - (outer_side_board_width - screw_dist - v1 - connect_width) / 2
    text_loc(0) = p1(0) - dim_dist2: text_loc(1) = (p1(1) + p3(1)) / 2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p3, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2

    p1(1) = p1(1) + screw_dist
    p3(0) = start2(0): p3(1) = p1(1)
    text_loc(0) = (p1(0) + p3(0)) / 2: text_loc(1) = p1(1) + dim_dist2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p3, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2

    p3(0) = p1(0): p3(1) = p1(1) + (outer_side_board_width - screw_dist - v1 - connect_width) / 2
    text_loc(0) = p1(0) - dim_dist2: text_loc(1) = (p1(1) + p3(1)) / 2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p3, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2

    ' 右側連接處螺絲孔
    ' 焊面
    SelectActiveLayer "鈑金"
    p1(0) = start2(0) + comp_length - connect_width / 2: p1(1) = start2(1) + v1 + v5
    AddLinedCircles p1, d1 / 2, comp_width - v1 - connect_width - 2 * v5, 2, 1

    ' 穿面
    If tube_hole_type = "橢圓孔" Then
        reflect_point(0) = p1(0) - 2 * (p1(0) - middle): reflect_point(1) = p1(1)
        AddLinedCircles reflect_point, d1 / 2, comp_width - v1 - connect_width - 2 * v5, 2, 1
    End If
    
    SelectActiveLayer "尺寸"
    p3(0) = p1(0): p3(1) = p1(1) - v5
    text_loc(0) = p1(0) + dim_dist2: text_loc(1) = (p1(1) + p3(1)) / 2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p3, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2
    
    p1(1) = start2(1) + comp_width - v5 - connect_width
    p3(0) = p1(0): p3(1) = p1(1) + v5
    text_loc(0) = p1(0) + dim_dist2: text_loc(1) = (p1(1) + p3(1)) / 2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p3, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2

    ' 上連接處螺絲孔
    If outer_side_board_up_hole = "是" Then
        SelectActiveLayer "鈑金"
        ' 焊面
        p1(0) = start2(0) + connect_width + v6: p1(1) = start2(1) + comp_width - connect_width / 2
        AddLinedCircles p1, d1 / 2, 0, 1, 0, _
                        text_dist:=dim_dist2, text_dir:=2, text_height:=dim_text_height1, arrow_size1:=arrow_size1

        ' 穿面
        If tube_hole_type = "橢圓孔" Then
            reflect_point(0) = p1(0) - 2 * (p1(0) - middle): reflect_point(1) = p1(1)
            AddLinedCircles reflect_point, d1 / 2, 0, 1, 0
        End If

        SelectActiveLayer "尺寸"
        p3(0) = p1(0) - v6: p3(1) = p1(1)
        text_loc(0) = (p1(0) + p3(0)) / 2: text_loc(1) = p1(1) + dim_dist2
        Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p3, text_loc)
        AcadDimAligned.TextHeight = dim_text_height2
        AcadDimAligned.ArrowheadSize = arrow_size2
    End If

    ' 下方連接處螺絲孔
    SelectActiveLayer "鈑金"
    ' TODO: 20 12 8
    ' AMADA 橢圓刀孔徑 Phi 8 X 12
    max_d = 12
    d2 = 8.45
    x1 = 20
    x2 = 8
    x3 = 10
    ' 焊面
    ' p1(0) = start2(0) +  connect_width + x1: p1(1) = start2(1) + x2
    ' AddArcwithLines p1, d2/2, max_d
    ' p3(0) = p1(0) + comp_length - 2*connect_width - x1 - x3: p3(1) = p1(1)
    ' AddArcwithLines p3, d2/2, max_d
    
    p1(0) = start2(0) + connect_width + x1: p1(1) = start2(1) + x2
    AddLinedCircles start:=p1, radius:=d2 / 2, distance:=outer_side_board_in_length - x1 - x3, num_circle:=2, direction:=0, _
                    text_dist:=dim_dist2, text_dir:=3, text_height:=dim_text_height2, _
                    arrow_size1:=arrow_size1
    p3(0) = p1(0) + outer_side_board_in_length - x1 - x3: p3(1) = p1(1)

    ' 下方連接處螺絲孔標註
    SelectActiveLayer "尺寸"
    p2(0) = p1(0) - x1: p2(1) = p1(1)
    text_loc(0) = (p1(0) + p2(0)) / 2: text_loc(1) = p1(1) - dim_dist2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p2, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2

    FarchordPoint(0) = p1(0) - d2 / 2 / Sqr(2): FarchordPoint(1) = p1(1) - d2 / 2 / Sqr(2)
    chordPoint(0) = p1(0) + d2 / 2 / Sqr(2): chordPoint(1) = p1(1) + d2 / 2 / Sqr(2)
    Set dim_obj = ThisDrawing.ModelSpace.AddDimDiametric(chordPoint, FarchordPoint, LeaderLength:=5)
    dim_obj.TextHeight = dim_text_height2
    dim_obj.ArrowheadSize = arrow_size2

    p2(0) = p3(0) + x3: p2(1) = p3(1)
    text_loc(0) = (p3(0) + p2(0)) / 2: text_loc(1) = p3(1) - dim_dist2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p3, p2, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2

    p2(0) = p3(0): p2(1) = p3(1) - x2
    text_loc(0) = p3(0) - dim_dist3: text_loc(1) = (p3(1) + p2(1)) / 2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p3, p2, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2

    SelectActiveLayer "鈑金"
    
    ' 穿面
    If tube_hole_type = "橢圓孔" Then
        reflect_point(0) = p3(0) - 2 * (p3(0) - middle): reflect_point(1) = p3(1)
        AddLinedCircles start:=reflect_point, radius:=d2 / 2, distance:=outer_side_board_in_length - x1 - x3, _
                        num_circle:=2, direction:=0


        ' reflect_point(0) = p1(0) - 2*(p1(0)-middle): reflect_point(1) = p1(1)
        ' AddArcwithLines reflect_point, d2/2, max_d
        ' reflect_point(0) = p3(0) - 2*(p3(0)-middle): reflect_point(1) = p3(1)
        ' AddArcwithLines reflect_point, d2/2, max_d
    End If


    ' 銅管孔
    p1(0) = start2(0) + connect_width + inner_dist: p1(1) = start2(1) + v1 + ext + stick_dist * 3 / 4
    AddFinCircles p1, tube_num_row, tube_num_stick, row_dist, stick_dist, out_sb_out_d

    If num_inner_side_board > 0 Then
        SelectActiveLayer "螺絲"
        p2(0) = p1(0): p2(1) = p1(1) + 1.5 * stick_dist
        AddFinCircles p2, tube_num_row, 1, row_dist, stick_dist, d1
        p2(0) = p1(0): p2(1) = p1(1) + (tube_num_stick - 1 - 1.5) * stick_dist
        AddFinCircles p2, tube_num_row, 1, row_dist, stick_dist, d1
    End If

    SelectActiveLayer "尺寸"
    chordPoint(0) = p2(0) - d1 / 2 / Sqr(2): chordPoint(1) = p2(1) - d1 / 2 / Sqr(2)
    FarchordPoint(0) = p2(0) + d1 / 2 / Sqr(2): FarchordPoint(1) = p2(1) + d1 / 2 / Sqr(2)
    Set dim_obj = ThisDrawing.ModelSpace.AddDimDiametric(chordPoint, FarchordPoint, LeaderLength:=5)
    dim_obj.TextHeight = dim_text_height2
    dim_obj.ArrowheadSize = arrow_size2

    p2(0) = p1(0): p2(1) = p1(1) + Int(tube_num_stick / 2) * stick_dist
    chordPoint(0) = p2(0) - out_sb_out_d / 2 / Sqr(2): chordPoint(1) = p2(1) - out_sb_out_d / 2 / Sqr(2)
    FarchordPoint(0) = p2(0) + out_sb_out_d / 2 / Sqr(2): FarchordPoint(1) = p2(1) + out_sb_out_d / 2 / Sqr(2)
    Set dim_obj = ThisDrawing.ModelSpace.AddDimDiametric(chordPoint, FarchordPoint, LeaderLength:=5)
    dim_obj.TextHeight = dim_text_height2
    dim_obj.ArrowheadSize = arrow_size2
    
    p1(0) = p1(0) + (tube_num_row - 1) * row_dist: p1(1) = p1(1) + (tube_num_stick - 1) * stick_dist
    p3(0) = p1(0) + v3: p3(1) = p1(1)
    text_loc(0) = (p1(0) + p3(0)) / 2: text_loc(1) = p1(1) + dim_dist2 + 0.25 * stick_dist + connect_width / 2
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, p3, text_loc)
    AcadDimAligned.TextHeight = dim_text_height2
    AcadDimAligned.ArrowheadSize = arrow_size2

    SelectActiveLayer "鈑金"
    
End Sub


Public Sub heater() ' 一般熱排
    Set lay1 = ThisDrawing.Layers.Add("鈑金") ' 增加一個名為“鈑金”的圖層
    Set lay2 = ThisDrawing.Layers.Add("尺寸") ' 增加一個名為“尺寸”的圖層
    lay2.color = 3 ' 圖層設置為綠色
    Set lay3 = ThisDrawing.Layers.Add("螺絲") ' 增加一個名為“螺絲”的圖層
    lay3.color = 6 ' 圖層設置為洋紅
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

    layout_origin = 1200
    layout_range = 500
    start(0) = layout_origin: start(1) = layout_origin
    
    ' TODO: 輸入不合法 報錯
    ' TODO: 管徑 下拉式選單
    
    ' 15HP 5/16'' 7R 24T 1350 P19.05
    tube_diameter = 2.5
    ' tube_num_row = "7"
    tube_num_stick = 24
    efficient_dist = 1350
    material = "錏板"

    fans_board_thickness = 0.8
    fan_diameter = 389
    num_motor = 3
    motor_frame_diagonal = 485.08
    num_screw = 5
    screw_dist = 520
    num_part_screw = 3
    is_expand = "抽唇"

    outer_side_board_in_length = 234
    connect_width = 13.5
    out_sb_out_d = 13.2
    tube_hole_type = "橢圓孔"
    out_sb_in_d = 13.2
    outer_side_board_thickness = 1.6
    inner_side_board_thickness = 2#


    ' ' 6HP 5/16'' 5R 16T 1140 P19.05
    ' tube_diameter = 2.5
    ' tube_num_row = 5
    ' tube_num_stick = 16
    ' efficient_dist = 1140
    ' material = "錏板"

    ' fans_board_thickness = 1.0
    ' fan_diameter = 295
    ' num_motor = 3
    ' motor_frame_diagonal = 385
    ' num_screw = 4
    ' screw_dist = 333
    ' num_part_screw = 3
    ' is_expand = "抽唇"

    ' outer_side_board_in_length = 171
    ' connect_width = 11
    ' out_sb_out_d = 8.45
    ' tube_hole_type = "圓孔"
    ' out_sb_in_d = 8.45
    ' outer_side_board_thickness = 1.0
    ' inner_side_board_thickness = 0.8


    ' ' 10HP 3/8'' 5R 20T 1454 P22
    ' tube_diameter = 3
    ' tube_num_row = 5
    ' tube_num_stick = 20
    ' efficient_dist = 1454
    ' material = "錏板"

    ' fans_board_thickness = 1.0
    ' fan_diameter = 295
    ' num_motor = 4
    ' motor_frame_diagonal = 385
    ' num_screw = 4
    ' screw_dist = 405
    ' num_part_screw = 3
    ' is_expand = "抽唇"

    ' outer_side_board_in_length = 210
    ' connect_width = 14
    ' out_sb_out_d = 14
    ' tube_hole_type = "橢圓孔"
    ' out_sb_in_d = 13.2
    ' outer_side_board_thickness = 1.0
    ' inner_side_board_thickness = 2.0

    ' tube_num_row_array = Split("3/2", "/")

    
    ' ' ' 2HP 5/16'' 5R 15T 390 P19.05
    ' tube_diameter = 2.5
    ' tube_num_row = 5
    ' tube_num_stick = 15
    ' efficient_dist = 390
    ' material = "錏板"

    ' fans_board_thickness = 0.8
    ' fan_diameter = 295
    ' num_motor = 1
    ' motor_frame_diagonal = 385
    ' num_screw = 4
    ' screw_dist = 315
    ' num_part_screw = 30
    ' is_expand = "抽唇"

    ' outer_side_board_in_length = 176
    ' connect_width = 11
    ' out_sb_out_d = 8.45
    ' tube_hole_type = "圓孔"
    ' ' out_sb_in_d =
    ' outer_side_board_thickness = 1.0
    ' inner_side_board_thickness = 20


    ' TODO: 考慮以0.5為級距去表示板厚 e.g., 0.8->1.0, 1.6->1.5
    ' real_thickness = thickness
    ' scale = Int(thickness/0.5+1)
    ' thickness = scale * 0.5

    ' 排支數換算鰭片長寬 & 穿管孔徑
    If tube_diameter = 2.5 Then
        row_dist = 19.05
        stick_dist = 25.4
        in_sb_d = 8.45
    ElseIf tube_diameter = 3 Then
        in_sb_d = 10.2
        If tube_head = "P22" Then
            row_dist = 22
            stick_dist = 25.4
        ElseIf tube_head = "P19.05" Then
            row_dist = 19.05
            stick_dist = 25.4
        End If
    ElseIf tube_diameter = 4 Then
        in_sb_d = 13.2
        row_dist = 33
        stick_dist = 38.1
    ElseIf tube_diameter = 5 Then
        in_sb_d = 16.5
        row_dist = 33
        stick_dist = 38.1
    Else
        MsgBox "           警告: " & Chr(10) & Label2.Caption & ": " & tube_diameter
    End If

    fin_length = tube_num_stick * stick_dist
    fin_width = tube_num_row * row_dist

    ' 風扇孔徑換算風扇大小
    is_expand = "抽唇"
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
        is_expand = "不抽唇"
    End If
    d1 = 5
    d2 = 3.2
    d3 = 11.5
    d4 = 8.45

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
    f_v1 = 2 ' 伸縮
    f_v2 = 1 ' 伸縮
    ' 隔板常數
    p_v1 = 11  ' 連接處寬
    p_v2 = 3
    p_v3 = 20 ' 非鰭片預留空間
    p_v4 = 6  ' 螺絲孔與邊緣距離
    ' 外端板常數
    o_v1 = 21
    inner_dist = outer_side_board_in_length - row_dist * (tube_num_row - 1) - o_v1
    o_v2 = outer_side_board_in_length + 2 * outer_side_board_thickness ' 風斗板上下板寬度
    o_v2 = Int(Format(o_v2, "."))


    ' 風斗板連接處寬度
    If efficient_dist < 1000 Then
        o_v4 = 12
    ElseIf efficient_dist >= 1000 And efficient_dist < 1500 Then
        o_v4 = 15
    ElseIf efficient_dist > 1500 Then
        o_v4 = 20
    End If
    o_v6 = 60  ' 外端板上方連結處螺絲孔
    o_v7 = 29 ' 常更動 外端板下方連接處 設為固定29
    o_v8 = o_v4 / 2 - outer_side_board_thickness
    
    partition_material = "鋁平板"
    power = tube_num_row * tube_num_stick * efficient_dist / 4 / 11 / 330
    power_decimal = power * 100
    motor_frame_length = motor_frame_diagonal / Sqr(2)
    

    ' 銅管管徑表示法轉換
    If tube_diameter = Int(tube_diameter) Then
        tube_diameter_r = tube_diameter & "/" & 8 & "''"
    Else
        tube_diameter_r = tube_diameter * 2 & "/" & 16 & "''"
    End If

    inner_side_board_length = fin_width
    inner_side_board_width = fin_length
    outer_side_board_length = outer_side_board_in_length + 2 * connect_width

    outer_side_board_width = Int(Format(fin_length + 2 * f_v1, ".")) + connect_width + o_v7

    o_v3 = outer_side_board_width + 2 * f_v2 + 2 * outer_side_board_thickness - o_v7 - connect_width
    o_v3 = Int(Format(o_v3, ".")) ' 風斗位置寬度

    fans_board_length = efficient_dist + 2 * connect_width + 2 * outer_side_board_thickness
    fans_board_length = Int(Format(fans_board_length, "."))
    fans_board_width = o_v3 + 2 * o_v2 + 2 * o_v4

    partition_length = o_v3 - 2 * fans_board_thickness - 2 * f_v2
    partition_length = Int(Format(partition_length + 1, ".")) + 2 * p_v1
    partition_width = Int(Format(inner_dist - p_v3, ".")) + 2 * p_v1
    
    comp_start(0) = start(0)
    comp_start(1) = start(1) + partition_width + fans_board_width + 2 * layout_range
    comp_title = Format(power, ".00") & " HP" & Space(3) & tube_diameter_r & Space(3) & tube_num_row & "R" & _
                 tube_num_stick & "T" & Space(3) & efficient_dist & "  m / m"
    If tube_diameter = 3 Then
        comp_title = comp_title & Space(3) & "(" & tube_head & ")"
    End If
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, 70

    If tube_hole_type = "橢圓孔" Then
        folding = "正折"
        num_inner_side_board = 2
    ElseIf tube_hole_type = "圓孔" Then
        folding = "文武邊"
        num_inner_side_board = 0
    End If

    ' 一般熱排 內端板
    If num_inner_side_board > 0 Then
        comp_start(0) = start(0) + outer_side_board_length + layout_range
        comp_start(1) = start(1) + partition_width + layout_range
        inner_side_board comp_start, inner_side_board_length, inner_side_board_width, d2
        ' 文字
        comp_start(0) = comp_start(0): comp_start(1) = comp_start(1) - text_dist
        comp_title = "內端板 " & Format(inner_side_board_thickness, "0.0") & "t " & partition_material
        ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height
        For Each in_num_row In tube_num_row_array
            comp_start(1) = comp_start(1) - text_height - 30
            comp_title = inner_side_board_length / tube_num_row * in_num_row & " X " & _
                         inner_side_board_width & " X " & _
                         num_inner_side_board & "只"
            ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height
        Next
    End If
    

    ' 一般熱排 外端板
    comp_start(0) = start(0)
    comp_start(1) = start(1) + partition_width + layout_range
    outer_side_board comp_start, outer_side_board_length, outer_side_board_width, o_v7, _
                     f_v1, o_v1, p_v3, o_v8, o_v6, fin_length, fin_width, d2, _
                     outer_side_board_thickness, d4

    comp_start(1) = comp_start(1) - text_dist
    
    comp_title = "外端板 " & Format(outer_side_board_thickness, "0.0") & "t " & material & "  " & folding
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height
    comp_start(1) = comp_start(1) - text_height - 30
    comp_title = outer_side_board_length & " X " & outer_side_board_width & " X " & "2只"
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, dim_text_height1

    ' 一般熱排 隔板
    If num_motor > 1 Then
        comp_start(0) = start(0): comp_start(1) = start(1)
        partition comp_start, partition_length, partition_width, p_v1, p_v2, p_v3, d2, p_v4
        
        comp_start(0) = comp_start(0): comp_start(1) = comp_start(1) - text_dist
        comp_title = "隔板 " & Format(fans_board_thickness, "0.0") & "t " & material & "  " & "正折"
        ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height
        comp_start(1) = comp_start(1) - text_height - 30
        comp_title = partition_length & " X " & partition_width & " X " & num_motor - 1 & "只"
        ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height
    End If
    
    ' 一般熱排 風斗板
    ' TODO:  以start為推斷基準 不要以隔板起點
    If num_motor > 1 Then ' 隔板存在
        If num_inner_side_board > 0 Then ' 內端板存在
            comp_start(0) = start(0) + inner_side_board_length + outer_side_board_length + _
                            2 * layout_range
        Else
            comp_start(0) = start(0) + partition_length + layout_range
        End If
        comp_start(1) = start(1)
    Else
        comp_start(0) = start(0) + outer_side_board_length + layout_range
        comp_start(1) = start(1) + partition_width + layout_range
    End If
    
    fans_board comp_start, fans_board_length, fans_board_width, d1, o_v2, o_v3, o_v4, o_v6, _
                           d3, fans_board_thickness

    comp_start(0) = comp_start(0): comp_start(1) = comp_start(1) - text_dist
    comp_title = "風斗板 " & Format(fans_board_thickness, "0.0") & "t " & material & "  " & "請參照折面"
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height
    comp_start(1) = comp_start(1) - text_height - 30
    comp_title = fans_board_length & " X " & fans_board_width & " X " & "1只"
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height

End Sub


Private Sub CommandButton1_Click()
    UserForm1.Hide

    t2 = ComboBox2.Text
    t3 = TextBox3.Text
    ' t3 = Val(TextBox3.Text)
    t4 = Val(TextBox4.Text)
    t5 = Val(TextBox5.Text)
    t6 = ComboBox6.Text
    t7 = Val(ComboBox7.Text)
    t8 = Val(TextBox8.Text)
    t9 = Val(TextBox9.Text)
    t10 = Val(TextBox10.Text)
    t11 = Val(TextBox11.Text)
    t12 = Val(TextBox12.Text)
    t13 = Val(TextBox13.Text)
    ' t14 = ComboBox14.Text
    t15 = Val(TextBox15.Text)
    t16 = Val(TextBox16.Text)
    t17 = Val(TextBox17.Text)
    t18 = ComboBox18.Text
    ' t19 = Val(TextBox19.Text)
    t20 = Val(ComboBox20.Text)
    ' t21 = Val(TextBox21.Text)
    t22 = Val(ComboBox22.Text)
    t23 = ComboBox23.Text

    tmp = Split(t2, "  ")
    tube_diameter = tmp(0)
    If tube_diameter = 3 Then
        tube_head = tmp(1)
    End If

    tube_num_row_array = Split(t3, "/")
    tube_num_row = 0
    For Each r In tube_num_row_array
        tube_num_row = tube_num_row + r
    Next

    tube_num_stick = t4
    efficient_dist = t5
    material = t6

    fans_board_thickness = t7
    fan_diameter = t8
    num_motor = t9
    motor_frame_diagonal = t10
    num_screw = t11
    screw_dist = t12
    num_part_screw = t13

    outer_side_board_in_length = t15
    connect_width = t16
    out_sb_out_d = t17
    tube_hole_type = t18
    out_sb_in_d = t19
    outer_side_board_thickness = t20
    inner_side_board_thickness = t22
    
    outer_side_board_up_hole = t23
    heater
End Sub


Private Sub ComboBox1_Change()

End Sub


Private Sub UserForm_Activate()
    ComboBox2.AddItem "2.5"
    ComboBox2.AddItem "3.0  P19.05"
    ComboBox2.AddItem "3.0  P22"
    ComboBox2.AddItem "4.0"
    ComboBox2.AddItem "5.0"

    ComboBox6.AddItem "錏板"
    ComboBox6.AddItem "鋁花板"
    ComboBox6.AddItem "鋁平板"
    ComboBox6.AddItem "SUS304"

    ComboBox7.AddItem "0.8"
    ComboBox7.AddItem "1.0"
    ComboBox7.AddItem "1.2"
    ComboBox7.AddItem "1.5"
    ComboBox7.AddItem "1.6"
    ComboBox7.AddItem "2.0"
    ComboBox7.AddItem "3.0"

    ComboBox20.AddItem "0.8"
    ComboBox20.AddItem "1.0"
    ComboBox20.AddItem "1.2"
    ComboBox20.AddItem "1.5"
    ComboBox20.AddItem "1.6"
    ComboBox20.AddItem "2.0"
    ComboBox20.AddItem "3.0"

    ComboBox22.AddItem "0.8"
    ComboBox22.AddItem "1.0"
    ComboBox22.AddItem "1.2"
    ComboBox22.AddItem "1.5"
    ComboBox22.AddItem "1.6"
    ComboBox22.AddItem "2.0"
    ComboBox22.AddItem "3.0"
    
    ComboBox18.AddItem "圓孔"
    ComboBox18.AddItem "橢圓孔"

    ComboBox23.AddItem "是"
    ComboBox23.AddItem "否"

End Sub



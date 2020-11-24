




' Public Sub AddRectDim(ByVal p1, ByVal p2, ByVal text_gap)
'     Dim length_text(2)  As Double
'     Dim width_text(2)  As Double
'     Dim x(2)  As Double

'     length_text(0) = (p2(0) - p1(0)) / 2: length_text(1) = p1(1) - text_gap
'     width_text(0) = p1(0) - text_gap: width_text(1) = (p2(1) - p1(1)) / 2
'     'Add the dimension in the drawing.
'     x(0) = p2(0): x(1) = p1(1)
'     Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, x, length_text)
'     x(0) = p1(0): x(1) = p2(1)
'     Set AcadDimAligned2 = ThisDrawing.ModelSpace.AddDimAligned(p1, x, width_text)
'     'Format the dimension object according to your needs.
'     With AcadDimAligned
'         .TextWidth = 10
'         .TextGap = 5              'The distance of the dimension text from the dimension line.
'         .Arrowhead1Type = 0         'acArrowOblique in early binding
'         .Arrowhead2Type = 0         'For the standard dimension arrow put 0 here.
'         .ArrowheadSize = 10
'         .ExtensionLineExtend = 10   'The amount to extend the extension line beyond the dimension line.
'         .DimensionLineColor = acGreen
'         .ExtensionLineColor = acGreen
'     End With
'     With AcadDimAligned2
'         .TextWidth = 10
'         .TextGap = 5               'The distance of the dimension text from the dimension line.
'         .Arrowhead1Type = 0         'acArrowOblique in early binding
'         .Arrowhead2Type = 0         'For the standard dimension arrow put 0 here.
'         .ArrowheadSize = 10
'         .ExtensionLineExtend = 10   'The amount to extend the extension line beyond the dimension line.
'         .DimensionLineColor = acGreen
'         .ExtensionLineColor = acGreen
'     End With
' End Sub


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


Public Sub AddLinedCircles(ByVal start, ByVal radius, ByVal distance, ByVal num_circle, ByVal direction)
    Dim c(2)  As Double
    Dim cir_obj As AcadCircle
    c(0) = start(0): c(1) = start(1)
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(c, radius)

    For i = 1 To num_circle - 1
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
    Next
End Sub


Public Sub SelectActiveLayer(ByVal layer_name)
    For Each lay0 In ThisDrawing.Layers ' 在所有的圖層中進行循環
        If lay0.Name = layer_name Then ' 如果找到圖層名
            ThisDrawing.ActiveLayer = lay0 ' 把當前圖層設為已經存在的圖層
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


    ' 輪廓
    ' Dim cc(2)  As Double
    ' cc(0) = 3000: cc(1) = 3000
    ' rect_start = starts(0)
    ' rect_end = ends(0)
    ' l = Abs(rect_start(0) - rect_end(0))
    ' w = Abs(rect_start(1) - rect_end(1))
    ' AddRect rect_start, l, w

    ' cc(0) = 4000: cc(1) = 4000
    ' rect_start = starts(1)
    ' rect_end = ends(1)
    ' l = Abs(rect_start(0) - rect_end(0))
    ' w = Abs(rect_start(1) - rect_end(1))
    ' AddRect rect_start, l, w
    ' i = 0
    ' For Each rect_start In starts
    '     ' rect_start = starts(i)
    '     rect_end = ends(i)
    '     l = Abs(rect_start(0)-rect_end(0))
    '     w = Abs(rect_start(1)-rect_end(1))
    '     AddRect starts(i), l, w
    '     i = i + 1
    ' Next
    ' AddRect start, 2000, 1000

    ' 折線
    ' Set line_obj = ThisDrawing.ModelSpace.AddLine(start, p1)

End Sub


Public Sub AddHill(ByVal start, ByVal l1, ByVal l2, ByVal l3, ByVal w1, ByVal w2, ByVal direction)
    Dim p1(2) As Double
    Dim p2(2) As Double
    Dim line_obj As AcadLine
    If direction = "h" Then
        p1(0) = start(0) + l1: p1(1) = start(1)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(start, p1)
        p2(0) = p1(0): p2(1) = p1(1) + w1
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        p1(0) = p2(0) + l2: p1(1) = p2(1)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        p2(0) = p1(0): p2(1) = p1(1) - w2
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        p1(0) = p2(0) + l3: p1(1) = p2(1)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
    ElseIf direction = "h_flip" Then
        p1(0) = start(0) + l1: p1(1) = start(1)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(start, p1)
        p2(0) = p1(0): p2(1) = p1(1) - w1
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        p1(0) = p2(0) + l2: p1(1) = p2(1)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        p2(0) = p1(0): p2(1) = p1(1) + w2
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
        p1(0) = p2(0) + l3: p1(1) = p2(1)
        Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
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
End Sub


Public Sub AddMainDims(ByVal h1, ByVal h2, ByVal h_dist, ByVal h_dir, _
                       ByVal v1, ByVal v2, ByVal v_dist, ByVal v_dir, _
                       ByVal text_height, ByVal arrow_size)
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
    AcadDimAligned.ArrowheadSize = arrow_size
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(v1, v2, v_text)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size
End Sub


Public Sub AddTwoCrossRects(ByVal start, ByVal length, ByVal width, ByVal l1, ByVal l2, _
                            ByVal w1, ByVal w2)
    
    Dim start_r1(2)  As Double, start_r2(2)  As Double
    Dim a(2)  As Double, b(2)  As Double, c(2)  As Double, d(2)  As Double, _
        e(2)  As Double, f(2)  As Double, g(2)  As Double, h(2)  As Double, i(2)  As Double
    Dim t1(2)  As Double, t2(2)  As Double, t3(2)  As Double, t4(2)  As Double, _
        t5(2)  As Double, t6(2)  As Double, t7(2)  As Double, t8(2)  As Double
    start_r1(0) = start(0): start_r1(1) = start(1) + w1
    start_r2(0) = start(0) + l1: start_r2(1) = start(1)
    in_length = length - l1 - l2
    in_width = width - w1 - w2
    
    
    text_height = 30
    arrow_size = 30
    dim_dist = 100
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
    t8(0) = start_r1(0) + 0.5 * l1: t8(1) = start_r2(1) + width + dim_dist

    SelectActiveLayer "鈑金"
    AddRect start_r1, length, in_width
    AddRect start_r2, in_length, width

    SelectActiveLayer "主尺寸"
    AddMainDims start_r1, c, dim_dist, "down", start_r2, g, dim_dist, "left", text_height, arrow_size

    SelectActiveLayer "其他尺寸"
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(a, b, t3)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size
    TextOutsideAlign = 1
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(b, e, t4)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(e, f, t5)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(e, d, t6)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(e, h, t7)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(h, i, t8)
    AcadDimAligned.TextHeight = text_height
    AcadDimAligned.ArrowheadSize = arrow_size
    SelectActiveLayer "鈑金"
End Sub


Public Sub up_board(ByVal start, ByVal length, ByVal width)
    ' Dim start(2) ' 宣告2x2陣列用於儲存所有長方形兩頂點
    Dim ends(2) ' 宣告2x2陣列用於儲存所有長方形兩頂點
    Dim fold_lens(2, 1) As Double ' 宣告3x2陣列用於儲存所有折線間距
    Dim a_start(2) As Double, a_end(2) As Double, b_start(2) As Double, b_end(2) As Double

    l1 = 10: l2 = 10: w1 = 0: w2 = 12 ' 連接處
    a1 = 32: a2 = 20 ' 折線間距
    b1 = 20: b2 = 10: b3 = 10
    v1 = 52.5
    v2 = 12
    start(1) = start(1) + a1 + a2 + b1 + b2 + b3
    AddTwoCrossRects start, length, width - (a1 + a2 + b1 + b2 + b3), l1, l2, w1, w2
    ends(0) = start(0) + l1 + l2: ends(1) = start(1) + w1 + w2
    start(0) = start(0) + length - w2 - 3
    AddConnect start, ends, 15, 49.5, 597, a1 + a2, b1 + b2 + b3
    ' a_start(0) = start(0) + v2: a_start(1) = start(1) + a1 + a2 + b1 + b2 + b3
    ' a_end(0) = start(0) + length - v2: a_end(1) = start(1) + b1 + b2 + b3
    ' b_start(0) = start(0) + v2 + v1: b_start(1) = a_end(1)
    ' b_end(0) = start(0) + length - v1 - v2: b_end(1) = start(1)

    ' a_start(0) = 1000: a_start(1) = 1000
    ' a_end(0) = 1800: a_end(1) = 600
    ' b_start(0) = 1200: b_start(1) = 600
    ' b_end(0) = 1600: b_end(1) = 200

    ' starts(0) = a_start: starts(1) = b_start
    ' ends(0) = a_end: ends(1) = b_end
    ' fold_lens(0, 0) = a1: fold_lens(1, 0) = a2
    ' ' fold_lens(0, 1) = b1: fold_lens(1, 1) = b2: fold_lens(2, 1) = b3

    ' Dim ss(1)
    ' ss(0) = start(0): ss(1) = start(1) + a1 + a2 + b1 + b2 + b3
    ' AddTwoCrossRects ss, length, width, l1, l2, 0, w2
    ' AddConnect starts, ends, fold_lens
    ' ByVal start, ByVal end, ByVal l1, ByVal l2, ByVal l3, ByVal w1, ByVal w2
End Sub


Public Sub inner_board(ByVal start, ByVal length, ByVal width)
    Dim p1(2)  As Double
    Dim p2(2)  As Double
    Dim center(2)  As Double ' 螺絲位置 圓心
    Dim cir_obj As AcadCircle ' 定義圓的變數名稱

    s1 = start(0) + 200 + length: s2 = start(1) ' 作圖起始點
    ' start(0) = start(0) + 1000: start(1) = start(1)
    v1 = 10
    
    in_length = length - 2 * v1
    in_width = width - 2 * v1

    ' 鈑金
    AddTwoCrossRects start, length, width, v1, v1, v1, v1

    ' 螺絲孔
    r = 5 ' 半徑
    range = 15 '螺絲孔距離
    center(0) = start(0) + v1 + range: center(1) = start(1) + v1 + range
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r) ' 畫圓
    center(0) = start(0) + v1 + range: center(1) = start(1) + v1 + range + in_width - 2 * range
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r)
    center(0) = start(0) + v1 + range + in_length - 2 * range: center(1) = start(1) + v1 + range + in_width - 2 * range
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r)
    center(0) = start(0) + v1 + range + in_length - 2 * range: center(1) = start(1) + v1 + range
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r)
End Sub


Public Sub padding_board(ByVal start, ByVal length, ByVal width, ByVal v1, ByVal c1, ByVal c2)
    Dim p1(2)  As Double
    Dim p2(2)  As Double
    Dim line_obj As AcadDimAligned

    l1 = c1 + v1
    w1 = c2
    in_length = length - 2 * l1
    in_width = width - 2 * w1

    ' 鈑金
    AddTwoCrossRects start, length, width, l1, l1, w1, w1
    p1(0) = start(0) + v1: p1(1) = start(1) + w1
    p2(0) = start(0) + v1: p2(1) = start(1) + w1 + in_width
    Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
    p1(0) = start(0) + length - v1: p1(1) = start(1) + w1
    p2(0) = start(0) + length - v1: p2(1) = start(1) + w1 + in_width
    Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
End Sub


Public Sub gauge_board(ByVal start, ByVal length, ByVal width, ByVal num_gauge, ByVal side_dist)
    
    Dim in_length As Double ' 配電箱 內長
    Dim in_width As Double ' 配電箱 內寬
    Dim center(2)  As Double ' 螺絲位置 圓心
    Dim cir_obj As AcadCircle ' 定義圓的變數名稱
    
    l1 = 15
    w1 = 25
    w2 = 15
    v4 = 120 ' 錶頭間距
    v5 = 10  ' 連接螺絲與邊緣垂直距離
    v6 = 50 ' TODO: wrong vars
    d1 = 68  ' 錶頭直徑
    d2 = 4  ' 螺絲孔徑
    r1 = d1 / 2
    r2 = d2 / 2
    in_length = length - 2 * l1
    in_width = width - w2 - w1
    v7 = in_length - 2 * v6 ' TODO: wrong vars
    v8 = r1 + 4.525 ' TODO: wrong vars

    
    ' 鈑金
    AddTwoCrossRects start, length, width, l1, l1, w1, w2

    ' 連接螺絲孔
    center(0) = start(0) + l1 + v6: center(1) = start(1) + v5
    AddLinedCircles center, r2, v7 / 2, 3, 0

    ' 錶頭孔
    g1 = start(0) + l1 + side_dist + r1: g2 = start(1) + w1 + in_width / 2
    center(0) = g1: center(1) = g2
    AddLinedCircles center, r1, v4, num_gauge, 0

    ' 錶頭旁螺絲孔
    center(0) = g1: center(1) = g2 + v8
    AddLinedCircles center, r2, v4, num_gauge, 0
    center(0) = g1 - v8 * (Sqr(3) / 2): center(1) = g2 - 0.5 * v8
    AddLinedCircles center, r2, v4, num_gauge, 0
    center(0) = g1 + v8 * (Sqr(3) / 2): center(1) = g2 - 0.5 * v8
    AddLinedCircles center, r2, v4, num_gauge, 0
    
End Sub


Public Sub door(ByVal start, ByVal length, ByVal width)
    Dim in_length As Variant ' 配電箱 內長
    Dim in_width As Variant ' 配電箱 內寬

    Dim center(2)  As Double ' 螺絲位置 圓心
    Dim cir_obj As AcadCircle ' 定義圓的變數名稱
    
    Dim p1(2)  As Double

    s1 = start(0): s2 = start(1) ' 作圖起始點
    l1 = 19
    l2 = 12
    w1 = 14
    w2 = 14
    v2 = 14
    v3 = 12
    v4 = 5
    v5 = 9
    v6 = 35
    v7 = 23 ' 手把 長
    v8 = 62 ' 手把 寬
    v9 = (width - v8) / 2 - v2
    in_length = length - l1 - l2
    in_width = width - w1 - w2
    v10 = (in_width - v8) / 2

    ' 鈑金
    AddTwoCrossRects start, length, width, l1, l2, w1, w2
    p1(0) = s1 + l1 + v6: p1(1) = s2 + w2 + v10
    AddRect p1, v7, v8

    ' 螺絲孔
    r = 4 ' 半徑
    center(0) = s1 + length - l2 - v4: center(1) = s2 + v5
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r) ' 畫圓
    center(0) = s1 + length - l2 - v4: center(1) = s2 + width - v5
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r)

    r = 3.5 ' 半徑
    center(0) = s1 + l1 + v6 + 0.5 * v7: center(1) = s2 + v2 + v9 - r
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r) ' 畫圓
    center(0) = s1 + l1 + v6 + 0.5 * v7: center(1) = s2 + v2 + v9 + v8 + r
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r)
End Sub


Public Sub main_part(ByVal start, ByVal length, ByVal width, ByVal depth, ByVal main_length)
    ' 主體
    Dim line_obj As AcadLine
    Dim p1(2)  As Double
    Dim p2(2)  As Double
    Dim ends(2)  As Double
    v1 = 10
    v2 = 70
    v3 = 32
    v4 = depth + (length - 200) / 2
    v4 = 430
    dim_dist = 200
    text_height = 40: arrow_size = 20

    ' 鈑金
    p1(0) = start(0) + v1: p1(1) = start(1) + v3
    AddHill p1, v2, v4, 200, v3, 50, "h_flip"
    AddHill p1, 14, 447, 25, v1, v1, "v_flip"
    p1(1) = p1(1) + 14 + 25 + 447 ' TODO: 14 + 25 + 447=486
    AddHill p1, v2, main_length + 2 * depth, v2, v3, v3, "h"
    p1(0) = p1(0) + length - 2 * v1: p1(1) = p1(1) - 486
    AddHill p1, 14, 447, 25, v1, v1, "v"
    p1(0) = p1(0) - v4 - v2 - 200: p1(1) = p1(1) + (50 - v3)
    AddHill p1, 200, v4, v2, 50, v3, "h_flip"

    ' 螺絲孔
    ' r = 2
    ' p = 
    ' AddRectCircles p, r, length, width
    ' AddRectCircles p, r, length, width
    ' r = 5
    ' p = 
    ' l = 1: w = 
    ' AddRectCircles p, r, l, w

    ' 折線



    ' 標註
    ' SelectActiveLayer "主尺寸"
    ' h1 = 
    ' AddMainDims h1, h2, dim_dist, "down", v1, v2, dim_dist, "left", text_height, arrow_size


End Sub


Public Sub electrical_box(ByVal length, ByVal width, ByVal thickness, _
                          ByVal material, ByVal num_gauge) ' 配電箱
    Set lay1 = ThisDrawing.Layers.Add("鈑金") ' 增加一個名為“鈑金”的圖層
    Set lay2 = ThisDrawing.Layers.Add("主尺寸") ' 增加一個名為“主尺寸”的圖層
    lay2.color = 1 ' 圖層設置為紅色
    Set lay3 = ThisDrawing.Layers.Add("其他尺寸") ' 增加一個名為“其他尺寸”的圖層
    lay3.color = 3 ' 圖層設置為綠色
    
    ThisDrawing.ActiveLayer = lay1 ' 將“鈑金”設置為當前圖層

    On Error Resume Next ' 如果有錯誤, 不管他
    ' 刪除所有作圖
    For Each oEntity In ThisDrawing.ModelSpace
        oEntity.Delete
    Next

    ' 700 X 550 配電箱
    Dim start(2)  As Double ' 全圖作圖起始點
    Dim comp_start(2)  As Double ' 配件作圖起始點
    ' 內板
    Dim in_borad_length As Double, in_borad_width As Double
    start(0) = 1000: start(1) = 1000

    length = 700
    width = 550
    num_gauge = 4
    thickness = 1
    
    depth = 180
    dim_dist = 100
    text_dist = 100
    text_height = 40
    door_dist = 30
    range = 8 ' 配電箱 門開關順利所預留間隙
    side_dist = 61 ' 錶孔至邊緣水平距離
    gauge_dist = 120 ' 錶孔間距

    
    ' 配電箱 主體
    ' 圖面
    comp_start(0) = start(0): comp_start(1) = start(1) + 1500
    comp_length = length + 2 * depth + 2 * (30 + 20 + 20 + 10)
    comp_width = width
    main_part comp_start, comp_length, comp_width, depth, length
    ' 文字
    comp_start(1) = comp_start(1) - dim_dist - text_dist
    comp_title = thickness & "  " & material & "     " & comp_length & " X " & comp_width & "  主板金"
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height

    comp_start(1) = comp_start(1) + comp_width + 500
    comp_title = length & " X " & width & "  配電盒 (上蓋外包)"
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height

    ' 配電箱 門
    comp_start(0) = start(0): comp_start(1) = start(1)
    comp_length = length - 2 * door_dist - range + 19 + 12
    comp_width = width - 2 * door_dist - range + 14 + 14
    door comp_start, comp_length, comp_width
    comp_start(1) = comp_start(1) - dim_dist - text_dist
    ' SelectActiveLayer "主尺寸"
    comp_title = thickness & "  " & material & "     " & comp_length & " X " & comp_width & "  門板"
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height

    ' 配電箱 內板
    comp_start(0) = start(0) + 1000: comp_start(1) = start(1)
    comp_length = length
    inner_board comp_start, 670, 450
    comp_start(1) = comp_start(1) - dim_dist - text_dist
    comp_title = thickness & "  " & material & "     " & comp_length & " X " & comp_width & "  配電盤"
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height
    

    ' 配電箱 上板
    ' a1 = 32: a2 = 20 ' 折線間距
    ' b1 = 20: b2 = 10: b3 = 10
    ' w2 = 12
    ' comp_start(0) = start(0) + 2000: comp_start(1) = start(1) + 1500
    ' comp_length = length + 2 * (thickness + 12)
    ' comp_width = depth + thickness + (a1 + a2 + b1 + b2 + b3 + w2)
    ' up_board comp_start, comp_length, comp_width
    ' comp_start(1) = comp_start(1) - dim_dist - text_dist
    ' comp_title = thickness & "  " & material & "     " & comp_length & " X " & comp_width
    ' ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height

    ' 配電箱 底板
    ' bottom_board start, 717, 258

    ' 配電箱 錶板 *
    l1 = 15
    d = 68 ' 錶孔直徑
    comp_start(0) = start(0) + 2000: comp_start(1) = start(1) + 500
    comp_length = 2 * l1 + (num_gauge - 1) * gauge_dist + 2 * side_dist + d
    comp_width = 162 ' 錶鈑高通常固定為162
    gauge_board comp_start, comp_length, comp_width, num_gauge, side_dist
    comp_start(1) = comp_start(1) - dim_dist - text_dist
    comp_title = thickness & "  " & material & "     " & comp_length & " X " & comp_width & "  錶板"
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height

    ' 配電箱 配合機房箱墊板 *
    ' TODO: 變量
    v1 = 73 ' 配合外門高
    c1 = 20
    c2 = 15
    t = 1 ' 伸縮量
    comp_start(0) = start(0) + 2000: comp_start(1) = start(1) - 500
    comp_length = length - 2 * t + 2 * (c1 + v1)
    comp_width = depth - 2 * t + 2 * c2
    padding_board comp_start, comp_length, comp_width, v1, c1, c2
    comp_start(1) = comp_start(1) - dim_dist - text_dist
    comp_title = thickness & "  " & material & "     " & comp_length & " X " & comp_width & "  墊板"
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

    length = Val(TextBox1.Text)
    width = Val(TextBox2.Text)
    thickness = Val(TextBox3.Text)
    material = TextBox4.Text
    num_gauge = Val(TextBox5.Text)
    electrical_box length, width, thickness, material, num_gauge
End Sub


Private Sub Label2_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub UserForm_Click()

End Sub










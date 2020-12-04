




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

Public fin_length As Double
Public fin_width As Double

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
    
    ' TODO: 參數拉出去
    ' TODO: dimension in right position
    text_height = 25
    arrow_size = 20
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
    t8(0) = start_r1(0) + 0.5 * l1 - 50: t8(1) = start_r2(1) + width + dim_dist

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


Public Sub fans_board(ByVal start, ByVal d1, ByVal o_v2, ByVal o_v3, ByVal o_v4, ByVal o_v5, _
                      ByVal o_v6, ByVal comp_width, ByVal d3)
    Dim line_obj As AcadLine
    Dim p1(2) As Double
    Dim p2(2) As Double

    fan_thickness = thickness
    fan_efficient_dist = efficient_dist
    fan_num_motor = num_motor
    fan_motor_frame_length = motor_frame_length
    fan_motor_frame_width = motor_frame_width
    comp_length = fan_efficient_dist + 2*o_v4

    ' 鈑金
    AddHill start, 0, comp_length, 0, o_v4, o_v4, "h_flip"
    p1(0) = start(0): p1(1) = start(1) + o_v2
    AddHill p1, 0, comp_length, 0, o_v2, o_v2, "h_flip"
    p1(0) = p1(0): p1(1) = p1(1) + o_v3
    AddHill p1, 0, comp_length, 0, o_v3, o_v3, "h_flip"
    p1(0) = p1(0): p1(1) = p1(1) + o_v2
    AddHill p1, 0, comp_length, 0, o_v2, o_v2, "h_flip"
    p1(0) = p1(0): p1(1) = p1(1) + o_v4
    AddHill p1, 0, comp_length, 0, o_v4, o_v4, "h_flip"
    p2(0) = p1(0) + comp_length: p2(1) = p1(1)
    Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)

    ' 螺絲
    p1(0) = start(0) + o_v4/2: p1(1) = start(1) - o_v5
    AddRectCircles p1, d1/2, comp_length-o_v4, comp_width-2*o_v5
    p1(0) = start(0) + o_v4/2: p1(1) = start(1) + o_v2 - fan_thickness - o_v6
    AddRectCircles p1, d1/2, comp_length-o_v4, o_v3+2*(fan_thickness+o_v6)

    ' 風斗與馬達架孔
    ' AddLinedCircles start, radius, distance, num_stick, 1
    ' For i = 1 To fan_num_motor
    '     AddRectCircles p1, d3/2, fan_motor_frame_length, fan_motor_frame_width
    ' Next
End Sub


Public Sub inner_side_board(ByVal start, ByVal row_dist, ByVal stick_dist)
    in_tube_num_row = tube_num_row
    i_tube_num_stick = tube_num_stick
    ' TODO: 14
    AddFinCircles start, i_tube_num_row, i_tube_num_stick, row_dist, stick_dist, 14
End Sub


Public Sub partition(ByVal start, ByVal v1, ByVal v2, ByVal v3, ByVal d1, ByVal dist1)
    Dim comp_length As Double
    Dim comp_width As Double
    Dim p(2) As Double

    p_fin_length = fin_length
    p_inner_dist = inner_dist
    p_num_part_screw = num_part_screw
    p_screw_dist = screw_dist

    ' 鈑金
    comp_length = p_fin_length + 2*v1
    comp_width = p_inner_dist - v3 + 2*v1
    AddTwoCrossRects start, comp_length, comp_width, v1, v1, v1, v1

    ' 螺絲
    p(0) = start(0) + v1 + (p_fin_length-p_screw_dist)/2: p(1) = start(1) + comp_width - dist1
    AddLinedCircles p, d1/2, p_screw_dist/(p_num_part_screw-1), p_num_part_screw, 0

    ' 文字
    comp_title = "隔板 " & thickness & "t " & material & "  " & comp_length & " X " & comp_width _ 
                  & " X " & "2只"
    ThisDrawing.ModelSpace.AddText comp_title, start, text_height
    
End Sub


Public Sub outer_side_board(ByVal start)
End Sub


Public Sub heater() ' 一般熱排
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

    ' 10 HP 3分5R20T 1454 mm
    Dim start(2)  As Double ' 全圖作圖起始點
    Dim comp_start(2)  As Double ' 配件作圖起始點
    start(0) = 1000: start(1) = 1000
    
    ' TODO: 輸入不合法 報錯
    ' TODO: 管徑 彎頭 考慮下拉式選單
    tube_head = 22
    tube_diameter = 3
    tube_num_row = 5
    tube_num_stick = 20
    efficient_dist = 1454
    fan_diameter = 12
    motor_frame_length = 272
    motor_frame_width = 272
    thickness = 1
    material = "錏板"
    partition_material = "鋁板"
    num_motor = 2
    inner_dist = 100
    connect_width = 14
    num_screw = 4
    screw_dist = 405
    num_part_screw = 3

    ' 排支數換算鰭片長寬
    If tube_diameter = 2.5 Then
        row_dist = 19.05
        stick_dist = 25.4
    ElseIf tube_diameter = 3 Then
        If tube_head = 22 Then
            row_dist = 22
            stick_dist = 25.4
        ElseIf tube_head = 19.05 Then
            row_dist = 19.05
            stick_dist = 25.4
        End If
    ElseIf tube_diameter = 4 Or tube_diameter = 5 Then
        row_dist = 33
        stick_dist = 38.1
    ' TODO: 
    ' Else
        ' error message
    End If

    fin_length = tube_num_stick * stick_dist
    fin_width = tube_num_row * row_dist 
    d1 = 5
    d2 = 4
    d3 = 11.5

    screw1 = 5
    screw2 = 3.2
    text_height = 30
    text_dist = 200
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

    ' 一般熱排 風斗板
    comp_length = efficient_dist + 2*o_v4
    comp_width = o_v3 + 2*o_v2 + 2*o_v4

    comp_start(0) = start(0) + 1000: comp_start(1) = start(1)
    fans_board comp_start, d1, o_v2, o_v3, o_v4, o_v5, o_v6, comp_width, d3

    comp_start(0) = comp_start(0): comp_start(1) = comp_start(1) - text_dist
    comp_title = "風斗板 " & thickness & "t " & material & "  " & comp_length & " X " & comp_width _ 
                  & " X " & "1只"
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height

    ' 一般熱排 內端板
    inner_side_board comp_start, row_dist, stick_dist

    ' 一般熱排 隔板
    comp_length = fin_length + 2*p_v1
    comp_width = inner_dist - p_v3 + 2*p_v1

    comp_start(0) = start(0): comp_start(1) = start(1)
    partition comp_start, p_v1, p_v2, p_v3, d2, p_v4
    
    comp_start(0) = comp_start(0): comp_start(1) = comp_start(1) - text_dist
    comp_title = "隔板 " & thickness & "t " & material & "  " & comp_length & " X " & comp_width _ 
                  & " X " & "2只"
    ThisDrawing.ModelSpace.AddText comp_title, comp_start, text_height
    
    
    
    ' 一般熱排 外端板
    ' outer_side_board

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



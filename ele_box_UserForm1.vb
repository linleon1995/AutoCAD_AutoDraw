

Public Sub AddRectDim(ByVal p1, ByVal p2, ByVal text_gap)
    Dim length_text(2)  As Double
    Dim width_text(2)  As Double
    Dim x(2)  As Double

    length_text(0) = (p2(0) - p1(0)) / 2: length_text(1) = p1(1) - text_gap
    width_text(0) = p1(0) - text_gap: width_text(1) = (p2(1) - p1(1)) / 2
    'Add the dimension in the drawing.
    x(0) = p2(0): x(1) = p1(1)
    Set AcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(p1, x, length_text)
    x(0) = p1(0): x(1) = p2(1)
    Set AcadDimAligned2 = ThisDrawing.ModelSpace.AddDimAligned(p1, x, width_text)
    'Format the dimension object according to your needs.
    With AcadDimAligned
        .TextHeight = 10
        .TextGap = 5              'The distance of the dimension text from the dimension line.
        .Arrowhead1Type = 0         'acArrowOblique in early binding
        .Arrowhead2Type = 0         'For the standard dimension arrow put 0 here.
        .ArrowheadSize = 10
        .ExtensionLineExtend = 10   'The amount to extend the extension line beyond the dimension line.
        .DimensionLineColor = acGreen
        .ExtensionLineColor = acGreen
    End With
    With AcadDimAligned2
        .TextHeight = 10
        .TextGap = 5               'The distance of the dimension text from the dimension line.
        .Arrowhead1Type = 0         'acArrowOblique in early binding
        .Arrowhead2Type = 0         'For the standard dimension arrow put 0 here.
        .ArrowheadSize = 10
        .ExtensionLineExtend = 10   'The amount to extend the extension line beyond the dimension line.
        .DimensionLineColor = acGreen
        .ExtensionLineColor = acGreen
    End With
End Sub

Public Sub AddRect(ByVal p1, ByVal p2)
    Dim line_obj As AcadLine
    Dim p3(2)  As Double
    p3(0) = p1(0): p3(1) = p2(1):
    Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p3)
    Set line_obj = ThisDrawing.ModelSpace.AddLine(p3, p2)
    p3(0) = p2(0): p3(1) = p1(1):
    Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p3)
    Set line_obj = ThisDrawing.ModelSpace.AddLine(p3, p2)
End Sub


Public Sub inner_board(ByVal start, ByVal length, ByVal width)
    Dim p1(2)  As Double
    Dim p2(2)  As Double
    Dim center(2)  As Double ' 螺絲位置 圓心
    Dim cir_obj As AcadCircle ' 定義圓的變數名稱

    s1 = start(0) + 200 + length: s2 = start(1) ' 作圖起始點
    v1 = 10
    
    in_length = length - 2 * v1
    in_width = width - 2 * v1

    ' 鈑金
    p1(0) = s1 + v1: p1(1) = s2
    p2(0) = s1 + v1 + in_length: p2(1) = s2 + width
    AddRect p1, p2
    p1(0) = s1: p1(1) = s2 + v1
    p2(0) = s1 + length: p2(1) = s2 + v1 + in_width
    AddRect p1, p2

    ' 螺絲孔
    r = 5 ' 半徑
    range = 15 '螺絲孔距離
    center(0) = s1 + v1 + range: center(1) = s2 + v1 + range
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r) ' 畫圓
    center(0) = s1 + v1 + range: center(1) = s2 + v1 + range + in_width - 2 * range
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r)
    center(0) = s1 + v1 + range + in_length - 2 * range: center(1) = s2 + v1 + range + in_width - 2 * range
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r)
    center(0) = s1 + v1 + range + in_length - 2 * range: center(1) = s2 + v1 + range
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r)
End Sub


Public Sub padding_board(ByVal start, ByVal length, ByVal width)
    Dim p1(2)  As Double
    Dim p2(2)  As Double
    Dim line_obj As AcadLine
    s1 = start(0) + 200 + 2 * length: s2 = start(1) - 400 ' 作圖起始點
    v1 = 20
    v2 = 73
    v3 = 15

    in_length = length - 2 * (v1 + v2)
    in_width = width - 2 * v3

    ' 鈑金
    p1(0) = s1 + v1 + v2: p1(1) = s2
    p2(0) = s1 + v1 + in_length: p2(1) = s2 + width
    AddRect p1, p2
    p1(0) = s1: p1(1) = s2 + v3
    p2(0) = s1 + length: p2(1) = s2 + v3 + in_width
    AddRect p1, p2
    p1(0) = s1 + v1: p1(1) = s2 + v3
    p2(0) = s1 + v1: p2(1) = s2 + v3 + in_width
    Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
    p1(0) = s1 + length - v1: p1(1) = s2 + v3
    p2(0) = s1 + length - v1: p2(1) = s2 + v3 + in_width
    Set line_obj = ThisDrawing.ModelSpace.AddLine(p1, p2)
End Sub


Public Sub gauge_board(ByVal start, ByVal length, ByVal width)
    Set lay1 = ThisDrawing.Layers.Add("TBC") ' 增加一個名為“臨時圖層”的圖層
    lay1.Color = 3 ' 圖層設置為綠色
    ThisDrawing.ActiveLayer = lay1 ' 將當前圖層設置為新建圖層
    Dim in_length As Variant ' 配電箱 內長
    Dim in_width As Variant ' 配電箱 內寬

    Dim center(2)  As Double ' 螺絲位置 圓心
    Dim cir_obj As AcadCircle ' 定義圓的變數名稱
    
    Dim p1(2)  As Double
    Dim p2(2)  As Double

    s1 = start(0) + 2000: s2 = start(1) ' 作圖起始點
    v1 = 15
    v2 = 15
    v3 = 25
    v4 = 120 ' 錶頭間距
    v5 = 10  ' 連接螺絲距離
    v6 = 50
    d1 = 68  ' 錶頭直徑
    d2 = 4  ' 螺絲孔徑
    r1 = d1 / 2
    r2 = d2 / 2
    n1 = 4  ' 錶頭數量
    in_length = length - 2 * v1
    in_width = width - v2 - v3
    v7 = in_length - 2*v6
    v8 = r1 + 4.525

    
    ' 鈑金
    p1(0) = s1 + v1: p1(1) = s2
    p2(0) = s1 + v1 + in_length: p2(1) = s2 + width
    AddRect p1, p2
    p1(0) = s1: p1(1) = s2 + v3
    p2(0) = s1 + length: p2(1) = s2 + v3 + in_width
    AddRect p1, p2

    

    ' 連接螺絲孔
    center(0) = s1 + v1 + v6: center(1) = s2 + v5
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r2)
    center(0) = s1 + v1 + v6 + v7/2: center(1) = s2 + v5
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r2)
    center(0) = s1 + v1 + v6 + v7: center(1) = s2 + v5
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r2)

    g1 = s1 + v1 + (in_length - (n1 - 1) * v4) / 2: g2 = s2 + v3 + in_width / 2
    For i = 0 To n1 - 1
        ' 首先定位錶頭中心
        If i = 0 Then
            g_center1 = g1: g_center2 = g2
        Else
            g_center1 = g_center1 + v4
        End If

        ' 錶頭孔
        center(0) = g_center1: center(1) = g_center2
        Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r1) ' 畫圓

        ' 錶頭旁螺絲孔
        center(0) = g_center1: center(1) = g_center2 + v8
        Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r2) ' 畫圓
        center(0) = g_center1 - v8*(Sqr(3)/2): center(1) = g_center2 - 0.5*v8
        Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r2)
        center(0) = g_center1 + v8*(Sqr(3)/2): center(1) = g_center2 - 0.5*v8
        Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r2)
    Next
End Sub


Public Sub door(ByVal start, ByVal length, ByVal width)
    Dim in_length As Variant ' 配電箱 內長
    Dim in_width As Variant ' 配電箱 內寬

    Dim center(2)  As Double ' 螺絲位置 圓心
    Dim cir_obj As AcadCircle ' 定義圓的變數名稱
    
    Dim p1(2)  As Double
    Dim p2(2)  As Double

    s1 = start(0): s2 = start(1) ' 作圖起始點
    v1 = 19
    v2 = 14
    v3 = 12
    v4 = 5
    v5 = 9
    v6 = 35
    v7 = 23
    v8 = 62
    v9 = (width - v8) / 2 - v2
    in_length = length - v1 - v3
    in_width = width - 2 * v2
    v10 = (in_width - v8) / 2

    ' 鈑金
    p1(0) = s1 + v1: p1(1) = s2
    p2(0) = s1 + v1 + in_length: p2(1) = s2 + width
    AddRect p1, p2
    AddRectDim p1, p2, 200
    p1(0) = s1: p1(1) = s2 + v2
    p2(0) = s1 + length: p2(1) = s2 + v2 + in_width
    AddRect p1, p2
    AddRectDim p1, p2, 300
    p1(0) = s1 + v1 + v6: p1(1) = s2 + v2 + v10
    p2(0) = s1 + v1 + v6 + v7: p2(1) = s2 + v2 + v10 + v8
    AddRect p1, p2
    AddRectDim p1, p2, 100

    ' 螺絲孔
    r = 4 ' 半徑
    center(0) = s1 + length - v3 - v4: center(1) = s2 + v5
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r) ' 畫圓
    center(0) = s1 + length - v3 - v4: center(1) = s2 + width - v5
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r)

    r = 3.5 ' 半徑
    center(0) = s1 + v1 + v6 + 0.5 * v7: center(1) = s2 + v2 + v9 - r
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r) ' 畫圓
    center(0) = s1 + v1 + v6 + 0.5 * v7: center(1) = s2 + v2 + v9 + v8 + r
    Set cir_obj = ThisDrawing.ModelSpace.AddCircle(center, r)
End Sub


Public Sub electrical_box(ByVal length, ByVal width) ' 配電箱

    On Error Resume Next ' 如果有錯誤, 不管他


    ' 700 X 550 配電箱
    Dim start(2)  As Double ' 作圖起始點
    
    Dim p1(2)  As Double
    Dim p2(2)  As Double

    ' 刪除所有作圖
    For Each oEntity In ThisDrawing.ModelSpace
        oEntity.Delete
    Next

    start(0) = 1000: start(1) = 1000
    ' 配電箱 主體

    ' 配電箱 門
    door start, 663, 510
    ' 配電箱 內板
    inner_board start, 670, 450

    

    ' 配電箱 上板
    ' up_board

    ' ' 配電箱 底板
    ' bottom_board

    ' ' 配電箱 錶板
    gauge_board start, 580, 162

    ' 配電箱 配合機房箱墊板
    padding_board start, 884, 208


    


End Sub


Private Sub CommandButton1_Click()
    UserForm1.Hide

    length = Val(TextBox1.Text): width = Val(TextBox2.Text)
    electrical_box length, width
End Sub


Private Sub Label2_Click()

End Sub

Private Sub UserForm_Click()

End Sub





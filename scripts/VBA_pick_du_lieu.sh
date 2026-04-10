Sub copy_ten_sanpham()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    ' Lấy dòng cuối dựa vào cột G (Tên người mua) để bao phủ hết dữ liệu
    lastRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row
    
    Dim currentParent As String
    Dim i As Long
    Dim changedCount As Long
    Dim cellVal As String
    
    changedCount = 0
    currentParent = ""

    For i = 1 To lastRow
        ' Kiểm tra cột A và B
        Dim col As Integer
        For col = 1 To 2
            cellVal = Trim(CStr(ws.Cells(i, col).Value))
            
            ' KIỂM TRA DÒNG CHA: 
            ' Tìm ô có chứa chữ "Tên" hoặc "sản phẩm" (không phân biệt hoa thường)
            If UCase(cellVal) Like "*TÊN*" And UCase(cellVal) Like "*PHẨM*" Then
                currentParent = cellVal
                Exit For ' Thoát vòng lặp cột nếu đã tìm thấy cha
            End If
        Next col
        
        ' KIỂM TRA DÒNG CON: 
        ' Nếu cột A là số và chúng ta đã có tên sản phẩm cha
        If IsNumeric(ws.Cells(i, 1).Value) And ws.Cells(i, 1).Value <> "" Then
            If currentParent <> "" Then
                ' Ghi vào cột F
                ws.Cells(i, 6).Value = currentParent
                changedCount = changedCount + 1
            End If
        End If
    Next i
    
    If changedCount = 0 Then
        MsgBox "Vẫn chưa tìm thấy! Bạn hãy kiểm tra xem cột A có thực sự là số không, hoặc thử chọn vùng dữ liệu rồi chạy lại.", vbCritical
    Else
        MsgBox "Thành công! Đã cập nhật: " & changedCount & " ô.", vbInformation
    End If
End Sub
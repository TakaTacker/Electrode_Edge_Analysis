Attribute VB_Name = "EdgePeakAnalysis_v33"
Option Explicit

' 新規: 2ピーク取得 (peak1 + peak2)
Public Function Top2YInRange(ByRef xArr() As Double, ByRef yArr() As Double, _
                             ByVal xMin As Double, ByVal xMax As Double, ByVal minSep As Double, _
                             ByRef x1 As Double, ByRef y1 As Double, _
                             ByRef x2 As Variant, ByRef y2 As Variant) As Boolean
    Dim i As Long
    Dim found1 As Boolean
    Dim found2 As Boolean
    Dim yMax1 As Double
    Dim yMax2 As Double
    Dim xMax1 As Double
    Dim xMax2 As Double

    For i = LBound(xArr) To UBound(xArr)
        If xArr(i) >= xMin And xArr(i) <= xMax Then
            If Not found1 Or yArr(i) > yMax1 Then
                yMax1 = yArr(i)
                xMax1 = xArr(i)
                found1 = True
            End If
        End If
    Next i

    If Not found1 Then
        Top2YInRange = False
        Exit Function
    End If

    x1 = xMax1
    y1 = yMax1
    x2 = Empty
    y2 = Empty

    If minSep < 0 Then
        minSep = 0
    End If

    For i = LBound(xArr) To UBound(xArr)
        If xArr(i) >= xMin And xArr(i) <= xMax Then
            If Abs(xArr(i) - x1) >= minSep Then
                If Not found2 Or yArr(i) > yMax2 Then
                    yMax2 = yArr(i)
                    xMax2 = xArr(i)
                    found2 = True
                End If
            End If
        End If
    Next i

    If found2 Then
        x2 = xMax2
        y2 = yMax2
    End If

    Top2YInRange = True
End Function

' Result ヘッダ拡張 (既存A:Mの右側に追加)
Public Sub EnsureResultHeaders(ByVal ws As Worksheet)
    With ws
        If .Cells(1, 14).Value <> "x_L2_mm" Then
            .Cells(1, 14).Value = "x_L2_mm"
            .Cells(1, 15).Value = "yPeak_L2_um"
            .Cells(1, 16).Value = "h_L2_(y-baseline)/baseline"
            .Cells(1, 17).Value = "x_R2_mm"
            .Cells(1, 18).Value = "yPeak_R2_um"
            .Cells(1, 19).Value = "h_R2_(y-baseline)/baseline"
            .Cells(1, 20).Value = "PeakStatus"
        End If
    End With
End Sub

' 2ピーク結果の追記 (peak2が無い場合は空欄)
Public Sub AppendResultEx2(ByVal ws As Worksheet, ByVal rowIndex As Long, _
                           ByVal status As String, ByVal baseline As Double, _
                           ByVal xL1 As Double, ByVal yL1 As Double, _
                           ByVal xR1 As Double, ByVal yR1 As Double, _
                           ByVal xL2 As Variant, ByVal yL2 As Variant, _
                           ByVal xR2 As Variant, ByVal yR2 As Variant)
    Dim hL2 As Variant
    Dim hR2 As Variant

    If status <> "ERROR" Then
        If Not IsEmpty(xL2) And Not IsEmpty(yL2) Then
            hL2 = (CDbl(yL2) - baseline) / baseline
        Else
            hL2 = Empty
        End If

        If Not IsEmpty(xR2) And Not IsEmpty(yR2) Then
            hR2 = (CDbl(yR2) - baseline) / baseline
        Else
            hR2 = Empty
        End If

        ws.Cells(rowIndex, 14).Value = xL2
        ws.Cells(rowIndex, 15).Value = yL2
        ws.Cells(rowIndex, 16).Value = hL2
        ws.Cells(rowIndex, 17).Value = xR2
        ws.Cells(rowIndex, 18).Value = yR2
        ws.Cells(rowIndex, 19).Value = hR2
        ws.Cells(rowIndex, 20).Value = status
    Else
        ws.Cells(rowIndex, 14).Value = Empty
        ws.Cells(rowIndex, 15).Value = Empty
        ws.Cells(rowIndex, 16).Value = Empty
        ws.Cells(rowIndex, 17).Value = Empty
        ws.Cells(rowIndex, 18).Value = Empty
        ws.Cells(rowIndex, 19).Value = Empty
        ws.Cells(rowIndex, 20).Value = Empty
    End If
End Sub

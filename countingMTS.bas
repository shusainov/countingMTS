Attribute VB_Name = "Module1"
Sub CountingMTS()
Attribute CountingMTS.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' CountingMTS Macro
'
'

'
    Dim searchPosition As Integer
    
    Dim rng As Range, selectedRange As Range
    
    Dim summMinute As Double
    summMinute = 0
    Dim summMinuteCurency As Double
    summMinuteCurency = 0
    
    Dim summInternet As Double
    summInternet = 0
    Dim summInternetCurency As Double
    summInternetCurency = 0
    
    Set selectedRange = Selection
    
    For Each rng In selectedRange
    
'ѕодсчет минут
    searchPosition = InStr(rng, "минута")
    If (searchPosition > 0) And (rng.Cells(1, 2) > 0) Then
        rng.Cells(1, 3) = Left(rng, searchPosition - 2)
        rng.Cells(1, 4) = rng.Cells(1, 2) * 1.18
        summMinute = summMinute + rng.Cells(1, 3)
        summMinuteCurency = summMinuteCurency + rng.Cells(1, 4)
    End If
        
    searchPosition = InStr(rng, "секунда")
    If (searchPosition > 0) And (rng.Cells(1, 2) > 0) Then
        rng.Cells(1, 3) = CInt(Left(rng, searchPosition - 2)) / 60
        rng.Cells(1, 4) = rng.Cells(1, 2) * 1.18
        summMinute = summMinute + rng.Cells(1, 3)
        summMinuteCurency = summMinuteCurency + rng.Cells(1, 4)
    End If
    
'ѕодсчет интернета
    
    searchPosition = InStr(rng, "байт")
    If (searchPosition > 0) And (InStr(rng, "килобайт") = 0) Then
        rng.Cells(1, 3) = CDbl(Left(rng, searchPosition - 2)) / 1024 / 1024
        rng.Cells(1, 4) = rng.Cells(1, 2) * 1.18
        summInternet = summInternet + rng.Cells(1, 3)
        summInternetCurency = summInternetCurency + rng.Cells(1, 4)
    End If
       
    searchPosition = InStr(rng, "килобайт")
    If (searchPosition > 0) Then
        rng.Cells(1, 3) = CDbl(Left(rng, searchPosition - 2)) / 1024
        rng.Cells(1, 4) = rng.Cells(1, 2) * 1.18
        summInternet = summInternet + rng.Cells(1, 3)
        summInternetCurency = summInternetCurency + rng.Cells(1, 4)
    End If
    
    Next
    
    selectedRange.End(xlDown).Cells(2, 3) = summMinute
    selectedRange.End(xlDown).Cells(2, 4) = summMinuteCurency
    selectedRange.End(xlDown).Cells(3, 3) = summInternet
    selectedRange.End(xlDown).Cells(3, 4) = summInternetCurency
    selectedRange.End(xlDown).Cells(2, 5) = selectedRange.End(xlDown).Cells(1, 1) * 1.18
    Application.DisplayAlerts = False
    Range(selectedRange.End(xlDown).Cells(2, 5), selectedRange.End(xlDown).Cells(3, 5)).Merge
    Application.DisplayAlerts = True

End Sub


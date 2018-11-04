Attribute VB_Name = "Module1"
Sub print_unique()
Dim v

For Each ws In Worksheets

    ws.Range("J1") = "Ticker"
    ws.Range("K1") = "Total Stock Volume"

    v = getUniqueArray(ws.Range("A2", ws.Range("A" & Rows.Count).End(xlUp)))
    If IsArray(v) Then
    ws.Range("J2").Resize(UBound(v)) = v
   
    End If


    With ws.Range("K2")
    
        .Formula = "=SUMIF(A:A, J2, G:G)"
        .AutoFill Destination:=ws.Range("K2:K" & ws.Range("J" & Rows.Count).End(xlUp).Row), Type:=xlFillDefault
        
    End With

Next ws

End Sub



'Public is global
'Optional parameter: If it's not defined then it will take a value 'True'

Public Function getUniqueArray(inputRange As Range, _
                                Optional skipBlanks As Boolean = True, _
                                Optional matchCase As Boolean = True, _
                                Optional prepPrint As Boolean = True _
                                ) As Variant
               
'Object different datatype
'Variant is a variable type that is changing

Dim vDic As Object
Dim tArea As Range
Dim tArr As Variant, tVal As Variant, tmp As Variant
Dim noBlanks As Boolean
Dim cnt As Long
                      
'If an error happens go to the exit function

On Error GoTo exitFunc:
If inputRange Is Nothing Then GoTo exitFunc

'ReDim --> redefine

With inputRange
    If .Cells.Count < 2 Then
        ReDim tArr(1 To 1, 1 To 1)
        tArr(1, 1) = .Value2
        getUniqueArray = tArr
        GoTo exitFunc
    End If

    Set vDic = CreateObject("scripting.dictionary")
    If Not matchCase Then vDic.compareMode = vbTextCompare
    
    noBlanks = True
    
    For Each tArea In .Areas
        tArr = tArea.Value2
        For Each tVal In tArr
            If tVal <> vbNullString Then
                vDic.Item(tVal) = Empty
            ElseIf noBlanks Then
                noBlanks = False
            End If
        Next
    Next
End With

If Not skipBlanks Then If Not noBlanks Then vDic.Item(vbNullString) = Empty

'this is done just in the case of large data sets where the limits of
'transpose may be encountered
If prepPrint Then
    ReDim tmp(1 To vDic.Count, 1 To 1)
    For Each tVal In vDic.Keys
        cnt = cnt + 1
        tmp(cnt, 1) = tVal
    Next
    getUniqueArray = tmp
Else
    getUniqueArray = vDic.Keys
End If

exitFunc:
Set vDic = Nothing
End Function



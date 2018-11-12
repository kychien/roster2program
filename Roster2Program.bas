Attribute VB_Name = "Module2"
Sub ConvertProgram():
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Script to convert monster concert roster to program format.
    '
    '   2018 11 12 - Setup.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim songs() As String
    Dim comps() As String
    Dim pairs() As String
    Dim pos, pair As Integer
    Dim curr As String
    Dim x, y, size, times As Integer
    x = 1
    y = 4
    size = 1
    times = 0
    Dim time(4) As String
    curr = Cells(4, 1)
    
    'Figure out how many songs are being played
    Do While (Cells(y, 1) <> "End")
        If (Cells(y, 1) <> "") Then             'Song title found
            If (curr <> Cells(y, 1)) Then
                size = size + 1
                curr = Cells(y, 1)
            End If
        End If
        y = y + 1
    Loop
    
    'Figure out how many different times assuming a max of 4
    y = 4
    Do While (Cells(y, 3) <> "")
        If (IsInArr(Cells(y, 3), time) = False) Then
            time(times) = Cells(y, 3)
            times = times + 1
        End If
        y = y + 1
    Loop
    
    'Resize elements based on findings
    ReDim songs(size)
    ReDim comps(size)
    ReDim pairs(size * times, 17)
    pair = 1
    pos = 1
    x = 1
    y = 4
    
    'Copy Data to variables
    Do While (pos <= size)
        If (Cells(y, 1) <> "") Then         'New song was found
            songs(pos) = Cells(y, 1)        'Store song title
            comps(pos) = Cells(y, 2)        'Store song composer
            x = 4
            For i = 0 To times - 1            'Figure out which timeslot...
                If (Cells(y, 3) = "") Then  'Exit loop if no more times for that title
                    Exit For
                ElseIf (Cells(y, 3) = time(i)) Then
                    pair = 1
                    py = (i * size) + pos
                    Do While (Cells(y, x) <> "")    'Get all the player pairs at that row
                        pairs(py, pair) = Cells(y, x)
                        pair = pair + 1
                        x = x + 1
                    Loop
                    y = y + 1               'Assuming times are sorted, no need to alter i
                    x = 4                   ' in for loop
                End If
            Next i
            pos = pos + 1
        End If
          
        y = y + 1
    Loop
    
    'Format Data on new worksheet
    For i = 0 To times - 1                    'For each concert time...
        'Create the new workbook
        Dim newName As String
        newName = "Program" + CStr(i + 1)
        nameSheet (newName)
        
        'Fill out the data
        For j = 1 To size                   'For each song...
            y = 1 + (8 * (j - 1))
            
            Cells(y, 1) = songs(j)          'Place the title
            Cells(y, 3) = comps(j)          'Place the composer
            
            y = y + 1                       'Place all the pairs
            py = (i * times) + j
            pair = 1
            x = 1
            Do While (pairs(py, pair) <> "" And pair < 17)
                Cells(y, x) = pairs(py, pair)
                If (x < 3) Then
                    x = x + 1
                Else
                    y = y + 1
                    x = 1
                End If
                pair = pair + 1
                
            Loop
        Next j
        
        Sheets(newName).UsedRange.Columns.AutoFit

    Next i
End Sub

Function IsInArr(target As String, arr As Variant) As Boolean
    IsInArr = (UBound(Filter(arr, target)) > -1)
End Function

Function nameSheet(name As String)
    Dim ws As Worksheet
    With ThisWorkbook
        Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        ws.name = name
    End With
    ws.Activate
End Function

Attribute VB_Name = "Module2"
Sub ConvertProgram():
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Script to convert monster concert roster to program format.
    '
    '   2018 11 12 - Setup.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim curr As String
    Dim x, y, size As Integer
    x = 1
    y = 4
    size = 1
    curr = Cells(1, 4)
    
    'Figure out how many songs are being played
    Do While (Cells(1, y) <> "End")
        If (Cells(1, y) <> "") Then             'Song title found
            If (curr <> Cells(1, y)) Then
                size = size + 1
                curr = Cells(1, y)
            End If
        End If
        y = y + 1
    Loop
    
    'Resize elements based on findings
    Dim songs(size) As String
    Dim comps(size) As String
    Dim pairs(size, 21) As String
    Dim times(2) As String
    
    'Copy Data to variables
    
    
    'Format Data on new worksheet
    
    
            
End Sub


Attribute VB_Name = "StartGameModule"
Public GridNumber As String
Public Sub StartGame(Player As String, Enemy As String)
    
    Dim InvalidBoat As Boolean
    
    InvalidBoat = False
    
    'Player 1 setup
    Call PopulateLog(Player, InvalidBoat)
    Call FindBoats(Player, InvalidBoat)
    If InvalidBoat Then
    
        MsgBox "Not all boats are placed on the Grid"
        Range(Player & "LogIndicator").ClearContents
        Range(Player & "OurBoatRanges").ClearContents
        Range(Player & "EnemyBoatRanges").ClearContents
        Exit Sub
        
    End If
    
    Call FormatBoats(Player)
    
    'Player 2 setup
    Application.ScreenUpdating = False
    Call LoadComputerBoats
    Call PopulateLog(Enemy)
    Call FindBoats(Enemy)
    'Call FormatBoats(Enemy)
    Application.ScreenUpdating = True

End Sub

Private Sub PopulateLog(Player As String, Optional ByRef InvalidBoat As Boolean)

    Dim P1LI As Range, IndicatorCheck As Range
    Dim i As Integer
    
    Set P1LI = Range(Player & "LogIndicator")
    Set IndicatorCheck = Range(Player & "IndicatorCheck")
    
    'Grid contains 100 cells
    'counter for grid cell number
    i = 0

    For Each Col In Range(Player & "OurGrid").Columns
    
        For Each rw In Col.Rows
            
            i = i + 1
            
            P1LI(i) = UCase(rw.Value)

'            If Not IsEmpty(rw) Then
'                'assign ship letter from the grid to indicator log number
'                P1LI(i) = UCase(rw.Value)
'
'            End If
            
        Next rw
    
    Next Col
    
    'Check if all boats are placed minimum in one grid
    For Each IC In IndicatorCheck
        
        If IC <= 0 Then
            InvalidBoat = True
        
        End If
        
    Next IC
    
    Debug.Print "PopulateLog OK!"
End Sub

Private Sub FindBoats(Player As String, Optional ByRef InvalidBoat As Boolean)
    
    Dim BI As String, BN As String, BD As String
    Dim BS As Integer, r As Integer, a As Integer, b As Integer
    Dim GridFields As Variant

    GridFields = Array("", "C", "D", "B", "S", "PB")
    For Each rwboats In Range(Player & "Boats").Rows
        
        'declare boat configuration
        BI = rwboats.Cells(1, 1).Value   'Boat Indicator
        BN = rwboats.Cells(1, 2).Value   'Boat Name
        BS = rwboats.Cells(1, 3).Value   'Boat Size
        
        'Loop Our Grid section to find boats on it and check if they r placed correctly
        For Each Col In Range(Player & "OurGrid").Columns
        
            For Each rw In Col.Rows

                If BI = UCase(rw.Value) Then
                    
                    'check for direction and size
                    r = 1
                    a = 1
                    b = 1
                    BD = ""
                    
                    Do
                        If rw.Cells.Offset(r, 0).Value = rw.Value Then
                        
                            BD = "down"    'Boat Direction
                            a = a + 1
                            r = r + 1
                        ElseIf rw.Cells.Offset(0, r).Value = rw.Value Then
                            BD = "right"
                            b = b + 1
                            r = r + 1
                        Else
                            GoTo BoatInvalid
                        End If
                    
                    Loop Until r = BS
                    
                    If (a Or b = BS) And _
                        WorksheetFunction.CountIf(Range(Player & "LogIndicator"), BI) = BS Then
                            
                            'assign player 1 cell address according to direction of placed boat
                            If BD = "down" Then
                            
                                'BS_OurGrid = rwboats.Cells(1, 7)
                                BS_OurGrid = rw.Address & ":" & rw.Offset(BS - 1, 0).Address
                                rwboats.Cells(1, 7) = BS_OurGrid
                                                                
                                'BS_EnemyGrid = rwboats.Cells(1, 8)
                                BS_EnemyGrid = rw.Offset(0, 17).Address & ":" & rw.Offset(BS - 1, 17).Address
                                rwboats.Cells(1, 8) = BS_EnemyGrid
                                
                            Else
                            
                                BS_OurGrid = rw.Address & ":" & rw.Offset(0, BS - 1).Address
                                rwboats.Cells(1, 7) = BS_OurGrid
                                
                                BS_EnemyGrid = rw.Offset(0, 17).Address & ":" & rw.Offset(0, BS - 1 + 17).Address
                                rwboats.Cells(1, 8) = BS_EnemyGrid
                                                                
                            End If
                            
                            GoTo NextBoat
                    Else
                        GoTo BoatInvalid
                    End If
                   
                ElseIf Not IsInArray(UCase(rw.Value), GridFields) Then
                    GoTo GridError
                
                End If
                
            Next rw
        
        Next Col
       
NextBoat:
    Next rwboats
    Debug.Print "FindBoat OK!"
Exit Sub
    
BoatInvalid:
    InvalidBoat = True
    MsgBox "Boat " & BN & " placed incorrectly!"
    Exit Sub
    
GridError:
    InvalidBoat = True
    MsgBox "There is placed boat, that is not recognizable by our navy"
    Exit Sub

End Sub
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Private Sub FormatBoats(Player As String)

    Dim BRange As Range
    Dim BColor As Integer
    Dim P1Grid As Worksheet
    
    Set P1Grid = ThisWorkbook.Sheets(Player)
    
    For Each rwboats In Range(Player & "Boats").Rows
    
        Set BRange = rwboats.Cells(1, 7)  'Boat range on grid in excel spreadsheet range values
        BColor = rwboats.Cells(1, 4).Value  'Boat Color
        
        P1Grid.Range(BRange).Interior.ColorIndex = BColor
        P1Grid.Range(BRange).BorderAround xlContinuous, xlMedium
    
    Next rwboats
    Debug.Print "FormatBoats OK!"
End Sub

Sub LoadComputerBoats()

    Dim FolderPath As String
    FolderPath = "I:\excel\VBA\Project Based Excel VBA Course\Battleship\Enemy Grids\"

    Dim FileName As String, EGVariant As String
    Dim WS As Worksheet
    Dim WB As Workbook
    Dim EGRange As Range
    
    EGVariant = ""
    GridNumber = ""
    
    Application.ScreenUpdating = False
    
    Set WS = ThisWorkbook.Sheets("Player2")
    Set EGRange = WS.Range("Player2OurGrid")
    
    'Choose random grid from presets or choose a specific one
    Do Until IsNumeric(EGVariant)
        GridVariant = MsgBox("Do you want to choose Enemy Grid variant?", vbYesNo)
 
        If GridVariant = 6 Then
        
            GridVariantForm.Show
            EGVariant = GridNumber
            'EGVariant = GridVariantForm.GridNumber
        
        Else
        
            'Enemy Grid variant number
            EGVariant = Format([RandBetween(1, 6)], "00")
        
        End If
    Loop
    
    FileName = Dir(FolderPath & "*" & EGVariant & ".xls*")
    
    'Open workbook with enemy grid, copy drip and paste it to game
    Workbooks.Open (FolderPath & FileName)
    Set WB = Workbooks(FileName)
    
    WB.Sheets("Grid").Range("Player1OurGrid").Copy
    EGRange.PasteSpecial Paste:=xlPasteAll
    
    'Close workbook
    WB.Close

End Sub

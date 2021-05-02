Attribute VB_Name = "NewGameModule"
Public Sub NewGame()
    
    'Declare variables
    Dim gridP1o As Range
    Dim GridP1E As Range
    Dim GridP2O As Range
    Dim GridP2E As Range
    Dim Answer As String

    'Message box for starting a new game.
    Answer = MsgBox("Would you like to start a New Game?", vbQuestion + vbYesNo + vbDefaultButton2, "New Game")
    If Answer = vbNo Then
        Exit Sub
    End If

    'Clear contents of the Indicator, Range (Our Grid) and Range (Enemy Grid) columns for Player 1 log.
    Range("Player1LogIndicator").ClearContents
    Range("Player1OurBoatRanges").ClearContents
    Range("Player1EnemyBoatRanges").ClearContents
    
    'Set Attacked Flag range values to 0 for Player 1
    Range("Player1LogAttackedFlag").Value = 0
    
    'Clear contents of the Indicator, Range (Our Grid) and Range (Enemy Grid) columns for Player 2 log.
    Range("Player2LogIndicator").ClearContents
    Range("Player2OurBoatRanges").ClearContents
    Range("Player2EnemyBoatRanges").ClearContents
    
    'Set Attacked Flag range values to 0 for Player 2
    Range("Player2LogAttackedFlag").Value = 0
    
    'Clear and format players grids.
    
    Set gridP1o = Range("Player1OurGrid")
    Set GridP1E = Range("Player1EnemyGrid")
    Set GridP2O = Range("Player2OurGrid")
    Set GridP2E = Range("Player2EnemyGrid")
    ClearFormatGrid gridP1o
    ClearFormatGrid GridP1E
    ClearFormatGrid GridP2O
    ClearFormatGrid GridP2E



End Sub


Sub ClearFormatGrid(Grid)
'Clear and format players grids.

    Grid.Clear
    Grid.Font.Italic = True
    Grid.Font.Size = 16
    Grid.BorderAround xlContinuous, xlMedium
    Grid.Borders(xlInsideHorizontal).LineStyle = xlDot
    Grid.Borders(xlInsideVertical).LineStyle = xlDot
    Grid.Cells.VerticalAlignment = xlCenter
    Grid.Cells.HorizontalAlignment = xlCenter

End Sub






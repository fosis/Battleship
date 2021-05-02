Attribute VB_Name = "FireshotModule"
Private Sub DoTurn()

    Dim ComputerTarget As String, Enemy As String, Player As String, SunkAddress As String, TargetBoardID As String
    Dim Hit As Boolean, Sunk As Boolean, EndGame As Boolean
    Dim BoatColor As Integer
    
    Enemy = "Player1"
    Player = "Player2"
    ComputerTarget = ""
    TargetBoardID = ""
    BoatIndicator = ""
    Call FindTarget_Computer(TargetBoardID, BoatIndicator, ComputerTarget, Hit, Sunk, SunkAddress, BoatColor)
    
    Call ChangeGridFormat(ComputerTarget, Player, Enemy, Hit, Sunk, SunkAddress, BoatColor)
    
    Call HitMessage(Player, Enemy, TargetBoardID, BoatIndicator, Hit, Sunk)
    
    Debug.Print ComputerTarget
    
    Call CheckEndGame(Player, Enemy, EndGame)
    If EndGame Then
        MsgBox Player & "has Won!" & vbNewLine & "All your ships has been sunk"
        MsgBox "Press New Game button to play again!"
        Exit Sub
    End If
    MsgBox "It's your turn now." & vbNewLine & "Please choose target and press Fire! button"

End Sub

Public Sub FireShot()

    Dim Player As String, Enemy As String, TargetBoardID As String, BoatIndicator As String
    Dim Target As Range
    Dim TargetCheck As Boolean
    Dim Hit As Boolean, Sunk As Boolean, EndGame As Boolean
    
    TargetCheck = True
    Set Target = Application.Selection.Resize(1, 1)
    Player = "Player1"
    Enemy = "Player2"
    TargetBoardID = ""
    BoatIndicator = ""
    
    Call FindTarget_Player(Target, TargetBoardID, BoatIndicator, TargetCheck, Hit, Sunk)
    
    'If Target is not in Enemy Grid do not continue
    If TargetCheck = False Then
        Exit Sub
    End If
    
    Call HitMessage(Player, Enemy, TargetBoardID, BoatIndicator, Hit, Sunk)
    
    Call CheckEndGame(Player, Enemy, EndGame)
    If EndGame Then
        MsgBox Player & "has Won!" & vbNewLine & "You have sunk all enemy ships"
        MsgBox "Press New Game button to play again!"
        Exit Sub
    End If
    
    Call DoTurn
    
End Sub

Private Sub CheckEndGame(Player As String, Enemy As String, Optional ByRef EndGame As Boolean)

    'Checks if game has ended
    Dim BoatsRemaining As Range
    
    Set BoatsRemaining = Sheets(Enemy & " Log").Range(Enemy & "BoatsRemaining")
    
    If BoatsRemaining <= 0 Then
    
        EndGame = True
        Sheet2.FireButton.Visible = False
    
    End If
End Sub

Private Sub HitMessage(Player As String, Enemy As String, TargetBoardID As String, Optional BoatIndicator, Optional Hit, Optional Sunk)

    'Declare variables
    Dim BoatName As String, GridAddress As String
    Dim P1LOG As Worksheet, P2LOG As Worksheet
    
    'Assign variables
    Set P1LOG = Sheet3
    Set P2LOG = Sheet5
    
    If Enemy = "Player2" Then
        With P2LOG
            If BoatIndicator <> "" Then
                BoatName = WorksheetFunction.VLookup(BoatIndicator, Range(Enemy & "Boats"), 2, 0)
            End If
            
            GridAddress = WorksheetFunction.VLookup(TargetBoardID, Range(Enemy & "Log"), 6, 0)
            
        End With
    Else
        With P1LOG
        
            If BoatIndicator <> "" Then
                BoatName = WorksheetFunction.VLookup(BoatIndicator, Range(Enemy & "Boats"), 2, 0)
            End If
                
            GridAddress = WorksheetFunction.VLookup(TargetBoardID, Range(Enemy & "Log"), 6, 0)
        
        End With
    End If

    If Player = "Player1" Then
        'Message for Player 1
        If Not Hit Then
            'miss message
            MsgBox "You shot at " & GridAddress & vbNewLine & vbNewLine & " and..." & vbNewLine & vbNewLine & "It's Miss!"
        ElseIf Hit And Not Sunk Then
            'hit message
            MsgBox "You shot at " & GridAddress & vbNewLine & vbNewLine & " and..." & vbNewLine & vbNewLine & "It's Hit!"
        Else 'sunkw
            'sunk message
            MsgBox "You shot at " & GridAddress & vbNewLine & vbNewLine & " and..." & vbNewLine & vbNewLine & "It's Hit!" & vbNewLine & vbNewLine & "You sunk the enemy " & BoatName
        
        End If
        
    Else
        'Message for Player 2
        If Not Hit Then
            'miss message
            MsgBox "The computer fires at " & GridAddress & " and misses"
        ElseIf Hit And Not Sunk Then
            'hit message
            MsgBox "The computer fires at " & GridAddress & " and hits your " & BoatName
        Else 'sunk
            'sunk message
            MsgBox "The computer fires at " & GridAddress & " and sinks your " & BoatName
        
        End If
        
    End If
End Sub
Private Function FindTarget_Player(TargetCell As Range, ByRef TargetBoardID, Optional ByRef BoatIndicator, _
                                Optional ByRef TargetCheck As Boolean, Optional ByRef Hit As Boolean, _
                                Optional ByRef Sunk As Boolean) _
                                As Range

    'Check if selected cell is in EnemyGrid range
    Dim EnemyGrid As Range, CorrectTarget As Range, EnemyLog As Range, EnemyAF As Range, EnemyLogInd As Range, EnemyBoats As Range
    Dim BoardIDColumn As String, BoardIDRow As String, Player As String, Enemy As String, SunkAddress As String
    Dim BoatColor As Integer
    
    Set EnemyGrid = Range("Player1EnemyGrid")
    Set EnemyLog = Range("Player2Log")
    Set EnemyAF = Range("Player2LogAttackedFlag")
    Set EnemyLogInd = Range("Player2LogIndicator")
    Set EnemyBoats = Range("Player2Boats")
    
    Player = "Player1"
    Enemy = "Player2"
    Hit = False
    
    Set CorrectTarget = Application.Intersect(TargetCell, EnemyGrid)
    If CorrectTarget Is Nothing Then
        TargetCheck = False
        MsgBox "Target cell not in Enemy Grid." & vbNewLine & "Please choose cell in Enemy Grid."
        GoTo InvalidTarget
    End If
    
    'Change excel cell address to match BoardID format
    BoardIDColumn = WorksheetFunction.Text(CorrectTarget.Column - 19, "00")
    BoardIDRow = WorksheetFunction.Text(CorrectTarget.Row - 3, "00")
    TargetBoardID = BoardIDColumn & BoardIDRow

    AttFlagRow = WorksheetFunction.Match(TargetBoardID, EnemyLog.Columns(1).Value2, 0)
    
        'Check if cell has already been attacked
        If EnemyAF.Rows(AttFlagRow).Value = 1 Then
        
            TargetCheck = False
            MsgBox "Target cell has already been attacked!" & vbNewLine & "Please choose another cell in Enemy Grid."
            GoTo InvalidTarget
            
        'Change status of Attacked flag
        Else
            
            'Check if you hit ship
            If EnemyLogInd.Rows(AttFlagRow).Value <> "" Then
            
                Hit = True
                BoatIndicator = EnemyLogInd.Rows(AttFlagRow).Value
                EnemyAF.Rows(AttFlagRow).Value = 1
                
                'Check if sunk
                If WorksheetFunction.VLookup(BoatIndicator, EnemyBoats, 6, 0) = 1 Then
                    Sunk = True
                    SunkAddress = WorksheetFunction.VLookup(BoatIndicator, EnemyBoats, 8, 0)
                    BoatColor = WorksheetFunction.VLookup(BoatIndicator, EnemyBoats, 4, 0)
                    Call ChangeGridFormat(CorrectTarget.Address, Player, Enemy, Hit, Sunk, SunkAddress, BoatColor)
                Else
                    Sunk = False
                    Call ChangeGridFormat(CorrectTarget.Address, Player, Enemy, Hit)
                End If
                
            Else
                
                Hit = False
                EnemyAF.Rows(AttFlagRow).Value = 1
                Call ChangeGridFormat(CorrectTarget.Address, Player, Enemy)
            
            End If
            
        End If
    
    Set FindTarget_Player = CorrectTarget
    
    Debug.Print "Target OK!"
InvalidTarget:
End Function

Function ChangeGridFormat(GridCell As String, _
                            Player As String, _
                            Enemy As String, _
                            Optional Hit As Boolean, _
                            Optional Sunk As Boolean, _
                            Optional SunkAddress As String, _
                            Optional BoatColor As Integer _
                            )
    Dim OffsetGridCell As Range

    Sheets(Player).Range(GridCell).Borders(xlDiagonalDown).LineStyle = xlContinuous
    Sheets(Player).Range(GridCell).Borders(xlDiagonalUp).LineStyle = xlContinuous
    Sheets(Player).Range(GridCell).Borders(xlDiagonalDown).Weight = xlMedium
    Sheets(Player).Range(GridCell).Borders(xlDiagonalUp).Weight = xlMedium
    
    Set OffsetGridCell = Range(GridCell).Offset(0, -17)
    Sheets(Enemy).Range(OffsetGridCell.Address).Borders(xlDiagonalDown).LineStyle = xlContinuous
    Sheets(Enemy).Range(OffsetGridCell.Address).Borders(xlDiagonalUp).LineStyle = xlContinuous
    Sheets(Enemy).Range(OffsetGridCell.Address).Borders(xlDiagonalDown).Weight = xlMedium
    Sheets(Enemy).Range(OffsetGridCell.Address).Borders(xlDiagonalUp).Weight = xlMedium
    
    If Hit Then
        
        Sheets(Player).Range(GridCell).Interior.ColorIndex = 3
        Sheets(Enemy).Range(OffsetGridCell.Address).Interior.ColorIndex = 3
    
    End If
    
    If Sunk Then
    
        Sheets(Player).Range(SunkAddress).BorderAround xlContinuous, xlThick, ColorIndex:=BoatColor
    
    End If

End Function
Function FindTarget_Computer(ByRef TargetBoardID, Optional ByRef BoatIndicator, Optional ByRef ComputerTarget, _
                            Optional ByRef Hit, Optional ByRef Sunk, Optional ByRef SunkAddress, _
                            Optional BoatColor)

    Dim HuntingRow As Range, TargetCellGrid As Range, EnemyBoats As Range
    Dim RandomTargetArray() As Variant, AvCell As Variant
    Dim LowerBound As Integer, UpperBound As Integer
    Dim TargetCellID As String
    Dim TColumn As Double, TRow As Double
    Dim WS As Excel.Worksheet
    Dim x As Integer, i As Integer
    
    Set WS = ActiveWorkbook.Worksheets("Player1 Log")
    Set HuntingRow = WS.Range("Player1HuntingRow")
    Set EnemyBoats = Range("Player1Boats")
    
    With WS
        If HuntingRow > 0 Then
            'Hunt the boat
            LowerBound = 1
            UpperBound = .Range("Player1AvailableAttacks")
            
            RandomNum = Int(LowerBound + Rnd * (UpperBound - LowerBound - 1))
            'Size of array equals to number of available attacks
            ReDim RandomTargetArray(UpperBound - 1)
    
            x = 0
            
            AvCell = WS.Range("Player1Hunting").Rows(HuntingRow)
            
            For i = 6 To 9
            
                If AvCell(1, i) <> "N/A" And AvCell(1, i) <> "Attacked" Then

                    RandomTargetArray(x) = AvCell(1, i + 4)
                    x = x + 1
                    
                End If
            
            Next
            
'            For Each AvCell In ws.Range("Player1Hunting").Rows(HuntingRow) '.Range(Cells(1, 6), Cells(1, 9))
'
'                'Check if cell has N/A status, if no, add it to the array
'                If AvCell <> "N/A" And AvCell <> "Attacked" Then
'
'                    RandomTargetArray(x) = AvCell.Offset(0, 4)
'                    x = x + 1
'                End If
'
'            Next AvCell
            
            'BoardID to attack by computer
            TargetCellID = RandomTargetArray(RandomNum)
            'Row number of AttackedFlag range
            AttFlagRow = WorksheetFunction.Match(TargetCellID, Range("Player1Log").Columns(1).Value2, 0)
            'Set value to 1
            Debug.Print AttFlagRow
            Range("Player1LogAttackedFlag").Rows(AttFlagRow).Value = 1
        
        Else
            'Randomly choose unattacked cells
            LowerBound = 1
            UpperBound = Range("Player1UnattackedCells")
            
            RandomNum = Int(LowerBound + Rnd * (UpperBound - LowerBound - 1))
            
            'Size of array equals to number of unattacked cells
            ReDim RandomTargetArray(UpperBound - 1)
                
            x = 0
            For Each AvCell In Range("Player1Log").Rows
    
                'Check if cell has changed flag, if no, add it to the array
                If AvCell.Cells(, 5) = 0 Then
    
                    RandomTargetArray(x) = AvCell
                    x = x + 1
    
                End If
    
            Next AvCell
            
            'BoardID to attack by computer
            TargetCellID = RandomTargetArray(RandomNum)(1, 1)
            'Row number of AttackedFlag range
            AttFlagRow = WorksheetFunction.Match(TargetCellID, Range("Player1Log").Columns(1).Value2, 0)
            'Set value to 1
            Debug.Print AttFlagRow
            Range("Player1LogAttackedFlag").Rows(AttFlagRow).Value = 1
    
        End If
        
        'Check if hit
        BoatIndicator = Range("Player1LogIndicator").Rows(AttFlagRow).Value
        If BoatIndicator <> "" Then
            Hit = True
            
            'Check if sunk
            If WorksheetFunction.VLookup(BoatIndicator, EnemyBoats, 6, 0) = 1 Then
            
                SunkAddress = WorksheetFunction.VLookup(BoatIndicator, EnemyBoats, 8, 0)
                BoatColor = WorksheetFunction.VLookup(BoatIndicator, EnemyBoats, 4, 0)
                Sunk = True
                
            End If
        End If
        
    End With
        'Transform BoardId to Grid cell address
        TColumn = Left(TargetCellID, 2)
        TRow = Right(TargetCellID, 2)
        Set TargetCellGrid = Cells(TRow + 3, TColumn + 2 + 17)
        
        TargetBoardID = TargetCellID
    
        ComputerTarget = TargetCellGrid.Address
'Debug.Print Hit

End Function

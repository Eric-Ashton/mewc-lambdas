Attribute VB_Name = "maze_solver_color"
' deploy: shared
'=========================================
' Sub: maze_solver_color_with_diagonal
'=========================================
' Solves a maze using breadth-first search logic starting from a cell with value 0,
' propagating outward to all 8 neighboring cells (including diagonals)
' that share the same background color.
'
' Each newly reached cell is filled with a number representing the number of steps from the starting cell.
' Filling stops when no further changes can be made.
'
' SELECTION REQUIRED:
'   A rectangular range of cells representing a maze.
'   - One of the cells must contain the value 0 (this is the starting point).
'   - All traversable paths should share the same background fill color as the start cell.
'   - Walls/obstacles must have a different fill color.

Sub maze_solver_color_with_diagonal()
    Dim mazeRange As Range
    Dim cell As Range
    Dim neighborCell As Range
    Dim startColor As Long
    Dim directionOffsets As Variant
    Dim targetRow As Long, targetCol As Long
    Dim currentNumber As Long
    Dim changesMade As Boolean
    
    ' Get the selected range and assign to mazeRange
    On Error Resume Next
    Set mazeRange = Selection
    On Error GoTo 0
    If mazeRange Is Nothing Then Exit Sub ' Exit if no range is selected

    ' Find the fill color of the first non-empty cell with a value of 0
    startColor = -1 ' Default to -1 (no start color found)
    For Each cell In mazeRange
        If Not IsEmpty(cell.Value) And cell.Value = 0 Then
            startColor = cell.Interior.Color
            Exit For
        End If
    Next cell

    If startColor = -1 Then Exit Sub ' Exit if no valid cell with 0 is found

    ' Define neighbor offsets (8 directions)
    directionOffsets = Array(Array(-1, 0), Array(1, 0), Array(0, -1), Array(0, 1), Array(-1, 1), Array(1, -1), Array(1, 1), Array(-1, -1))
    
    ' Start filling process
    currentNumber = 1
    Do
        changesMade = False
        
        ' Iterate through the maze range
        For Each cell In mazeRange
            If (IsEmpty(cell.Value) Or cell.Value = "") And cell.Interior.Color = startColor Then
                Dim hasNeighbor As Boolean
                hasNeighbor = False
                
                For Each Offset In directionOffsets
                    targetRow = cell.Row + Offset(0)
                    targetCol = cell.Column + Offset(1)
                    
                    ' Ensure the target cell coordinates are within the worksheet's valid range (1 or greater)
                    ' This prevents the "Application-defined" error before we try to Set the cell object.
                    If targetRow > 0 And targetCol > 0 Then
                        Set neighborCell = mazeRange.Worksheet.Cells(targetRow, targetCol)
                        
                        ' Confinement Check: Ensure the neighbor is actually within the selected mazeRange
                        If Not Intersect(neighborCell, mazeRange) Is Nothing Then
                            ' Original Logic: Check if this valid neighbor has the required number
                            If Not IsEmpty(neighborCell.Value) And IsNumeric(neighborCell.Value) And neighborCell.Value = currentNumber - 1 Then
                                hasNeighbor = True
                                Exit For ' Found a valid neighbor, no need to check others
                            End If
                        End If
                    End If
                Next Offset
                
                ' If a valid neighbor was found, set the next number
                If hasNeighbor Then
                    cell.Value = currentNumber
                    changesMade = True
                End If
            End If
        Next cell
        
        ' Increment to the next number
        currentNumber = currentNumber + 1
    Loop While changesMade
End Sub


'=========================================
' Sub: maze_solver_color_no_diagonal
'=========================================
' Solves a maze using breadth-first search logic starting from a cell with value 0,
' propagating outward to only 4 neighboring cells (up, down, left, right)
' that share the same background color.
'
' Each newly reached cell is filled with a number representing the number of steps from the starting cell.
' Filling stops when no further changes can be made.
'
' SELECTION REQUIRED:
'   A rectangular range of cells representing a maze.
'   - One of the cells must contain the value 0 (this is the starting point).
'   - All traversable paths should share the same background fill color as the start cell.
'   - Walls/obstacles must have a different fill color.
'
' Notes:
'   - This version uses 4-directional movement (no diagonals).

Sub maze_solver_color_no_diagonal()
    Dim mazeRange As Range
    Dim cell As Range
    Dim neighborCell As Range
    Dim startColor As Long
    Dim directionOffsets As Variant
    Dim targetRow As Long, targetCol As Long
    Dim currentNumber As Long
    Dim changesMade As Boolean
    
    ' Get the selected range and assign to mazeRange
    On Error Resume Next
    Set mazeRange = Selection
    On Error GoTo 0
    If mazeRange Is Nothing Then Exit Sub ' Exit if no range is selected

    ' Find the fill color of the first non-empty cell with a value of 0
    startColor = -1 ' Default to -1 (no start color found)
    For Each cell In mazeRange
        If Not IsEmpty(cell.Value) And cell.Value = 0 Then
            startColor = cell.Interior.Color
            Exit For
        End If
    Next cell

    If startColor = -1 Then Exit Sub ' Exit if no valid cell with 0 is found

    ' *** CHANGE: Define neighbor offsets for 4 directions only (up, down, left, right) ***
    directionOffsets = Array(Array(-1, 0), Array(1, 0), Array(0, -1), Array(0, 1))
    
    ' Start filling process
    currentNumber = 1
    Do
        changesMade = False
        
        ' Iterate through the maze range
        For Each cell In mazeRange
            If (IsEmpty(cell.Value) Or cell.Value = "") And cell.Interior.Color = startColor Then
                Dim hasNeighbor As Boolean
                hasNeighbor = False
                
                For Each Offset In directionOffsets
                    targetRow = cell.Row + Offset(0)
                    targetCol = cell.Column + Offset(1)
                    
                    ' Boundary Check: Ensure the target cell coordinates are within the worksheet's valid range
                    If targetRow > 0 And targetCol > 0 Then
                        Set neighborCell = mazeRange.Worksheet.Cells(targetRow, targetCol)
                        
                        ' Confinement Check: Ensure the neighbor is actually within the selected mazeRange
                        If Not Intersect(neighborCell, mazeRange) Is Nothing Then
                            ' Logic: Check if this valid neighbor has the required number
                            If Not IsEmpty(neighborCell.Value) And IsNumeric(neighborCell.Value) And neighborCell.Value = currentNumber - 1 Then
                                hasNeighbor = True
                                Exit For ' Found a valid neighbor, no need to check others
                            End If
                        End If
                    End If
                Next Offset
                
                ' If a valid neighbor was found, set the next number
                If hasNeighbor Then
                    cell.Value = currentNumber
                    changesMade = True
                End If
            End If
        Next cell
        
        ' Increment to the next number
        currentNumber = currentNumber + 1
    Loop While changesMade
End Sub

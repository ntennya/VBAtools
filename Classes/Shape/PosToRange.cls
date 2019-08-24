VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PosToRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'** PosToRange                                                     *************
'*******************************************************************************
'**
'** -- Explanation -------------------------------------------------------------
'**
'** This class enable to get the nearest cell position with coordinate values.
'**
'** -- Methods -----------------------------------------------------------------
'**
'** getOptimizedRow(top): Returns row index of cell that include the argument
'**                       "top" as Y-axis.
'**
'** getOptimizedCol(left): Returns column index of cell that include the argument
'**                        "left" as X-axis.
'**
'** getOptimizedCellPosition(top,left):
'**        Returns Range object(1 cells) that include the coordinate:(left,top).
'**
'** -- Example of use ----------------------------------------------------------
'**
'**  Dim ptrng = New PosToRange
'**  ptrng.init Sheet1
'**  Dim rng as Range
'**  Set rng = ptrng.getOptimizedCellPosition(300,200)  ...Will get the cell
'**                                                     that locates (200,300).
'**
'**
'** version 0.0.0                          Created by ntennya(GitHub user name)
'*******************************************************************************
Option Explicit

Private Const ROWMAX As Long = 10000
Private Const COLMAX = 1000

Private Mws As Worksheet

'*******************************************************************************
'*** Constructor ***************************************************************
'*******************************************************************************
Public Function init(ws As Worksheet)
  Set Mws = ws
End Function

'*******************************************************************************
'*** Public Methods ************************************************************
'*******************************************************************************
'=== Getter ================================================
Public Function getOptimizedRow(top As Long) As Long
  Dim l As Long: l = 0
  Dim row As Long
  For row = 1 To ROWMAX
    l = l + Mws.Cells(row, 1).Height
    If l >= top Then
      getOptimizedRow = row
      Exit Function
    End If
  Next row
  
  getOptimizedRow = -1
End Function


Public Function getOptimizedCol(left As Long) As Long
  Dim l As Long: l = 0
  Dim col As Long
  For col = 1 To COLMAX
    l = l + Mws.Cells(1, col).Width
    If l >= left Then
      getOptimizedCol = col
      Exit Function
    End If
  Next col
  
  getOptimizedCol = -1
End Function


Public Function getOptimizedCellPosition(top As Long, left As Long, Optional rowOffset As Long = 0, Optional colOffset As Long = 0) As Range
  Set getOptimizedCellPosition = Mws.Cells(getOptimizedRow(top), getOptimizedCol(left)).Offset(rowOffset, colOffset)
End Function

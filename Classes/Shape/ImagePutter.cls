VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImagePutter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'** ImagePutter                                                    *************
'*******************************************************************************
'**
'** -- Need Modules/Classes ----------------------------------------------------
'** *PosToRange
'**
'** -- Explanation -------------------------------------------------------------
'** This Class enable to put images with setting name by three way.
'** (absolute coordinate, nearest cell of coordinate, row/col of cell)
'**
'** -- Methods -----------------------------------------------------------------
'** Common arguments:
'**    imageFP : The fullPath of the image file for put.
'**    name    : The name for the put image. The opportunity is to identify
'**              this image in VBA.
'**   rowOffset
'**   colOffset: The row/column offset of Cell position.
'**
'** putImage(imageFP, name, top, left):
'**   Puts a image to the absolute position (left,top)
'**
'** putImageToOptimizedCell(imageFP, name, top, left, rowOffset, colOffset):
'**   Puts a image to the nearest position of (left,top)
'**
'** putImageToCell(imageFP, name, row, col, rowOffset, colOffset):
'**   Puts a image to Cells(row,col).
'**
'** -- Example of use ----------------------------------------------------------
'**
'**  With New ImagePutter
'**    .init Sheet1
'**    .putImageToOptimizedCell "C:\hoge\fuga.jpg" ,"FUGA", 300, 200
'**  End With
'**
'**                 ... Will put image "C:\hoge\fuga.jpg" to the nearest
'**                     cell of (200,300) with naming it "FUGA".
'**
'** version 0.0.0                          Created by ntennya(GitHub user name)
'*******************************************************************************
Option Explicit

Private Mpos As Postorange
Private Mws As Worksheet

'*******************************************************************************
'*** Constructor ***************************************************************
'*******************************************************************************
Public Function init(ws As Worksheet)
  Set Mws = ws
  Set Mpos = New Postorange
  Mpos.init ws
End Function

'*******************************************************************************
'*** Public Methods ************************************************************
'*******************************************************************************
'=== Setter ================================================
Public Function putImage(imageFP As String, name As String, top As Long, left As Long)
With Mws.Pictures.Insert(imageFP)
  .top = top
  .left = left
  .name = name
End With
End Function


Public Function putImageToOptimizedCell(imageFP As String, name As String, top As Long, left As Long, Optional rowOffset As Long = 0, Optional colOffset As Long = 0)
With Mws.Pictures.Insert(imageFP)
  Dim rng As Range: Set rng = Mpos.getOptimizedCellPosition(top, left, rowOffset, colOffset)
  .top = rng.top
  .left = rng.left
  .name = name
End With
End Function


Public Function putImageToCell(imageFP As String, name As String, row As Long, col As Long, Optional rowOffset As Long = 0, Optional colOffset As Long = 0)
With Mws.Pictures.Insert(imageFP)
  Dim rng As Range: Set rng = Mws.Cells(row, col).Offset(rowOffset, colOffset)
  .top = rng.top
  .left = rng.left
  .name = name
End With
End Function

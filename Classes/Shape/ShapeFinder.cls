VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapeFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'** ShapeFinder                                                    *************
'*******************************************************************************
'**
'** -- Explanation -------------------------------------------------------------
'** This Class enale to find a shape in a worksheet by name.
'** Please use "ImagePutter" class if you will put images to worksheet that
'** forces to set a name to image object.
'** -- Methods -----------------------------------------------------------------
'**  getShape(name) : Find the shape that has the name same as argument.
'**
'** -- Example of use ----------------------------------------------------------
'**
'**  With New ShapeFinder
'**    .init Sheet1
'**    Dim cat as Shape
'**    Set cat = .getShape("TOM")
'**  End Dith
'**                    ... Will get a pointer for the shape that name is "TOM".
'**                    If There is no shape that name is "TOM" in Sheet1,
'**                    This function will return "Nothing"(= "null" in VBA).
'**
'** version 0.0.0                          Created by ntennya(GitHub user name)
'*******************************************************************************
Option Explicit

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
Public Function getShape(name As String) As Shape
  Dim shp As Shape
  For Each shp In Mws.Shapes
    If shp.name = name Then
      Set getShape = shp
      Exit Function
    End If
  Next
  
  Set getShape = Nothing   'No work, but need for readable.
End Function




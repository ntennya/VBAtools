VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'** ArrayList                                                      *************
'*******************************************************************************
'** Target user : Intermediate to Expert programmers.
'** -- Need Modules/Classes ----------------------------------------------------
'**
'**
'** -- Explanation -------------------------------------------------------------
'** This is Dynamic array. The size will be expanded(Redim Preserve)
'** "EXPANDFREQ" on demmand.
'**
'**  For example...
'**   Size of Marr: 11(Ubound(Marr)=10), EXPANDFREQ=10
'**   If the element count changed from 11 to 12, the Marr size will change
'**   from 11 to 21 (+10).
'**
'**  Attention...
'**  (1) EXPANDFREQ should be changed by the forecast of size.
'**   If the forecast of element number is too big... such as 10000, please
'**   adjust the coustant "EXPANDFREQ" to big. Because of, the "Redim preserve"
'**   method will wast the memory space.
'**
'**  (2) The differance between "value(Variant)" and "object variable".
'**   In VBA, the pointer that is called "Object Variable" needs special treatment.
'**   For example, if assign a object variable to an object, it needs "Set"
'**   statement.
'**
'**    (NG) Dim pointer : pointer = new Tuple
'**    (OK) Dim pointer : Set pointer = new Tuple
'**
'**   So, the method will be used properly. For example,
'**    "add", "getValue" methods are for value elements.
'**    "addObject", "getObject" methods are for object elements.
'**
'** -- Methods -----------------------------------------------------------------
'**
'** add/AddObject      : Add a value or object.
'** getValue/getObject : Get a value or object.
'** count              : Get the element count.
'** contains           : Will return True if the array has the value.
'** appendValueList
'** appendValueList    : Append another ArrayList from next index.
'**
'** -- Example of use ----------------------------------------------------------
'**
'**  Dim alist as New ArrayList
'**
'**  alist.add "aaa"
'**  alist.add "bbb"
'**  ...
'**
'**  Dim val
'**  For Each val In alist.getArray
'**    Debug.Print val
'**  Next
'**
'**
'**
'**
'** version 0.0.0                          Created by ntennya(GitHub user name)
'*******************************************************************************
Option Explicit

Private Const EXPANDFREQ As Integer = 10

Dim Marr() As Variant

Dim Mcnt As Long
Dim MmaxSize As Long
Dim Mfreq As Long
'*******************************************************************************
'*** Constructor ***************************************************************
'*******************************************************************************
Private Sub Class_Initialize()
  Mcnt = 0
  MmaxSize = 0
  Mfreq = EXPANDFREQ
End Sub


'*******************************************************************************
'*** Public Methods ************************************************************
'*******************************************************************************
'=== Setter ================================================
Public Function add(value As Variant)
  If Mcnt >= MmaxSize Then
    MmaxSize = MmaxSize + Mfreq
    ReDim Preserve Marr(MmaxSize)
  End If
  Marr(Mcnt) = value
  Mcnt = Mcnt + 1
End Function


Public Function addObject(obj As Object)
  If Mcnt >= MmaxSize Then
    MmaxSize = MmaxSize + Mfreq
    ReDim Preserve Marr(MmaxSize)
  End If
  Set Marr(Mcnt) = obj
  Mcnt = Mcnt + 1
End Function


Public Function appendValueList(another As ArrayList)
  If another.count = 0 Then Exit Function
  
  Dim val
  For Each val In another.getList
    add val
  Next
End Function


Public Function appendObjectList(another As ArrayList)
  If another.count = 0 Then Exit Function
  
  Dim val, obj As Object
  For Each obj In another.getList
    Set obj = val  'Type Cast
    addObject obj
  Next
End Function


'=== Getter ================================================
Public Function count()
  count = Mcnt
End Function


Public Function getValue(index As Long) As Variant
  getValue = Marr(index)
End Function


Public Function getObject(index As Long) As Object
  Set getObject = Marr(index)
End Function


Public Function getList()
  Dim retarr() As Variant
  retarr = Marr       'Deep copy
  ReDim Preserve retarr(Mcnt - 1)
  getList = retarr
End Function


Public Function contains(value As Variant) As Boolean
  contains = False
  
  Dim val
  For Each val In Marr
    If val = value Then
      contains = True
      Exit Function
    End If
  Next
End Function



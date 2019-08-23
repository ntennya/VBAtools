VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tuple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'**  Tuple                                                         *************
'*******************************************************************************
'** Target user : All programmers.
'** -- Explanation -------------------------------------------------------------
'** Tuple enable to return two value.
'**
'**  Attention...
'**  (1) The differance between "value(Variant)" and "object variable"
'**   In VBA, the pointer that is called "Object Variable" needs special treatment.
'**   For example, if assign a object variable to an object, it needs "Set"
'**   statement.
'** -- Methods -----------------------------------------------------------------
'**
'** setValues
'** setObjects
'** setValueAndObject
'** setObjectAndValue : Set two variables to the tuple. Please use these properly
'**                     depending on the variable type(object or Variant).
'**
'** getValue1,getObject1 : Get the 1st(order in set...) valiant/object.
'** getValue2,getObject2 : Get the 2nd(order in set...) valiant/object.
'**
'** -- Example of use ----------------------------------------------------------
'**
'** Public Function ageAndBuyList() As Touple
'**   Dim ret As New Tuple                     '<<Important!!!>>  To return an
'**                                            'object, it needs to "New"
'**                                            'IN this function.
'**   Dim age As Integer
'**   Dim buylist As New ArrayList
'**   ...
'**   Call ret.setValueAndObject(age, buylist) 'Set valiant and Object
'**   ageAndBuyList = ret                      'Return Tuple.
'** End Function
'**
'**
'** version 0.0.0                          Created by ntennya(GitHub user name)
'*******************************************************************************
Option Explicit

Private Mval1 As Variant, Mval2 As Variant
'*******************************************************************************
'*** Public Methods ************************************************************
'*******************************************************************************
'=== Setter ================================================
Public Function setValues(val1 As Variant, val2 As Variant)
  Mval1 = val1
  Mval2 = val2
End Function


Public Function setObjects(val1 As Object, val2 As Object)
  Set Mval1 = val1
  Set Mval2 = val2
End Function


Public Function setValueAndObject(val1 As Variant, val2 As Object)
  Mval1 = val1
  Set Mval2 = val2
End Function


Public Function setObjectAndValue(val1 As Object, val2 As Variant)
  Set Mval1 = val1
  Mval2 = val2
End Function


'=== Getter ================================================
Public Function getValue1() As Variant
  getValue1 = Mval1
End Function


Public Function getValue2() As Variant
  getValue2 = Mval2
End Function


Public Function getObject1() As Variant
  Set getObject1 = Mval1
End Function


Public Function getObject2() As Variant
  Set getObject2 = Mval2
End Function

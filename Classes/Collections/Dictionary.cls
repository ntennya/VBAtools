VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'** Dictionary                                                     *************
'*******************************************************************************
'** Target user : Intermediate to Expert programmers.
'** -- Need Modules/Classes ----------------------------------------------------
'**
'**
'** -- Explanation -------------------------------------------------------------
'** This is Dictionary(in VB.Net/VBA/C#) / HashMap(in Java and so on)
'**
'** The key should be not an object. But object is able to use for value.
'** (Please refer below attention(1).)
'**
'**  Attention...
'**  (1) The differance between "value(Variant)" and "object variable".
'**   In VBA, the pointer that is called "Object Variable" needs special treatment.
'**   For example, if assign a object variable to an object, it needs "Set"
'**   statement.
'**
'**    (NG) Dim pointer : pointer = new Tuple
'**    (OK) Dim pointer : Set pointer = new Tuple
'**
'**   So, the method will be used properly. For example,
'**    "setValue", "getValue" methods are for value elements.
'**    "setObject", "getObject" methods are for object elements.
'**
'**
'** -- Methods -----------------------------------------------------------------
'**
'**  setValue/setObject : Add a key/value pair to Dictionary.
'**  getValue/getObject : Get a value/object using the key.
'**  containsKey        : Will return True if the key exists.
'**
'** -- Example of use ----------------------------------------------------------
'**
'**
'**
'**
'**
'**
'**
'**
'**
'**
'**
'** version 0.0.0                          Created by ntennya(GitHub user name)
'*******************************************************************************
Option Explicit

Private Const ERRORSTR = "<<@@# ERRORSTR #@@>>>"

Dim Mlist As ArrayList
'*******************************************************************************
'*** Constructor ***************************************************************
'*******************************************************************************
Private Sub Class_Initialize()
  Set Mlist = New ArrayList
End Sub


'*******************************************************************************
'*** Private Methods ***********************************************************
'*******************************************************************************
Private Function changeValue(key As Variant, value As Variant) As Boolean
  changeValue = False
  If Not containsKey(key) Then Exit Function
    
  Dim kv As Tuple
  Set kv = getObject(key)
  kv.setValues key, value
  
  changeValue = True
End Function


Private Function changeObject(key As Variant, value As Object) As Boolean
  changeObject = False
  If Not containsKey(key) Then Exit Function
    
  Dim kv As Tuple
  Set kv = getObject(key)
  kv.setValueAndObject key, value
  
  changeObject = True
End Function
'*******************************************************************************
'*** Public Methods ************************************************************
'*******************************************************************************
'=== Setter ================================================
Public Function setValue(key As Variant, value As Variant)
  If containsKey(key) Then
    changeValue key, value
    Exit Function
  End If

  Dim tpl As New Tuple
  tpl.setValues key, value
  Mlist.addObject tpl
End Function


Public Function setObject(key As Variant, value As Object)
  If containsKey(key) Then
    changeObject key, value
    Exit Function
  End If
  Dim tpl As New Tuple
  tpl.setValueAndObject key, value
  Mlist.addObject tpl
End Function


'=== Getter ================================================
Public Function containsKey(key As Variant) As Boolean
  containsKey = False
  If Mlist.count = 0 Then Exit Function
  
  Dim val, tpl As Tuple
  For Each val In Mlist.getList
    Set tpl = val   'type cast
    If tpl.getValue1 = key Then
      containsKey = True
      Exit Function
    End If
  Next
End Function


Public Function getPair(key As Variant) As Tuple
  Set getPair = Nothing   'No work but need for readablility.
  If Not containsKey(key) Then Exit Function

  Dim val, tpl As Tuple
  For Each val In Mlist.getList
    Set tpl = val   'type cast
    If tpl.getValue1 = key Then
      Set getPair = tpl
      Exit Function
    End If
  Next
End Function


Public Function getValue(key As Variant) As Variant
  getValue = ERRORSTR
  If Not containsKey(key) Then Exit Function
  
  getValue = getPair(key).getValue2
End Function


Public Function getObject(key As Variant) As Object
  Set getObject = Nothing
  If Not containsKey(key) Then Exit Function
  
  Set getObject = getPair(key).getObject2
End Function


Public Function getDictionary() As ArrayList
  getDictionary = Mlist.getList    'deepcopy
End Function





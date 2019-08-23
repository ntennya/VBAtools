Attribute VB_Name = "ClipboardFuncs"
'*******************************************************************************
'** Clipboard Funcs                                                *************
'*******************************************************************************
'**
'** -- Explanation -------------------------------------------------------------
'**
'** These function enable clipboard treatment easy.
'**
'** version 0.0.0                          Created by ntennya(GitHub user name)
'*******************************************************************************
Option Explicit

'*******************************************************************************
'** getClipboardText                                               *************
'*******************************************************************************
'** -- Explanation -------------------------------------------------------------
'** Get text from clipboard.
'*******************************************************************************
Public Function getTextFromClipboard(Optional showMsg As Boolean = True)
getTextFromClipboard = ""
On Error GoTo ERREND

  With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  'DataObject
    .GetFromClipboard
    getTextFromClipboard = .getText
  End With
  Exit Function
  
ERREND:
  If showMsg Then Call MsgBox("[getTextFromClipboard] failed." + vbNewLine + "May be current clipboard target is not text(String).", vbOKOnly + vbExclamation, "Clipboard function Failed")
End Function

'*******************************************************************************
'** setTextToClipboard                                             *************
'*******************************************************************************
'** -- Explanation -------------------------------------------------------------
'** Set text to clipboard.
'*******************************************************************************
Public Function setTextToClipboard(text As String, Optional showMsg As Boolean = True)
On Error GoTo ERREND

  With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  'DataObject
      .setText text
      .PutInClipboard
  End With
  Exit Function
  
ERREND:
  If showMsg Then Call MsgBox("[setTextToClipboard] failed.", vbOKOnly + vbExclamation, "Clipboard function Failed")
End Function

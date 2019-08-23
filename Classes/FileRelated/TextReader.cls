VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'** textReader                                                     *************
'*******************************************************************************
'** Target user : All programmers.
'** -- Need Modules/Classes ----------------------------------------------------
'**
'**
'** -- Explanation -------------------------------------------------------------
'** This will make text file read more easy and rapid.
'**
'**    Available format: Unicode/ASCII/UTF-8/UTF-8N
'**
'** -- Methods -----------------------------------------------------------------
'**
'**  readText : Read a text file.
'**    fileFP  : Target file fullpath.
'**    format  : 0:ASCII / -1:Unicode / -2:systemDefault
'**    showMsg : Show a message if an error occurd in this method.
'**
'**  readUTF8 : Read a UTF-8 text file.
'**    fileFP  : Target file fullpath.
'**    showMsg : Show a message if an error occurd in this method.
'**
'**  existsFile : Will return "True" if the file exists.
'**
'** -- Example of use ----------------------------------------------------------
'**
'**  Dim text as String
'**  Dim tr as New TextReader
'**  text = tr.readText("C:\Debug\test.txt")
'**
'**
'**  (If you want to get the text line by line, please try like below.)
'**
'**  Dim line
'**  For Each line in Split(text,vbNewLine)
'**    Debug.Print line
'**  Next
'**
'**
'** version 0.0.0                          Created by ntennya(GitHub user name)
'*******************************************************************************
Option Explicit

Private MfileFP As String
'*******************************************************************************
'*** Public Methods ************************************************************
'*******************************************************************************
'=== Getter ================================================
Public Function existsFile() As Boolean
    existsFile = (Dir(MfileFP) <> "")
End Function


Public Function readText(Optional fileFP As String = "", Optional format As Integer = -1, Optional showMsg As Boolean = True) As String
On Error GoTo ERREND

  If fileFP <> "" Then MfileFP = fileFP
  If Not existsFile Then
    MfileFP = ""
    GoTo ERREND
  End If
  
  With CreateObject("Scripting.FileSystemObject").OpenTextFile(fileFP, 1, True, format)  'format: 0:ASCII / -1:Unicode / -2:systemDefault
    readText = .ReadAll
  End With

  Exit Function
ERREND:
 If showMsg Then Call MsgBox("[read] failed.", vbOKOnly + vbExclamation, "TextFileAccessor method Failed")
End Function


Public Function readUTF8(Optional fileFP As String = "", Optional showMsg As Boolean = True) As String
On Error GoTo ERREND

  If fileFP <> "" Then MfileFP = fileFP
  If Not existsFile Then
    MfileFP = ""
    GoTo ERREND
  End If

  With CreateObject("ADODB.Stream")
    .Charset = "UTF-8"
    .Open
    .LoadFromFile fileFP
    readUTF8 = .readText
    .Close
  End With

  Exit Function
ERREND:
 If showMsg Then Call MsgBox("[readUtf8] failed.", vbOKOnly + vbExclamation, "TextFileAccessor method Failed")
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'** TextWriter                                                     *************
'*******************************************************************************
'** Target user : All programmers.
'** -- Need Modules/Classes ----------------------------------------------------
'** *TextReader
'**
'** -- Explanation -------------------------------------------------------------
'** This will make text file creation more easy and rapid.
'**
'**    Available format: Unicode/ASCII/UTF-8/UTF-8N
'**
'** [Attention!!!]
'** This object will OVERWRITE the existing same-name file WITHOUT CHECK.
'** Please pay attention only this.
'**
'** -- Methods -----------------------------------------------------------------
'**
'** writeText : Write text to file.
'**        text    : text(String) for write.
'**        fileFP  : The fullpath of output file. (Ex: "C:\debug\test.txt")
'**        showMsg : Show a message if an error occurd in this method.
'**        unicode : True:Unicode(default)/False:ASCII
'**
'** writeUTF8 : Write text to file with UTF-8 format.
'**        text      : text(String) for write.
'**        fileFP    : The fullpath of output file. (Ex: "C:\debug\test.txt")
'**        removeBom : Remove the "BOM".(Save as UTF-8N)
'**        showMsg   : Show a message if an error occurd in this method.
'**
'** appendText/appendUTF8 : Append text to original file.
'**   Arguments are same as writeText/writeUTF8. In adition to...
'**     readFormat : The format of read file.(See TextReader explanation.)
'**     joinStr    : The string that joins original text and the text to add.
'**                  ( default: vbNewLine (CR+LF) )
'** -- Example of use ----------------------------------------------------------
'**
'**  With New TextWriter
'**    .writeText "aaa"+VbNewLine+"bbb", "C:\debug\test.txt"
'**    .appendText "ccc"    (If fileFP is omitted, previous fileFP is applied.)
'**  End With
'**
'**  ... Will created a file...  aaa
'**                              bbb
'**                              ccc
'**
'**
'** version 0.0.0                          Created by ntennya(GitHub user name)
'*******************************************************************************
Option Explicit

Private MfileFP As String
'*******************************************************************************
'*** Public Methods ************************************************************
'*******************************************************************************
'=== Setter ================================================
Public Function writeText(ByVal text As String, Optional fileFP As String = "", Optional showMsg As Boolean = True, Optional unicode As Boolean = True)
  If text = "" Then text = ""  ' The argument "text" may Null if the argument is text property of component.
On Error GoTo ERREND
  If fileFP <> "" Then MfileFP = fileFP

  With CreateObject("Scripting.FileSystemObject").CreateTextFile(fileFP, True, unicode)
    .Write text
    .Close
  End With

  Exit Function
ERREND:
 If showMsg Then Call MsgBox("[writeText] failed.", vbOKOnly + vbExclamation, "TextFileAccessor method Failed")
End Function


Public Function writeUTF8(ByVal text As String, Optional fileFP As String = "", Optional removeBom As Boolean = True, Optional showMsg As Boolean = True)
    If text = "" Then text = ""  ' The argument "text" may Null if the argument is text property of component.
With CreateObject("ADODB.Stream")
On Error GoTo ERREND

  If text = "" And removeBom And showMsg Then
    Call MsgBox("Unable to Save empty string to [UTF-8/Non Bom File.]", vbExclamation + vbOKOnly, "Unable to Save empty string")
  End If

  If fileFP <> "" Then MfileFP = fileFP
  
  .Type = 2
  .Charset = "UTF-8"
  .Open
  .writeText text
  
  .Position = 0
  .Type = 1   'adTypeBinary

  If removeBom Then .Position = 3
  Dim tmpTxt() As Byte
  tmpTxt = .read
  .Close
  
  .Open
  .Write tmpTxt
  .SaveToFile fileFP, 2    '2 means "overrite"  if 1: error
  .Close

  Exit Function
End With
ERREND:
 If showMsg Then Call MsgBox("[writeUTF8] failed.", vbOKOnly + vbExclamation, "TextFileAccessor method Failed")
End Function


Public Function appendText(ByVal text As String, Optional fileFP As String = "", Optional showMsg As Boolean = True, Optional unicode As Boolean = True, Optional readFormat As Integer = -1, Optional joinStr As String = vbNewLine)
  Dim orgtext As String: orgtext = ""
On Error GoTo ERREND
  If fileFP <> "" Then MfileFP = fileFP
  
With New TextReader
  If Dir(MfileFP) <> "" Then
    orgtext = .readText(MfileFP, readFormat, showMsg)
  End If
End With

  Call writeText(orgtext + joinStr + text, MfileFP, showMsg, unicode)

  Exit Function
ERREND:
 If showMsg Then Call MsgBox("[appendText] failed.", vbOKOnly + vbExclamation, "TextFileAccessor method Failed")
End Function


Public Function appendUTF8(ByVal text As String, Optional fileFP As String = "", Optional removeBom As Boolean = True, Optional showMsg As Boolean = True, Optional joinStr As String = vbNewLine)
  Dim orgtext As String: orgtext = ""
On Error GoTo ERREND
  If fileFP <> "" Then MfileFP = fileFP
  
With New TextReader
  If Dir(MfileFP) <> "" Then
    orgtext = .readUTF8(MfileFP, showMsg)
  End If
End With

  Call writeUTF8(orgtext + joinStr + text, fileFP, removeBom, showMsg)

  Exit Function
ERREND:
 If showMsg Then Call MsgBox("[appendUTf8] failed.", vbOKOnly + vbExclamation, "TextFileAccessor method Failed")
End Function

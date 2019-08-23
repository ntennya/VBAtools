VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSearcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'** FileSearcher                                                   *************
'*******************************************************************************
'** Target user : All programmers.
'** -- Need Modules/Classes ----------------------------------------------------
'** *ArrayList
'**
'** -- Explanation -------------------------------------------------------------
'** The object of this class scans all child files/folders in a folder
'** by recursive call.
'**
'** -- Methods -----------------------------------------------------------------
'**
'**  getFiles : Get all child/grandchild/... files full path.
'**      folderFP : The root folder for scan.
'**
'**  getFolders : Get all child/grandchild/... folders full path.
'**      folderFP : The root folder for scan.
'**
'**  execute : Execute any procedure for each file when getFiles is called.
'**            Pleasse customize this method as you like.
'**      fileFP : The file full path for use.
'**
'** -- Example of use ----------------------------------------------------------
'**
'** Dim fs as New FileSearcher
'**
'** Dim fp
'** For Each fp In fs.getFiles
'**   Debug.Print fp
'** Next
'**
'**
'**
'** version 0.0.0                          Created by ntennya(GitHub user name)
'*******************************************************************************
Option Explicit

'*******************************************************************************
'*** Customizable Methods ******************************************************
'***     Please customize this as you like. ************************************
'*******************************************************************************
Private Function execute(fileFP As String)
  'Debug.Print fileFP
End Function


'*******************************************************************************
'*** Public Methods ************************************************************
'*******************************************************************************
'=== Setter ================================================
'=== Getter ================================================
Public Function getFiles(folderFP As String) As ArrayList
With CreateObject("Scripting.FileSystemObject")

  Dim retlist As New ArrayList
  
  Dim file
  For Each file In .GetFolder(folderFP).files
    retlist.add file.Path
    Call execute(file.Path)
  Next
  
  Dim fol, fch As FileSearcher, files As ArrayList
  For Each fol In .GetFolder(folderFP).SubFolders
    Set fch = New FileSearcher
    Set files = fch.getFiles(fol.Path)
    Call retlist.appendValueList(files)
  Next
  
  Set getFiles = retlist
End With
End Function


Public Function getFolders(folderFP As String, Optional includeMe = True) As ArrayList
With CreateObject("Scripting.FileSystemObject")

  Dim retlist As New ArrayList
  If includeMe Then retlist.add folderFP
  
  Dim fol
  For Each fol In .GetFolder(folderFP).SubFolders
    retlist.add fol.Path
  Next
  
  Dim fch As FileSearcher, files As ArrayList
  For Each fol In .GetFolder(folderFP).SubFolders
    Set fch = New FileSearcher
    Set files = fch.getFolders(fol.Path, False)
    Call retlist.appendValueList(files)
  Next
  
  Set getFolders = retlist
End With
End Function


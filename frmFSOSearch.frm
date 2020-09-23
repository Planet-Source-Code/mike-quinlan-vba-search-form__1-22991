VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFSOSearch 
   Caption         =   "Search (FSO)"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   OleObjectBlob   =   "frmFSOSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFSOSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
      Option Explicit
      Dim fso As New FileSystemObject
      Dim fld As Folder
      Dim lSize As Long
'"Assembled" by Mike Quinlan vb@qdog.com
'**Must add reference for Microsoft Scripting Runtime
'This is a VBA module (tested in word 2000) of a complete search function
'it is a work in progress but currently you give it the path and the file specs
'including wild cards.  It will search and populate a list pox with the results
'and also tell you the number of directories searched, files found and the total size.
Private Sub cmdSearch_Click()
    Dim nDirs As Integer
    Dim nFiles As Integer
    Dim sDir As String
    Dim sSrchString As String
    If cmdSearch.Caption <> "Cancel" Then
        List1.Clear
    End If
    
    If cmdSearch.Caption = "Cancel" Then
        cmdSearch.Caption = "Canceled"
        Exit Sub
    End If
    
    
    sDir = InputBox("Please enter the directory to search", _
    "Directory", "C:\")
    If sDir = "" Then
        MsgBox "Search Aborted", vbInformation, "User Abort"
        Exit Sub
    End If
    
    sSrchString = InputBox("Please enter the file name to search", _
    "Search String", "*.doc")
    If sSrchString = "" Then
        MsgBox "Search Aborted", vbInformation, "User Abort"
        Exit Sub
    End If
    MousePointer = 11
    cmdSearch.Caption = "Cancel"
    lblPath.Caption = "Searching " & vbCrLf & UCase(sDir) & "..."
    lSize = FindFile(sDir, sSrchString, nDirs, nFiles)
    MousePointer = 0
    If cmdSearch.Caption = "Canceled" Then
        lblPath.Caption = "Search Canceled"
        cmdSearch.Caption = "Search"
    Else
    lblPath.Caption = "Search Complete"
    End If
    'MsgBox Str(nFiles) & " files found in" & Str(nDirs) & _
    '" directories", vbInformation
    'MsgBox "Total Size = " & Format(lSize, "#,###,###,##0") & " bytes"
    lblSize.Caption = Format(lSize, "#,###,###,##0") & " Bytes"
End Sub

Private Function FindFile(ByVal sFol As String, sFile As String, _
    nDirs As Integer, nFiles As Integer) As Long
    Dim MyFolder As Folder
    Dim MyFile As File
    Dim FileName As String
    Dim f
    If cmdSearch.Caption = "Canceled" Then
        Exit Function
    End If
    
    Dim shit As String
    Set fld = fso.GetFolder(sFol)
    
    FileName = Dir(fso.BuildPath(fld.path, sFile), vbNormal Or _
    vbHidden Or vbSystem Or vbReadOnly)
    
    

    While Len(FileName) <> 0
        Set f = fso.GetFile(fso.BuildPath(fld.path, FileName))
       
      'check for files that are older than value in label
       shit = f.DateLastModified
        If f.DateLastModified < Now() - txtNumberofDays.Value Then
            FindFile = FindFile + FileLen(fso.BuildPath(fld.path, FileName))
            nFiles = nFiles + 1
            List1.AddItem fso.BuildPath(fld.path, FileName) & "     " & shit     ' Load ListBox
        End If
        FileName = Dir()  ' Get next file
        DoEvents
    Wend
    lblPath = "Searching " & vbCrLf & fld.path & "..."
    nDirs = nDirs + 1
    lbldir.Caption = nDirs
    If fld.SubFolders.Count > 0 Then
        For Each MyFolder In fld.SubFolders
        'below are directories you want to skip
            If MyFolder <> "C:\Recycled" Then
            If MyFolder <> "G:\CAMS\Installation Software" Then
            If MyFolder <> "G:\CAMS\Pseudo Holding" Then
            If MyFolder <> "G:\CAMS\LGK" Then
                DoEvents
                FindFile = FindFile + FindFile(MyFolder.path, sFile, nDirs, nFiles)
                
            End If
            End If
            End If
            End If
        Next
    End If
    lblFiles.Caption = nFiles
    
    
End Function


Attribute VB_Name = "RecursivePageCounting"
Option Explicit

Private Sub ProgressForPageCount(currentFileName As String, ByVal filesCounter As Object)
    ProgressForPageCounting.currentFileName.Caption = "Текущий файл: " & currentFileName
    ProgressForPageCounting.folderCount.Caption = filesCounter.Item("folders")
    ProgressForPageCounting.docCount.Caption = filesCounter.Item("docs")
    ProgressForPageCounting.pageCount.Caption = filesCounter.Item("pages")
    ProgressForPageCounting.Repaint
    DoEvents
End Sub

Function CollectionExplorer(ByVal path As String, _
                            ByRef report As String, _
                            ByVal filesCounter As Object, _
                            ByVal indent As String)
                        
    Dim key As Variant
    Dim tmpFilesCounter
    Set tmpFilesCounter = CreateObject("Scripting.Dictionary")
    For Each key In FolderExplorer(path)
        Dim normalFileName As String
        normalFileName = Mid(key, InStrRev(key, "\") + 1)
        Call ProgressForPageCount(normalFileName, filesCounter)
        If GetAttr(key) = 16 Or GetAttr(key) = 48 Then
            On Error Resume Next
            filesCounter.Item("folders") = filesCounter.Item("folders") + 1
            report = report & indent & "F: """ & UCase(normalFileName) & """ begin" & Chr(13)
            Call CollectionExplorer(key, report, filesCounter, indent & "|" & Chr(9))
            report = report & indent & "|______  end" & Chr(13)
        Else
            Dim ext As String, noExt As String
            ext = UCase(Mid(key, InStrRev(key, ".") + 1))
            'checking file has ext
            noExt = Mid(ext, InStrRev(ext, "\") + 1)
            If noExt <> ext Then
                If tmpFilesCounter.Exists("NoExt") Then
                    tmpFilesCounter.Item("NoExt") = tmpFilesCounter.Item("NoExt") + 1
                Else
                    tmpFilesCounter.Add "NoExt", 1
                End If
                report = report & indent & "N: " & normalFileName & Chr(13)
                GoTo NextIter
            End If
            If tmpFilesCounter.Exists(ext) Then
                tmpFilesCounter.Item(ext) = tmpFilesCounter.Item(ext) + 1
            Else
                tmpFilesCounter.Add ext, 1
            End If
            If ext = "DOC" Or ext = "DOCX" Or ext = "RTF" Then
                Dim tmp As Long
                Word.Application.Documents.Open key, Visible:=False, ReadOnly:=True
                tmp = Documents.Item(key).ComputeStatistics(2)
                filesCounter.Item("pages") = filesCounter.Item("pages") + tmp
                filesCounter.Item("docs") = filesCounter.Item("docs") + 1
                Word.Application.DisplayAlerts = False
                Word.Application.Documents.Item(key).Close (wdDoNotSaveChanges)
                Word.Application.DisplayAlerts = False
                report = report & indent & "D: " & UCase(normalFileName) & " (PAGE COUNT: " & tmp & ")" & Chr(13)
            Else
                report = report & indent & "N: " & normalFileName & Chr(13)
            End If
        End If
NextIter:
    Next
    If tmpFilesCounter.count > 0 Then
        report = report & indent & "*****************" & Chr(13)
        For Each key In tmpFilesCounter
            If key <> "DOC" And key <> "DOCX" And key <> "RTF" Then
                If filesCounter.Exists(key) Then
                    filesCounter.Item(key) = filesCounter.Item(key) + tmpFilesCounter.Item(key)
                Else
                    filesCounter.Add key, tmpFilesCounter.Item(key)
                End If
            End If
            report = report & indent & "*." & key & Chr(9) & tmpFilesCounter.Item(key) & Chr(13)
        Next
    End If
    
End Function

Function FolderExplorer(ByVal path As String) As Collection
    Set FolderExplorer = New Collection
    Dim strPath As String
    strPath = Dir(path & "\", vbDirectory)
    Do While strPath <> ""
        If strPath <> "." And strPath <> ".." Then
            FolderExplorer.Add (path & "\" & strPath)
        End If
        strPath = Dir()
    Loop
End Function

Sub newDoc(text As String)
    Dim newDoc As Document
    Set newDoc = Documents.Add
    With newDoc
        .Content.Font.name = "Calibri"
        .Content.Font.Size = 9
        .PageSetup.Orientation = wdOrientLandscape
    End With
    Word.Selection = text
    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
    End With
    With Selection.PageSetup
      .LeftMargin = CentimetersToPoints(0.5)
      .TopMargin = CentimetersToPoints(0.5)
      .BottomMargin = CentimetersToPoints(0.5)
      .RightMargin = CentimetersToPoints(0.5)
    End With
    Documents(newDoc).Activate
End Sub

Sub RecursivePageCounting()
    Dim report As String
    Dim generalInfo As String
    Dim indent As String: indent = "|"
    Dim path As String
    
    Dim filesCounter
    Set filesCounter = CreateObject("Scripting.Dictionary")
    filesCounter.Add "pages", 0
    filesCounter.Add "docs", 0
    filesCounter.Add "folders", 1
    
        
    With Word.Application.FileDialog(msoFileDialogFolderPicker)
           .Title = "Select a folder"
           .ButtonName = "Ok"
           .Filters.Clear
           .InitialFileName = "C:\Users\User\Documents"
           .InitialView = msoFileDialogViewList
           If .Show = 0 Then Exit Sub
           path = .SelectedItems(1)
    End With
        
    ProgressForPageCounting.Caption = path
    Word.Application.ScreenUpdating = False
    ProgressForPageCounting.Show
    
    generalInfo = Chr(9) & "Report for " & UCase(path) & Chr(13)  ' Chr(13) - line break (\n), Chr(9) - tab (\t)

    report = Chr(9) & "DETAILED REPORT" & Chr(13) & Chr(13) & "Explanation for file tree: " & Chr(13) & Chr(13) & _
    "F - folder" & Chr(13) & _
    "D - MS Word document" & Chr(13) & _
    "N - not MS Word document" & Chr(13) & Chr(13) & _
    "Root folder """ & UCase(Mid(path, InStrRev(path, "\") + 1)) & """ begin" & Chr(13)
    
    Call CollectionExplorer(path, report, filesCounter, indent & Chr(9))

    generalInfo = generalInfo & Chr(13) & "Pages count:" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & filesCounter.Item("pages") & Chr(13) & _
    "MS Word documents count:" & Chr(9) & Chr(9) & Chr(9) & filesCounter.Item("docs") & Chr(13) & _
    "Folders count (including the root folder):" & Chr(9) & filesCounter.Item("folders") & Chr(13) & Chr(13)
    If filesCounter.count > 3 Then
        Dim key As Variant
        generalInfo = generalInfo & "Other  files:" & Chr(13)
        For Each key In filesCounter
            If key <> "docs" And key <> "pages" And key <> "folders" Then
                generalInfo = generalInfo & "*." & key & Chr(9) & filesCounter.Item(key) & Chr(13)
            End If
        Next
        generalInfo = generalInfo & Chr(13)
    End If
    generalInfo = generalInfo & report & "|______  end" & Chr(13) & "END OF REPORT"
    
    ProgressForPageCounting.Hide
    Word.Application.ScreenUpdating = True
    
    Call newDoc(generalInfo)
End Sub





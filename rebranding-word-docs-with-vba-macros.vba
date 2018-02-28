Option Explicit
'I created this collection of subs to automate the rebranding of 1000+ word documents,
'inserting new headers/footers with embedded logos, bookmarks, and custom styles.
'It could be pointed at a folder of hundreds of documents and process them all in minutes.
'
'Additional reprocessing functionality was added when the new branding format was changed 
'mid-project, after some docs had already been updated with the now-suddenly-out-of-date branding. 
'
Public Const GENERAL_HEADER = "C:\files\general-header-with-image-object.docx"
Public Const GENERAL_FOOTER = "C:\files\general-footer-with-image-object.docx"
Public Const INTAKE_HEADER = "C:\files\intake-header-with-image-object.docx"
Public Const INTAKE_FOOTER = "C:\files\intake-footer-with-image-object.docx"
Public Const STYLES_DOTM = "C:\files\styles.dotm"
Public Const BLDG_BLOCKS = "C:\Users\#####\AppData\Roaming\Microsoft\Document Building Blocks\1033\15\Built-In Building Blocks.dotx"
'more properly (Environ("APPDATA") & "\Microsoft\Document Building Blocks\1033\15\Built-In Building Blocks.dotx", but can't use a variable in a Const obvi
Public pStrDocType As String
Public pInitialPages As Integer

Sub CheckConsts()
    MsgBox "GENERAL_HEADER = " & GENERAL_HEADER & vbCrLf & _
           "GENERAL_FOOTER = " & GENERAL_FOOTER & vbCrLf & _
           "INTAKE_HEADER = " & INTAKE_HEADER & vbCrLf & _
           "INTAKE_FOOTER = " & INTAKE_FOOTER & vbCrLf & _
           "STYLES_DOTM = " & STYLES_DOTM & vbCrLf & _
           "BLDG_BLOCKS = " & BLDG_BLOCKS
End Sub

Sub UpdateDocuments()
    Application.ScreenUpdating = False
    Dim strFolder As String, strFile As String, strFullName As String, wdDoc As Document
    
    'Which kind of processing is this?
    pStrDocType = InputBox("Are we processing 'general' or 'intake'?")
    If pStrDocType <> "general" And pStrDocType <> "intake" Then
        MsgBox "ERROR: must enter 'general' or 'intake' for doc type"
        Exit Sub
    End If
    
    'Where are the documents?
    strFolder = InputBox("Path to folder of target documents:")
    If strFolder = "" Then Exit Sub
    
    strFile = Dir(strFolder & "\*.doc*", vbNormal)
    While strFile <> ""
        strFullName = strFolder & "\" & strFile
        Set wdDoc = Documents.Open(filename:=strFullName, AddToRecentFiles:=False, Visible:=False)
        With wdDoc
            'check for read-only. If detected: save-as a .tmp w/o read-only, delete the original, and rename the copy
            If ActiveDocument.ReadOnly = True Then
                Dim strNewname As String
                strNewname = strFullName & ".tmp"
                .SaveAs2 filename:=strNewname, ReadOnlyRecommended:=False
                .Close SaveChanges:=False
                'delete the original
                Kill strFullName
                'rename the copy to the original's name
                Name strNewname As strFullName
                'reopen
                Documents.Open filename:=strFullName, AddToRecentFiles:=False, Visible:=False
            End If
            
            If pStrDocType = "general" Then
                  GeneralRebrand
            ElseIf pStrDocType = "intake" Then
                  IntakeRebrand
            End If
            ActiveDocument.Close SaveChanges:=True
        End With
      strFile = Dir()
    Wend
    Set wdDoc = Nothing
    Application.ScreenUpdating = True
End Sub

Sub GeneralRebrand()
    pStrDocType = "general"
    Call RebrandDocuments(GENERAL_HEADER, GENERAL_FOOTER)
End Sub

Sub IntakeRebrand()
    pStrDocType = "intake"
    Call RebrandDocuments(INTAKE_HEADER, INTAKE_FOOTER)
End Sub

Sub RebrandDocuments(headFile As String, footFile As String)
'
' RebrandDocuments Macro for rebranding documents
' NOTE: if called on its own, changes to the document are NOT SAVED,
' so it's good for testing just GENERAL or just INTAKE subs or processing one-offs
'

    'how many pages are there before we begin?
    pInitialPages = ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
    
    '''
    'delete old things
    deleteAboveTimeDateField
    deleteExistingHeaderFooter
    deleteLingeringBookmarks
    
    'activate the first page header/footer toggle
    ActiveDocument.Sections(1).PageSetup.DifferentFirstPageHeaderFooter = True
    'goto first page
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'HEADER:
    'insert desired text and graphic
    ActiveWindow.ActivePane.View.SeekView = wdSeekFirstPageHeader
    Selection.InsertFile (headFile)
    
    'delete extra CR added to header by insertion method
    Selection.EndKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, Count:=1
    
    'return to main window
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'FOOTER
    'insert desired text and graphic
    ActiveWindow.ActivePane.View.SeekView = wdSeekFirstPageFooter
    Selection.InsertFile (footFile)
    
    'delete extra CR added to footer by insertion method
    Selection.EndKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, Count:=1
    
    'return to main window
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'APPLY NEW STYLES
    With ActiveDocument
        .UpdateStylesOnOpen = True
        .AttachedTemplate = STYLES_DOTM
    End With
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'add page #s to all pages past page 1
    'must occur after applying styles as it calls on an embedded custom style
    addPagesAndOrPageNumbers
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Sub deleteAboveTimeDateField()
    'if there's fields...
    If ActiveDocument.Fields.Count >= 1 Then
        Dim fieldLoop As Field
        'loop through until you encounter the first TIME or DATE field
        'then go to just before that field, select everthing above it, and delete
        For Each fieldLoop In ActiveDocument.Fields
            If InStr(1, fieldLoop.Code.Text, "TIME", 1) Then
                Selection.GoTo What:=wdGoToField, Name:="Time", Which:=wdGoToFirst
                Selection.StartOf Unit:=wdStory, Extend:=wdExtend
                'Selection must be >1 character, or else it's just the Time/Date field itself
                If Selection.Characters.Count > 1 Then
                    Selection.Delete
                End If
                Exit Sub
            ElseIf InStr(1, fieldLoop.Code.Text, "DATE", 1) Then
                Selection.GoTo What:=wdGoToField, Name:="Date", Which:=wdGoToFirst
                Selection.StartOf Unit:=wdStory, Extend:=wdExtend
                'Selection must be >1 character, or else it's just the Time/Date field itself
                If Selection.Characters.Count > 1 Then
                    Selection.Delete
                End If
                Exit Sub
            End If
        Next fieldLoop
        'We shouldn't reach this point if we saw a time/date field, so...
        MsgBox "ERROR: fields detected, but none were TIME or DATE."
    Else
        MsgBox "ERROR: no fields detected."
    End If
End Sub

Sub deleteExistingHeaderFooter()
'
' deleteExistingHeaderFooter Macro
'
'
    Dim oSec As Section
    Dim oHead As HeaderFooter
    Dim oFoot As HeaderFooter

    For Each oSec In ActiveDocument.Sections
        For Each oHead In oSec.Headers
            If oHead.Exists Then oHead.Range.Delete
        Next oHead

        For Each oFoot In oSec.Footers
            If oFoot.Exists Then oFoot.Range.Delete
        Next oFoot
    Next oSec
    
    'delete any lingering contents of regular page header
    If ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Exists = True Then
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        Selection.WholeStory
        Selection.Delete Unit:=wdCharacter, Count:=1
        Selection.WholeStory
        Selection.Delete Unit:=wdCharacter, Count:=1
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    End If
    
    'delete lingering contents of 1st page header
    If ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Exists = True Then
        ActiveWindow.ActivePane.View.SeekView = wdSeekFirstPageHeader
        Selection.WholeStory
        Selection.Delete Unit:=wdCharacter, Count:=1
        Selection.WholeStory
        Selection.Delete Unit:=wdCharacter, Count:=1
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    End If

End Sub

Sub deleteLingeringBookmarks()
'
' This bookmark is the last in the new General header,
' and tends to linger even if all headers are deleted.
' That's a problem when we need to re-process documents for
' subsequent branding changes. So...
'
    'original "old" forms:
    If ActiveDocument.Bookmarks.Exists("staff_primary_email_1") = True Then
        ActiveDocument.Bookmarks(Index:="staff_primary_email_1").Delete
    End If
    If ActiveDocument.Bookmarks.Exists("staff_job_title_pa_1") = True Then
        ActiveDocument.Bookmarks(Index:="staff_job_title_pa_1").Delete
    End If
    
    'newly processed forms:
    If ActiveDocument.Bookmarks.Exists("staff_primary_email_99") = True Then
        ActiveDocument.Bookmarks(Index:="staff_primary_email_99").Delete
    End If
    If ActiveDocument.Bookmarks.Exists("staff_job_title_pa_99") = True Then
        ActiveDocument.Bookmarks(Index:="staff_job_title_pa_99").Delete
    End If

End Sub

Sub deleteOldBrandingLines()
'
' deleteOldBrandingLines, top 6 (general) or 3 (intake) lines creating a mock-header in old documents
'
    'goto first page
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1
    
    Dim numLines As Integer
    If pStrDocType = "general" Then
        numLines = 6
    ElseIf pStrDocType = "intake" Then
        numLines = 3
    End If
    
    'delete'm
    Selection.MoveDown Unit:=wdLine, Count:=numLines, Extend:=wdExtend
    Selection.Delete Unit:=wdCharacter, Count:=1

End Sub

Sub addPagesAndOrPageNumbers()
    '''''
    'add page #s to all pages past page 1, creating/deleting a page 2 if necessary
    '

    Dim objPages As Pages
    Set objPages = ActiveDocument.ActiveWindow.Panes(1).Pages
    
    If pInitialPages = 1 Then
        'if it only comes w/ one page, temporarily add a page 2 to gain access to the page header fields;
        'skip to the end, hit Enter until you have 2 pages total
        Selection.EndKey Unit:=wdStory
        Do Until objPages.Count = 2
            Selection.TypeParagraph
        Loop
    End If
    
    'jump to page 2 and add page number headers
    addPageNumberHeader
    
    If pInitialPages = 1 Then
        'if it was orginally only one page, then go to the end and back up until there's 1 page left
        Selection.EndKey Unit:=wdStory
        Do Until objPages.Count = 1
            Selection.TypeBackspace
        Loop
    End If
End Sub

Sub addPageNumberHeader()
    '
    'add page numbers to primary headers (of page 2 and above)
    '
    'go to page 2, otherwise we'd be editing the "first page header"
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=2
    
    'open the generic page header
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    
    'custom text to appear before the page#, otherwise it ended up in a weird auto-loading place...
    Selection.TypeText Text:="Page "
    
    Application.Templates(BLDG_BLOCKS).BuildingBlockEntries("Plain Number 3").Insert Where:=Selection.Range, RichText:=True
    'Alt version of above line, using environment variable:
    'Application.Templates(Environ("APPDATA") & "\Microsoft\Document Building Blocks\1033\15\Built-In Building Blocks.dotx" _
    '    ).BuildingBlockEntries("Plain Number 3").Insert Where:=Selection.Range, _
    '    RichText:=True
     
    'delete extra CR added to header by insertion method
    Selection.EndKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, Count:=1
     
    'select it all and apply the appropriate style
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.WholeStory
    Selection.Style = ActiveDocument.Styles("Header sec 2")
    
    'exit the header
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub

Sub ToggleBookmarks()
'
' ToggleBookmarks Macro
'(not used w/in automatic rebrand processing, but useful)
'
'
    ActiveWindow.View.ShowBookmarks = Not ActiveWindow.View.ShowBookmarks
End Sub

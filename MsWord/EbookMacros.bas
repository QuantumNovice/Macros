Attribute VB_Name = "Module1"
Const commonWordFileName As String = "ebookfindreplace.txt"
Const commonWordWildcardFileName As String = "ebookwildcardreplace.txt"

Sub FixEmDashes()
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = " - ([!^013-]{3,170})-([,."" ])"
        .replacement.text = " ^+\1^+\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = " -([!^013-]{3,200})-([^013,. ])"
        .replacement.text = " ^+\1^+\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = " —([,."" ])"
        .replacement.text = "—\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub FindNextCutoffParagraph()
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find.Font
        .Bold = False
    End With
    With Selection.Find
        .text = "([!0-9.""" & ChrW(8221) & _
            "^013^058…^t\?'\!\)\(^175^148^02])[^013]{1,2}([A-Za-z])"
        .replacement.text = "\1 \2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute
End Sub

Sub FindNextHardHyphenatedWord()
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = "([a-z])- ([a-z])"
        .replacement.text = "\1\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute
End Sub

Sub WhatChrIsThis()
    MsgBox "Ascii code #" & Asc(Selection.Range.text)
End Sub

Sub ChangeAllAllCapsToSentenceCase()
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
         .ClearFormatting
         .Wrap = wdFindContinue
         .Forward = True
         .Format = False
         .MatchWildcards = True
         .text = "[A-Z]{2}[A-Z ]{1,254}"
         .replacement.text = ""
         .Execute
         While .Found
            If Not IsCapRomanNumeral(Selection.Range.text) Then
                Debug.Print "Setting """ & Selection.text & """ to sentence case."
                Selection.Range.Case = wdTitleWord
            End If
            Selection.Collapse Direction:=wdCollapseEnd
            .Execute
         Wend
     End With
End Sub

Function IsCapRomanNumeral(text As String) As Boolean
    Dim length As Integer
    Dim theChar As Integer
    length = Len(text)
    
    IsCapRomanNumeral = True
    For i = 1 To length
        theChar = Asc(Mid(text, i, 1))
        'I, V, X, L, M, <space>
        If theChar <> 86 And theChar <> 88 And theChar <> 73 And _
           theChar <> 76 And theChar <> 77 And theChar <> 32 Then
            IsCapRomanNumeral = False
            Exit Function
        End If
    Next
End Function

Sub FixCapitalizedMonthsInSpanish()

    Dim months As Variant
    months = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", _
        "Octubre", "Noviembre", "Diciembre")
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    
    For i = 0 To 11
        Selection.HomeKey Unit:=wdStory
        With Selection.Find
            .text = "([!.^058\?""""'\!\)\(^175^148^013^09]) " & months(i)
            .replacement.text = "\1 " & LCase(months(i))
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next
End Sub

Sub LoadAndExecuteFileReplacements()
    'strings to search:
    Dim findBin(255) As String
    Dim replaceBin(255) As String
    Dim replaceCount As Integer
    
    Dim fileNum As Integer
    Dim dataLine As String
    
    Const delimiter As String = vbTab
    Dim findPart As String
    Dim replacePart As String
    Dim tempDelimiter As Integer
    
    Dim filePath As String
    Set cntnr = MacroContainer
    filePath = cntnr.Path & "\" & commonWordFileName
    
    fileNum = FreeFile()
    Open filePath For Input As #fileNum
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, dataLine ' read in data 1 line at a time
        ' decide what to do with dataline,
        ' depending on what processing you need to do for each case
        'Debug.Print dataLine
        tempDelimiter = InStr(1, dataLine, delimiter)
        If tempDelimiter <> 0 Then
            findPart = Left(dataLine, tempDelimiter - 1)
            replacePart = Mid(dataLine, tempDelimiter + 1)
            If Len(findPart) > 0 And Len(replacePart) > 0 Then
                findBin(replaceCount) = findPart
                replaceBin(replaceCount) = replacePart
                replaceCount = replaceCount + 1
            End If
        End If
    Loop
    
    Close #fileNum
    
    'loop through our pairs running corresponding find/replaces
    For i = 0 To replaceCount - 1
        If i Mod 10 = 0 Then
            DoEvents
        End If
        Debug.Print "Replacing """ & findBin(i) & """ with """ & replaceBin(i) & """"
        
        Selection.Find.ClearFormatting
        Selection.Find.replacement.ClearFormatting
        With Selection.Find
            .text = findBin(i)
            .replacement.text = replaceBin(i)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next
End Sub

Sub LoadAndExecuteWildcardReplacements()
    'strings to search:
    Dim findBin(255) As String
    Dim replaceBin(255) As String
    Dim replaceCount As Integer
    
    Dim fileNum As Integer
    Dim dataLine As String
    
    Const delimiter As String = vbTab
    Dim findPart As String
    Dim replacePart As String
    Dim tempDelimiter As Integer
    
    Dim filePath As String
    Set cntnr = MacroContainer
    filePath = cntnr.Path & "\" & commonWordWildcardFileName
    
    fileNum = FreeFile()
    Open filePath For Input As #fileNum
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, dataLine ' read in data 1 line at a time
        ' decide what to do with dataline,
        ' depending on what processing you need to do for each case
        'Debug.Print dataLine
        tempDelimiter = InStr(1, dataLine, delimiter)
        If tempDelimiter <> 0 Then
            findPart = Left(dataLine, tempDelimiter - 1)
            replacePart = Mid(dataLine, tempDelimiter + 1)
            If Len(findPart) > 0 And Len(replacePart) > 0 Then
                findBin(replaceCount) = findPart
                replaceBin(replaceCount) = replacePart
                replaceCount = replaceCount + 1
            End If
        End If
    Loop
    
    Close #fileNum
    
    'loop through our pairs running corresponding find/replaces
    For i = 0 To replaceCount - 1
        Debug.Print "Replacing """ & findBin(i) & """ with """ & replaceBin(i) & """"
        
        Selection.Find.ClearFormatting
        Selection.Find.replacement.ClearFormatting
        With Selection.Find
            .text = findBin(i)
            .replacement.text = replaceBin(i)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next
End Sub

Sub AddWordReplacement()
    Const delimiter As String = vbTab
    Const fileNa As String = "C:\Users\jeffr\Documents\Sion\Libros\ebookfindreplace.txt"
    
    Dim resp
    
    Dim fileNum As Integer
    Dim dataLine As String
    
    Dim findPart As String
    Dim replacePart As String
    
    resp = MsgBox("You have selected " & Len(Selection.text) & " characters, do you want to add this word the common word replacement file?", vbYesNo)
    If resp = vbYes Then
        fileNum = FreeFile()
        findPart = Selection.text
        replacePart = InputBox("Please enter the text you want to replace. Find/replace parameters are case sensitive and find whole word only.", , findPart)
        Open fileNa For Append As #fileNum
            dataLine = Replace(findPart, delimiter, "") & delimiter & Replace(replacePart, delimiter, "")
            Debug.Print dataLine
            Print #fileNum, dataLine
        Close fileNum
    End If
End Sub

Sub OpenWordReplacementFile()
    'todo create the file if it doesn't exist
    
    On Error GoTo jic

    RetVal = Shell("C:\Program Files (x86)\Notepad++\notepad++.exe " & filePath, 1)
    Exit Sub
jic:
    RetVal = Shell("C:\WINDOWS\notepad.exe " & filePath, 1)
End Sub

Sub CommonWordReplacements()
    resp = MsgBox("Do you want to check for incorrectly capitalized months in Spanish?", vbYesNo)
    If resp = vbYes Then
        FixCapitalizedMonthsInSpanish
    End If
    
    resp = MsgBox("Do you want to load the find/replace file from the Libros folder to cleanup common spelling/grammar errors?", vbYesNo)
    If resp = vbYes Then
        LoadAndExecuteFileReplacements
        
        LoadAndExecuteWildcardReplacements
    End If
    
End Sub


Sub MakeFootNotesAuto()
    Dim f As Footnote

    For Each f In ActiveDocument.Footnotes
        f.Range.FootnoteOptions.NumberingRule _
          = wdRestartContinuous
    Next f
End Sub

Sub ParagraphBeforeHeadingDontKeepWithNext()
    Dim lastP
    Dim count
    count = 0
    For Each p In ActiveDocument.Paragraphs
        count = count + 1
        If p.Style = ActiveDocument.Styles("Heading 2") Or p.Style = ActiveDocument.Styles("Heading 1") Then
            Debug.Print count
            lastP.Format.KeepWithNext = False
        End If
        Set lastP = p
    Next
    
End Sub

Sub AddChapterHyperlinks()
    'gather list of h1 p's
    'bookmark each h2
    'gather list of h2 p ranges with corresponding text
    'write at the beginning of each h1, hyperlinks to the h2s underneath
    Dim lastP
    Dim count
    count = 0
    Dim lastH1 As Range
    
    Dim h2s(100) As Range
    Dim isH3(100) As Boolean
    Dim h2Count As Integer
    Dim theText As String
    Dim textToAdd As String
    Dim formattedBookmark As Variant
    Dim hyperlinkCount As Integer
    
    If MsgBox("This script will add hyperlinks at the beginning of each chapter (immediately after each Heading 1), " & _
        "do you want to continue?", vbYesNo) <> vbYes Then
        Exit Sub
    End If
    
    If MsgBox("We usually eliminate any preexisting custom bookmarks to avoid conflicts, is this OK?", vbYesNo) = vbYes Then
        EliminateBookmarks
    Else
        MsgBox "OK, we'll leave the old ones alone.", vbInformation
    End If
    
    BookmarkHeadings2
    
    h2Count = 0
    For Each p In ActiveDocument.Paragraphs
        count = count + 1
        If p.Style = ActiveDocument.Styles("Heading 1") Then
            Debug.Print p.Range.text
            'it's time to write the last h1's hyperlinks
            If (h2Count > 0) Then
                counter = h2Count - 1
                For i = 0 To counter
                    theText = h2s(i).text
                    formattedBookmark = FormatBookmarkName(theText)
                    Debug.Print "Let's write some hyperlinks: " & theText & "," & formattedBookmark
                    If i = 0 Then
                        lastH1.Select
                        Selection.EndKey Unit:=wdLine
                        Selection.TypeParagraph
                    End If
                    hyperlinkCount = hyperlinkCount + 1
                    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
                        SubAddress:=formattedBookmark, ScreenTip:="", TextToDisplay:=theText
                    If isH3(h2Count) = True Then
                        Selection.Paragraphs.Indent
                    End If
                Next i
            End If
            Set lastH1 = p.Range
            h2Count = 0
        ElseIf p.OutlineLevel = wdOutlineLevel2 Or p.OutlineLevel = wdOutlineLevel3 Then
            Set h2s(h2Count) = p.Range
            If p.OutlineLevel = wdOutlineLevel3 Then
                isH3(h2Count) = True
            End If
            h2Count = h2Count + 1
        End If
    Next
    
    'put in the last set of bookmarks
    If (h2Count > 0) Then
        counter = h2Count - 1
        For i = 0 To counter
            theText = h2s(i).text
            formattedBookmark = FormatBookmarkName(theText)
            Debug.Print "Let's write some hyperlinks: " & theText & "," & formattedBookmark
            If i = 0 Then
                lastH1.Select
                Selection.EndKey Unit:=wdLine
                Selection.TypeParagraph
            End If
            hyperlinkCount = hyperlinkCount + 1
            ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
                SubAddress:=formattedBookmark, ScreenTip:="", TextToDisplay:=theText
        Next i
    End If
    MsgBox "All set, we wrote " & Trim(CStr(hyperlinkCount)) & " hyperlinks.", vbInformation
End Sub

Sub AddIndexHyperlinks()
    'presupposes the BookmarkHeadings2 has been called
    
    'gather list of h1 p's
    'bookmark each h2
    'gather list of h2 p ranges with corresponding text
    'write at the beginning of each h1, hyperlinks to the h2s underneath
    Dim lastP
    Dim count
    count = 0
    Dim lastH1 As Range
    
    Dim h2s(100)
    Dim h2Count
    Dim theText As String
    Dim textToAdd As String
    Dim formattedBookmark As Variant
    
    h2Count = 0
    For Each p In ActiveDocument.Paragraphs
        count = count + 1
        If p.Style = ActiveDocument.Styles("Heading 1") Then
            If Left(lastH1, 5) = "Index" Then
                'it's time to write the last h1's hyperlinks
                If (h2Count > 0) Then
                    counter = h2Count - 1
                    For i = 0 To counter
                        theText = h2s(i).text
                        formattedBookmark = FormatBookmarkName(theText)
                        Debug.Print "Let's write some hyperlinks: " & theText & "," & formattedBookmark
                        If i = 0 Then
                            lastH1.Select
                            Selection.EndKey Unit:=wdLine
                            Selection.TypeParagraph
                        End If
    
                        ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
                            SubAddress:=formattedBookmark, ScreenTip:="", TextToDisplay:=theText
                        
                    Next i
                End If
            End If
            Set lastH1 = p.Range
            h2Count = 0
        ElseIf p.Style = ActiveDocument.Styles("Heading 2") Then
            Set h2s(h2Count) = p.Range
            h2Count = h2Count + 1
        End If
    Next
    
End Sub

Sub BookmarkIndex()
    
    For Each p In ActiveDocument.Paragraphs
    
        If p.Style = ActiveDocument.Styles("Heading 2") And Len(p.Range.text) <= 2 Then
            Debug.Print FormatBookmarkName(p.Range.text)
            ActiveDocument.Bookmarks.Add FormatBookmarkName(p.Range.text), p.Range
        End If
    Next
End Sub

Sub BookmarkHeadings2()
    EliminateBookmarks
    
    For Each p In ActiveDocument.Paragraphs
    
        If p.Style = ActiveDocument.Styles("Heading 2") Then
            'Debug.Print FormatBookmarkName(p.Range.Text)
            ActiveDocument.Bookmarks.Add FormatBookmarkName(p.Range.text), p.Range
        End If
    Next
End Sub

Function FormatChapterBookmarkName(theText As String)
    Dim result As String
    result = "chapter_"
    'find the number within the text
    For i = 1 To Len(theText)
        temp = Asc(Mid(theText, i, 1))
        If temp >= 48 And temp < 58 Then
            result = result & Chr(temp)
        Else
            Exit For
        End If
    Next i
    FormatChapterBookmarkName = result
End Function

Function FormatQuestionBookmarkName(theText As String)
    Dim result As String
    result = "question_"
    'find the number within the text
    For i = 1 To Len(theText)
        temp = Asc(Mid(theText, i, 1))
        If temp >= 48 And temp < 58 Then
            result = result & Chr(temp)
        Else
            Exit For
        End If
    Next i
    FormatQuestionBookmarkName = result
End Function

Function FormatBookmarkName(theText As String)
    Dim result
    Dim ascii
    theText = UCase(theText)
    result = "b"
    For i = 1 To Len(theText)
        ascii = Asc(Mid(theText, i, 1))
        If (ascii >= 48 And ascii <= 57) Or (ascii >= 65 And ascii <= 90) Then
            result = result & Chr(ascii)
        End If
    Next
    FormatBookmarkName = LCase(Left(result, 10))
End Function

Sub EliminateBookmarks()
    For Each b In ActiveDocument.Bookmarks
        If b.Name <> "toc" And Left(b.Name, 1) <> "_" Then
            b.Delete
        End If
    Next
End Sub

Sub DeleteBlankParagraphAroundHeadings()
    Dim lastWasH As Boolean
    Dim latestPs(50) As Range
    Dim pCount As Integer
    Dim isBlank As Boolean
    Dim lastH As String
    Dim count As Integer
    
    count = 0
    lastWasH = False
    lastH = ""
    For Each p In ActiveDocument.Paragraphs
        count = count + 1
        isBlank = (Len(p.Range.text) = 1 And Asc(p.Range.text) = 13)
        
        If lastWasH And isBlank Then
            Debug.Print lastH
            p.Range.text = ""
        ElseIf p.OutlineLevel = wdOutlineLevel1 Or p.OutlineLevel = wdOutlineLevel2 Then
            lastWasH = True
            lastH = p.Range.text
            For i = 0 To pCount - 1
                Debug.Print lastH
                latestPs(i).text = ""
            Next
        ElseIf isBlank Then
            Set latestPs(pCount) = p.Range
            pCount = pCount + 1
        Else
            lastWasH = False
            lastH = ""
            pCount = 0
        End If
    Next
End Sub

Sub PutManualPageBreakBeforeHeadings()
'
' Macro4 Macro
'
'
    For Each p In ActiveDocument.Paragraphs
        count = count + 1
        If p.OutlineLevel = wdOutlineLevel1 Or p.OutlineLevel = wdOutlineLevel2 Then
            p.Range.Select
            Selection.HomeKey Unit:=wdLine
            Selection.MoveLeft Unit:=wdCharacter, count:=1
            Selection.TypeParagraph
            Selection.InsertBreak Type:=0
            Selection.Delete Unit:=wdCharacter, count:=1
        End If
    Next
End Sub

Sub FormatHyperlinks()
    For Each h In ActiveDocument.Hyperlinks
        Debug.Print h.Address & "|" & h.SubAddress
        h.Range.Font.Underline = wdUnderlineSingle
        h.Range.Font.TextColor = wdColorGray625
        h.Range.Case = wdTitleWord
    Next
End Sub

Sub EliminateHyperlinks()
    For Each h In ActiveDocument.Hyperlinks
        Debug.Print h.Address & "|" & h.SubAddress
        h.Delete
    Next
End Sub

Sub ExtrapolateFootnotes()
'
' Macro1 Macro
'
'
    Dim lastSelection As Range
    Dim numberOfFootnotes As Integer
    Dim footnoteRanges(200) As Range
    Dim superscript As Boolean
    tempResponse = InputBox("How many footnotes should we look for?")
    If tempResponse = "" Then
        Exit Sub
    End If
    numberOfFootnotes = CInt(tempResponse)
    If MsgBox("Should we only look in superscripts?", vbYesNo) = vbYes Then
        Selection.Find.ClearFormatting
        With Selection.Find.Font
            .superscript = True
            .Subscript = False
        End With
        Selection.Find.replacement.ClearFormatting
        With Selection.Find
            .replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
    Else
        Selection.Find.ClearFormatting
        Selection.Find.replacement.ClearFormatting
        With Selection.Find
            .text = Trim(CStr(i))
            .replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
    End If
    Selection.HomeKey Unit:=wdStory
    Set lastSelection = Selection.Range
    'go through the document always looking for the first occurance of
    'the next number.  If the user decides not to mark a footnote we
    'should select the previous selection to avoid skipping over a
    'whole bunch of numbers
    For i = 1 To numberOfFootnotes
        Do
            Selection.Find.text = Trim(CStr(i)) & ">"
            Selection.Find.Execute
            response = MsgBox("Should we convert the selected number into a blank footnote? " & _
            "If you click no we will continue searching for the same number. If you click cancel " & _
            "we will start searching for the same number. Press Ctrl-X to exit the macro.", vbYesNoCancel)
            If response = vbYes Then 'Add a blank footnote where the selected text was
                Selection.TypeBackspace
                Set lastSelection = Selection.Range
                With Selection
                    With .FootnoteOptions
                        .Location = wdBottomOfPage
                        .NumberingRule = wdRestartContinuous
                        .StartingNumber = 1
                        .NumberStyle = wdNoteNumberStyleArabic
                        .LayoutColumns = 0
                    End With
                    .Footnotes.Add Range:=Selection.Range, Reference:=""
                    Set footnoteRanges(i) = Selection.Range
                End With
            End If
        Loop Until response = vbYes Or response = vbCancel
        lastSelection.Select
    Next
End Sub

Sub StartingBookFromWpd()
'
' StartingBookFromWpd Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = "^b"
        .replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

        Selection.PageSetup.Orientation = wdOrientPortrait
    With Selection.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(1.27)
        .BottomMargin = CentimetersToPoints(1.27)
        .LeftMargin = CentimetersToPoints(1.27)
        .RightMargin = CentimetersToPoints(1.27)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.27)
        .FooterDistance = CentimetersToPoints(1.27)
        .PageWidth = CentimetersToPoints(21.59)
        .PageHeight = CentimetersToPoints(27.94)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = True
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
    End With
    ActiveWindow.ActivePane.View.ShowAll = Not ActiveWindow.ActivePane.View. _
        ShowAll
    ActiveDocument.ApplyQuickStyleSet2 ("Black & White (Classic)")
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = "^-"
        .replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = """"
        .replacement.text = """"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    Selection.WholeStory
    Selection.Delete Unit:=wdCharacter, count:=1
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    
    SetNormalStyle
End Sub

Sub FixBlankParagraphs()
'
' FixBlankParagraphs Macro
'
'
    If MsgBox("This operation will eliminate blank paragraphs around headings 1 and 2 " & _
        "and consolidate multiple blank spaces transforming 1 blank into none, " & _
        "2 blanks into 1, and 3+ blanks into 2. Are you sure you want to continue?", vbYesNo) <> vbYes Then
        Exit Sub
    End If
    'remove blank space characters from otherwise blank paragraphs
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "[ ^09]{1,200}[^013]"
        .replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    'delete blank paragraphs around headings 1 and 2
    DeleteBlankParagraphAroundHeadings

    'Prep big spaces
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = "[^013]{4,50}"
        .replacement.text = "^p{2blankparagraphs}^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    'Selection.Find.Execute
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'prep small spaces in text
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = "^p^p^p"
        .replacement.text = "^p{1blankparagraph}^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    'Selection.Find.Execute
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'protect double paragraphs in verse areas
    Selection.Find.ClearFormatting
    With Selection.Find.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = "^p^p"
        .replacement.text = "^p{1blankparagraph}^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    
    'eliminate formerlly small spaces
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = "^p^p"
        .replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'finalize big spaces
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = "{2blankparagraphs}"
        .replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'finalize small spaces
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = "{1blankparagraph}"
        .replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    If MsgBox("Finished removing excess blank paragraphs. Do you want to add manual page breaks before headings 1 and 2?") = vbYes Then
        PutManualPageBreakBeforeHeadings
    End If
    
    MsgBox "Done. As a next step, prepare to delete all tabs and double spaces.", vbInformation
End Sub

Sub ReplaceTabsWithSpace()
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = "^t"
        .replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub EliminateMultipleSpaces()
    'Eliminate spaces and tabs at the end of paragraphs
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "[ ^09]{1,200}[^013]"
        .replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = " {2,200}"
        .replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub PlaceIndexAtEnd()
'
' PlaceIndexAtEnd Macro
'
'
    If MsgBox("This script will place an index at the end of the file that will be recognizable by Kindle. " & _
        "This assumes you don't already have an automatic Table of Contents in your document. " & _
        "Do you want to continue?", vbYesNo) <> vbYes Then
        Exit Sub
    End If
    
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.TypeText text:="Index"
    Selection.Style = ActiveDocument.Styles("Heading 1")
    Selection.TypeParagraph
    With ActiveDocument
        .TablesOfContents.Add Range:=Selection.Range, RightAlignPageNumbers:= _
            True, UseHeadingStyles:=True, UpperHeadingLevel:=1, _
            LowerHeadingLevel:=2, IncludePageNumbers:=False, AddedStyles:="", _
            UseHyperlinks:=True, HidePageNumbersInWeb:=True, UseOutlineLevels:= _
            True
        .TablesOfContents(1).TabLeader = wdTabLeaderDots
        .TablesOfContents.Format = wdIndexIndent
    End With
End Sub

Sub EliminatePageNumbers()
'
' Macro5 Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = "[^013]{1,4}?oneparagraph? {1,200}[0-9]{1,3}[^013]{1,4}"
        .replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = "[^013]{1,4}[0-9]{1,3}[^013]{1,4}"
        .replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub ConvertToVerseStyles()
    If Not StyleExists("Verse") Then
        ActiveDocument.Styles.Add Name:="Verse", Type:=wdStyleTypeParagraph
        ActiveDocument.Styles("Verse").AutomaticallyUpdate = False
        With ActiveDocument.Styles("Verse").Font
            .Name = "+Body"
            .Size = 11
            .Bold = False
            .Italic = False
            .Underline = wdUnderlineNone
            .UnderlineColor = wdColorAutomatic
            .StrikeThrough = False
            .DoubleStrikeThrough = False
            .Outline = False
            .Emboss = False
            .Shadow = False
            .Hidden = False
            .SmallCaps = False
            .AllCaps = False
            .Color = wdColorAutomatic
            .Engrave = False
            .superscript = False
            .Subscript = False
            .Scaling = 100
            .Kerning = 0
            .Animation = wdAnimationNone
            .Ligatures = wdLigaturesNone
            .NumberSpacing = wdNumberSpacingDefault
            .NumberForm = wdNumberFormDefault
            .StylisticSet = wdStylisticSetDefault
            .ContextualAlternates = 0
        End With
        With ActiveDocument.Styles("Verse").ParagraphFormat
            .LeftIndent = CentimetersToPoints(0.5)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = LinesToPoints(1.08)
            .Alignment = wdAlignParagraphLeft
            .WidowControl = True
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = CentimetersToPoints(0)
            .OutlineLevel = wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .MirrorIndents = False
            .TextboxTightWrap = wdTightNone
            .CollapsedByDefault = False
        End With
        ActiveDocument.Styles("Verse").NoSpaceBetweenParagraphsOfSameStyle = False
        ActiveDocument.Styles("Verse").ParagraphFormat.TabStops.ClearAll
        With ActiveDocument.Styles("Verse").ParagraphFormat
            With .Shading
                .Texture = wdTextureNone
                .ForegroundPatternColor = wdColorAutomatic
                .BackgroundPatternColor = wdColorAutomatic
            End With
            .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
            .Borders(wdBorderRight).LineStyle = wdLineStyleNone
            .Borders(wdBorderTop).LineStyle = wdLineStyleNone
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
            .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
            With .Borders
                .DistanceFromTop = 1
                .DistanceFromLeft = 4
                .DistanceFromBottom = 1
                .DistanceFromRight = 4
                .Shadow = False
            End With
        End With
        ActiveDocument.Styles("Verse").Frame.Delete
    End If
    
    Selection.Style = ActiveDocument.Styles("Verse")
    
    'Delete tabs and spaces from the beginning of the line
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = "[^013][ ^09]{1,6}"
        .replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    With Selection.Find.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = "^p^p"
        .replacement.text = "^p{1blankparagraph}^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
End Sub

Function StyleExists(style_name As String) As Boolean
     
    style_exists = False
    On Error Resume Next
    StyleExists = ActiveDocument.Styles(style_name).NameLocal = style_name
     
End Function

Function DoesStyleExist(styleName As String) As Boolean
    DoesStyleExist = False
    For Each s In ActiveDocument.Styles
        Debug.Print s.NameLocal
        If s.NameLocal = styleName Then
            DoesStyleExist = True
            Exit Function
        End If
    Next
End Function

Sub FormatSelectedAsIndex()
    If Not DoesStyleExist("Index") Then
        ActiveDocument.Styles.Add Name:="Index", Type:=wdStyleTypeParagraph
    End If
    
    ActiveDocument.Styles("Index").AutomaticallyUpdate = True
    With ActiveDocument.Styles("Index").Font
        .Name = "+Body"
        .Size = 12
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Index").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.08)
        .Alignment = wdAlignParagraphJustify
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
    ActiveDocument.Styles("Index").NoSpaceBetweenParagraphsOfSameStyle = False
    ActiveDocument.Styles("Index").ParagraphFormat.TabStops.ClearAll
    ActiveDocument.Styles("Index").ParagraphFormat.TabStops.Add Position:= _
        CentimetersToPoints(0.5), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    ActiveDocument.Styles("Index").ParagraphFormat.TabStops.Add Position:= _
        CentimetersToPoints(1), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    ActiveDocument.Styles("Index").ParagraphFormat.TabStops.Add Position:= _
        CentimetersToPoints(1.5), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    ActiveDocument.Styles("Index").ParagraphFormat.TabStops.Add Position:= _
        CentimetersToPoints(2), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    ActiveDocument.Styles("Index").ParagraphFormat.TabStops.Add Position:= _
        CentimetersToPoints(2.5), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    ActiveDocument.Styles("Index").ParagraphFormat.TabStops.Add Position:= _
        CentimetersToPoints(3), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    ActiveDocument.Styles("Index").ParagraphFormat.TabStops.Add Position:= _
        CentimetersToPoints(3.5), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    With ActiveDocument.Styles("Index").ParagraphFormat
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorAutomatic
        End With
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    ActiveDocument.Styles("Index").LanguageID = wdSpanishModernSort
    ActiveDocument.Styles("Index").NoProofing = False
    ActiveDocument.Styles("Index").Frame.Delete
    Selection.Style = ActiveDocument.Styles("Index")
End Sub

Sub BoldAllCapsText()
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    Selection.Find.replacement.Font.Bold = True
    With Selection.Find
        .text = "([A-Z]{2}[A-Z ]{1,50})"
        .replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub ListPlain()
    Dim lp As Paragraph

    For Each lp In ActiveDocument.ListParagraphs
        Debug.Print lp.Range.text
        lp.Range.ListFormat.ConvertNumbersToText
    Next lp
End Sub

Sub SetNormalStyle()
'
' Macro7 Macro
'
'
    ActiveDocument.Styles("Normal").AutomaticallyUpdate = False
    With ActiveDocument.Styles("Normal").Font
        .Name = "+Body"
        .Size = 11
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Normal").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 8
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.08)
        .Alignment = wdAlignParagraphJustify
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0.75)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
    ActiveDocument.Styles("Normal").NoSpaceBetweenParagraphsOfSameStyle = _
        False
    ActiveDocument.Styles("Normal").ParagraphFormat.TabStops.ClearAll
    With ActiveDocument.Styles("Normal").ParagraphFormat
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorAutomatic
        End With
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    ActiveDocument.Styles("Normal").LanguageID = wdSpanishModernSort
    ActiveDocument.Styles("Normal").NoProofing = False
    ActiveDocument.Styles("Normal").Frame.Delete
End Sub

Sub ReapplyHeadingStyles()
    
    If MsgBox("This script will apply the proper heading style to paragraphs according to their Outline Level. " & _
        "Do you want to continue?", vbYesNo) <> vbYes Then
        Exit Sub
    End If
    
    For Each p In ActiveDocument.Paragraphs
        Select Case p.OutlineLevel
            Case wdOutlineLevelBodyText
            Case wdOutlineLevel1
                p.Style = ActiveDocument.Styles("Heading 1")
            Case wdOutlineLevel2
                p.Style = ActiveDocument.Styles("Heading 2")
            Case wdOutlineLevel3
                p.Style = ActiveDocument.Styles("Heading 3")
            Case wdOutlineLevel4
                p.Style = ActiveDocument.Styles("Heading 4")
        End Select
    Next
End Sub

Sub FixQuotes()
'
' Macro7 Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .text = ChrW(8220)
        .replacement.text = ChrW(8220)
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "'"
        .replacement.text = "'"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub



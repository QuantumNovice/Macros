<<<<<<< HEAD
Attribute VB_Name = "NewMacros"
Sub Macro1()
'
' Macro1 Macro
'
'
    With Selection.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = InchesToPoints(0.5)
        .BottomMargin = InchesToPoints(0.5)
        .LeftMargin = InchesToPoints(0.5)
        .RightMargin = InchesToPoints(0.5)
        .Gutter = InchesToPoints(0)
        .HeaderDistance = InchesToPoints(0.5)
        .FooterDistance = InchesToPoints(0.5)
        .PageWidth = InchesToPoints(8.5)
        .PageHeight = InchesToPoints(11)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
        .SectionDirection = wdSectionDirectionLtr
    End With
    With Selection.Sections(1)
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFrom = wdBorderDistanceFromPageEdge
            .AlwaysInFront = True
            .SurroundHeader = True
            .SurroundFooter = True
            .JoinBorders = False
            .DistanceFromTop = 24
            .DistanceFromLeft = 24
            .DistanceFromBottom = 24
            .DistanceFromRight = 24
            .Shadow = False
            .EnableFirstPageInSection = True
            .EnableOtherPagesInSection = True
            .ApplyPageBordersToAllSections
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
End Sub

Sub AssignmentLayoutGen()
Attribute AssignmentLayoutGen.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Assi"
'
' Assi Macro
'
'
    Call Macro1
    AssignmentNo = InputBox(prompt, "Assignment No", "1")
    Subject = InputBox(prompt, "Subject", "Eng Mechanics")
    RollNo = "45"
    RegNo = "18PWCIVxyz"
    Author = "Syed Haseeb Shah"
    
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Name: "
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:=Author
'
    Selection.TypeText Text:=vbLf
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Roll No: "
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:=RollNo
'
    Selection.TypeText Text:=vbLf
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Subject: "
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:=Subject
    
    Selection.TypeText Text:=vbLf
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Date: "
    Selection.Font.Bold = wdToggle
    Selection.InsertDateTime DateTimeFormat:="M/d/yyyy", InsertAsField:=False, _
         DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
        InsertAsFullWidth:=False
    Selection.TypeText Text:=vbLf
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Registration: "
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:=RegNo
End Sub

=======
Attribute VB_Name = "NewMacros"
Sub Macro1()
'
' Macro1 Macro
'
'
    With Selection.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = InchesToPoints(0.5)
        .BottomMargin = InchesToPoints(0.5)
        .LeftMargin = InchesToPoints(0.5)
        .RightMargin = InchesToPoints(0.5)
        .Gutter = InchesToPoints(0)
        .HeaderDistance = InchesToPoints(0.5)
        .FooterDistance = InchesToPoints(0.5)
        .PageWidth = InchesToPoints(8.5)
        .PageHeight = InchesToPoints(11)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
        .SectionDirection = wdSectionDirectionLtr
    End With
    With Selection.Sections(1)
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFrom = wdBorderDistanceFromPageEdge
            .AlwaysInFront = True
            .SurroundHeader = True
            .SurroundFooter = True
            .JoinBorders = False
            .DistanceFromTop = 24
            .DistanceFromLeft = 24
            .DistanceFromBottom = 24
            .DistanceFromRight = 24
            .Shadow = False
            .EnableFirstPageInSection = True
            .EnableOtherPagesInSection = True
            .ApplyPageBordersToAllSections
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
End Sub

Sub AssignmentLayoutGen()
Attribute AssignmentLayoutGen.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Assi"
'
' Assi Macro
'
'
    Call Macro1
    AssignmentNo = InputBox(prompt, "Assignment No", "1")
    Subject = InputBox(prompt, "Subject", "Eng Mechanics")
    RollNo = "45"
    RegNo = "18PWCIVxyz"
    Author = "Syed Haseeb Shah"
    
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Name: "
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:=Author
'
    Selection.TypeText Text:=vbLf
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Roll No: "
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:=RollNo
'
    Selection.TypeText Text:=vbLf
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Subject: "
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:=Subject
    
    Selection.TypeText Text:=vbLf
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Date: "
    Selection.Font.Bold = wdToggle
    Selection.InsertDateTime DateTimeFormat:="M/d/yyyy", InsertAsField:=False, _
         DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
        InsertAsFullWidth:=False
    Selection.TypeText Text:=vbLf
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Registration: "
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:=RegNo
End Sub

>>>>>>> daa931be89e26974d59dbe5bf462446143caa10a

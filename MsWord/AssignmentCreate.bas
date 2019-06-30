<<<<<<< HEAD
Attribute VB_Name = "NewMacros"

Sub Assi()
Attribute Assi.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Assi"
'
' Assi Macro
'
'
    AssignmentNo = InputBox(prompt, "AssignmentNo", "01")
    Subject = InputBox(prompt, "Subject", "Eng Mechanics")
    RollNo = "45"
    RegNo = "18PWCIVxyz"
    Author = "Syed Haseeb Shah"
    
    
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Name: "
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:=Author
    
    Selection.TypeText Text:=vbLf
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Roll No: "
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:=RollNo
    
    Selection.TypeText Text:=vbLf
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Subject: "
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:=Subject
    
    Selection.TypeText Text:=vbLf
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Date: "
    Selection.Font.Bold = wdToggle
    Selection.InsertDateTime DateTimeFormat:="M/d/yyyy", InsertAsField:=False, _
         DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
        InsertAsFullWidth:=False
    Selection.TypeText Text:=vbLf
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Subject: "
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:=Subject
    
    
End Sub

=======
Attribute VB_Name = "NewMacros"

Sub Assi()
Attribute Assi.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Assi"
'
' Assi Macro
'
'
    AssignmentNo = InputBox(prompt, "AssignmentNo", "01")
    Subject = InputBox(prompt, "Subject", "Eng Mechanics")
    RollNo = "45"
    RegNo = "18PWCIVxyz"
    Author = "Syed Haseeb Shah"
    
    
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Name: "
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:=Author
    
    Selection.TypeText Text:=vbLf
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Roll No: "
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:=RollNo
    
    Selection.TypeText Text:=vbLf
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Subject: "
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:=Subject
    
    Selection.TypeText Text:=vbLf
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Date: "
    Selection.Font.Bold = wdToggle
    Selection.InsertDateTime DateTimeFormat:="M/d/yyyy", InsertAsField:=False, _
         DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
        InsertAsFullWidth:=False
    Selection.TypeText Text:=vbLf
    Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="Subject: "
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:=Subject
    
    
End Sub

>>>>>>> daa931be89e26974d59dbe5bf462446143caa10a

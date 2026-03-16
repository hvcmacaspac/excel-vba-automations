Attribute VB_Name = "Module1"
Sub CleanPrevTSSheet()
    Dim wsPrev As Worksheet
    Set wsPrev = ThisWorkbook.Sheets("Prev TS")

    On Error Resume Next
    wsPrev.Columns("G:I").Delete Shift:=xlToLeft
    On Error GoTo 0
End Sub

Sub UpdateTrainingScoreSheet()
    Dim wsTS As Worksheet
    Dim lastRow As Long
    Set wsTS = ThisWorkbook.Sheets("Training Score by Project")

    ' Add columns for audit and conditional formatting
    
    wsTS.Columns("G:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    wsTS.Range("G1").Value = "SLA"
    wsTS.Range("H1").Value = "Previous TS"
    wsTS.Range("I1").Value = "Did Score Change"

    With wsTS.Range("G1:I1")
        .Interior.Color = RGB(255, 255, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With

    wsTS.Columns("G:I").ColumnWidth = 17
    
    ' Formulas

    lastRow = wsTS.Cells(wsTS.Rows.Count, "B").End(xlUp).Row
    wsTS.Range("G2").Formula = "=IF(F2<=30,""Needs Review"",""Good"")"
    wsTS.Range("H2").Formula = "=IFNA(XLOOKUP(C2,'Prev TS'!C:C,'Prev TS'!F:F),""New Project"")"
    wsTS.Range("I2").Formula = "=IF(H2=""New Project"",""New Project"",IF(F2=H2,""No"",""Yes""))"
    wsTS.Range("G2:I2").AutoFill Destination:=wsTS.Range("G2:I" & lastRow)
End Sub


Sub UpdateQIACandidatesSheet()
    Dim wsCdds As Worksheet
    Dim lastRow As Long
    Dim startDate As Date, endDate As Date

    ' Cdds = short for Candidates
    
    Set wsCdds = ThisWorkbook.Sheets("QIA Candidates")

    ' Add columns for audit and conditional formatting

    wsCdds.Range("H1").Value = "In ATS via Name"
    wsCdds.Range("I1").Value = "In ATS via Email"
    wsCdds.Range("J1").Value = "ATS Audit"
    wsCdds.Range("K1").Value = "Notes"

    With wsCdds.Range("H1:K1")
        .Interior.Color = RGB(255, 255, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With

    wsCdds.Columns("H:I").ColumnWidth = 40
    wsCdds.Columns("J:K").ColumnWidth = 17
    
    ' Formulas
    
    lastRow = wsCdds.Cells(wsCdds.Rows.Count, "B").End(xlUp).Row
    wsCdds.Range("H2").Formula = "=VLOOKUP(C2,'ATS Statuses'!A:A,1,FALSE)"
    wsCdds.Range("I2").Formula = "=VLOOKUP(D2,'ATS Statuses'!B:B,1,FALSE)"
    wsCdds.Range("H2:I2").AutoFill Destination:=wsCdds.Range("H2:I" & lastRow)

    
End Sub

Sub UpdateKATSCandidatesSheet()
    Dim wsKATS As Worksheet
    Dim lastRow As Long
    Set wsKATS = ThisWorkbook.Sheets("K-ATS Candidates")

    'Conditional formatting and formula
    
    wsKATS.Columns("R:R").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    With wsKATS.Range("R1")
        .Value = "SLA"
        .Interior.Color = RGB(255, 255, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With

    'Formulas
    
    lastRow = wsKATS.Cells(wsKATS.Rows.Count, "O").End(xlUp).Row
    wsKATS.Range("R2").FormulaR1C1 = _
        "=IF(AND(RC[-3]=""Conversation Booked"",RC[1]<=2),""Within SLA"",IF(RC[1]<=1,""Within SLA"",""Outside SLA""))"
    wsKATS.Range("R2").AutoFill Destination:=wsKATS.Range("R2:R" & lastRow)
End Sub

Sub UpdateAndCopyAllTasks()
    Dim wsAll As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long

    Set wsAll = ThisWorkbook.Sheets("All Tasks")
    Set wsTarget = ThisWorkbook.Sheets("Completed Tasks by Sourcer")

    ' Update All Tasks Sheet
    
    wsAll.Columns("I").Insert Shift:=xlToRight
    wsAll.Columns("K").Insert Shift:=xlToRight

    wsAll.Columns("H").TextToColumns Destination:=wsAll.Range("H1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(10, 1))
    wsAll.Columns("J").TextToColumns Destination:=wsAll.Range("J1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(10, 1))
    wsAll.Columns("L").TextToColumns Destination:=wsAll.Range("L1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(10, 1))

    wsAll.Range("M1").ClearContents
    wsAll.Range("L1").Value = "Last Update"

    wsAll.Columns("I:K").Delete Shift:=xlToLeft

    wsAll.Columns("H:J").ColumnWidth = 13.55

    wsAll.Columns("J").Insert Shift:=xlToRight
    With wsAll.Range("J1")
        .Value = "SLA"
        .Interior.Color = RGB(255, 255, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With
    wsAll.Columns("J").ColumnWidth = 17

    lastRow = wsAll.Cells(wsAll.Rows.Count, "H").End(xlUp).Row
    If lastRow >= 2 Then
        wsAll.Range("J2").FormulaR1C1 = "=IF(RC[-1]<RC[1],""Outside SLA"",""Within SLA"")"
        wsAll.Range("J2").AutoFill Destination:=wsAll.Range("J2:J" & lastRow)
    End If

    ' Copy DONE Rows to Completed Sheet

    lastRow = wsAll.Cells(wsAll.Rows.Count, "A").End(xlUp).Row
    With wsAll.Range("A1:K" & lastRow)
        .AutoFilter Field:=6, Criteria1:="DONE"
        On Error Resume Next
        .SpecialCells(xlCellTypeVisible).Copy
        On Error GoTo 0
    End With

    wsTarget.Range("A1").PasteSpecial xlPasteValues
    wsTarget.Range("A1").PasteSpecial xlPasteFormats

    Application.CutCopyMode = False
    wsAll.AutoFilterMode = False
End Sub



Sub Run_All_Updates()
    CleanPrevTSSheet
    UpdateTrainingScoreSheet
    UpdateQIACandidatesSheet
    UpdateKATSCandidatesSheet
    UpdateAndCopyAllTasks
    
    MsgBox "All sheets updated! Please double check all data and formulas.", vbInformation, "Update Complete"
    
End Sub

Sub UpdateWeekXX()
    Dim wsQIA As Worksheet, wsAllTasks As Worksheet, wsTraining As Worksheet, wsTarget As Worksheet
    Dim lastRowQIA As Long, lastRowAll As Long, lastRowTraining As Long
    Dim totalQIA As Long, goodQIA As Long
    Dim totalAllDone As Long, withinSLAAll As Long
    Dim totalTraining As Long, goodTraining As Long
    Dim pctQIAGood As Double, pctAllSLA As Double, pctTrainingGood As Double
    Dim notesVal As String, statusAllVal As String, slaAllVal As String, trainingVal As String
    Dim i As Long
    Dim weekDate As String

    Set wsQIA = ThisWorkbook.Sheets("QIA Candidates")
    Set wsAllTasks = ThisWorkbook.Sheets("All Tasks")
    Set wsTraining = ThisWorkbook.Sheets("Training Score by Project")
    Set wsTarget = ThisWorkbook.Sheets("Week XX")

    weekDate = Format(Date - 7, "mmm d")

    ' QIA Candidates
    
    lastRowQIA = wsQIA.Cells(wsQIA.Rows.Count, "H").End(xlUp).Row
    totalQIA = 0
    goodQIA = 0
    For i = 2 To lastRowQIA
        notesVal = Trim(wsQIA.Cells(i, "K").Value)
        If notesVal <> "" Then
            totalQIA = totalQIA + 1
            If StrComp(notesVal, "Good", vbTextCompare) = 0 Then
                goodQIA = goodQIA + 1
            End If
        End If
    Next i
    pctQIAGood = IIf(totalQIA = 0, 0, goodQIA / totalQIA)

    ' All Tasks
    
    lastRowAll = wsAllTasks.Cells(wsAllTasks.Rows.Count, "F").End(xlUp).Row
    totalAllDone = 0
    withinSLAAll = 0
    For i = 2 To lastRowAll
        statusAllVal = Trim(wsAllTasks.Cells(i, "F").Value)
        slaAllVal = Trim(wsAllTasks.Cells(i, "J").Value)
        If StrComp(statusAllVal, "Done", vbTextCompare) = 0 Then
            totalAllDone = totalAllDone + 1
            If StrComp(slaAllVal, "Within SLA", vbTextCompare) = 0 Then
                withinSLAAll = withinSLAAll + 1
            End If
        End If
    Next i
    pctAllSLA = IIf(totalAllDone = 0, 0, withinSLAAll / totalAllDone)

    ' Training Score by Project
    
    lastRowTraining = wsTraining.Cells(wsTraining.Rows.Count, "G").End(xlUp).Row
    totalTraining = 0
    goodTraining = 0
    For i = 2 To lastRowTraining
        trainingVal = Trim(wsTraining.Cells(i, "G").Value)
        If trainingVal <> "" Then
            totalTraining = totalTraining + 1
            If StrComp(trainingVal, "Good", vbTextCompare) = 0 Then
                goodTraining = goodTraining + 1
            End If
        End If
    Next i
    pctTrainingGood = IIf(totalTraining = 0, 0, goodTraining / totalTraining)

    ' Update "Week XX" Sheet
    
    With wsTarget
    
        ' Week-over-week comparison D column = prev week, E = current
        ' Rows 22-26 = QIA audit breakdown per category/sourcer
        
        .Range("E4").Value = pctQIAGood
        .Range("E7").Value = pctTrainingGood
        .Range("E18").Value = pctAllSLA
        .Range("E20").Value = pctQIAGood
        .Range("E22").Value = pctQIAGood
        .Range("E23").Value = pctQIAGood
        .Range("E24").Value = pctQIAGood
        .Range("E26").Value = pctQIAGood

        .Range("E4,E7,E18,E20,E22,E23,E24,E26").NumberFormat = "0.00%"

        ' Notes
        
        Dim noteText As String
        noteText = "Note: Status updated " & weekDate & " and beyond"

        .Range("F4").Value = noteText
        .Range("F20").Value = noteText
        .Range("F22").Value = noteText
        .Range("F23").Value = noteText
        .Range("F24").Value = noteText
        .Range("F26").Value = noteText

        .Range("F7").Value = "Note: " & goodTraining & "/" & totalTraining & " passed training score"
        .Range("F18").Value = "Note: " & withinSLAAll & "/" & totalAllDone & " tasks completed"

        ' Conditional Formatting
        
        Dim rngColor As Range
        Set rngColor = .Range("E4,E7,E18,E20,E22,E23,E24,E26")

        Dim cell As Range
        For Each cell In rngColor
            If IsNumeric(cell.Value) Then
                Select Case cell.Value
                    Case Is >= 0.85
                        cell.Interior.Color = RGB(0, 176, 80) ' Green
                    Case 0.76 To 0.84
                        cell.Interior.Color = RGB(255, 255, 0) ' Yellow
                    Case Else
                        cell.Interior.Color = RGB(255, 0, 0) ' Red
                End Select
            Else
                cell.Interior.ColorIndex = xlNone
            End If
        Next cell
    End With
End Sub


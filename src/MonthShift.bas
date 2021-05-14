Attribute VB_Name = "MonthShift"
Option Explicit

' Create new tab for new month and hide same month from previous year
Public Sub Shift_Tabs(ByVal iMonth As String, ByVal iYear As Integer)
    Dim previousYear As Integer
    Dim previousMonth As String
    Dim myMonth As String
    Dim newYearMonth As String
    
    Dim previousMonthSh As Worksheet
    Dim newMonthSh As Worksheet
    
    previousMonth = UCase(MonthName(month(DateValue("01 " & iMonth & " 2012")) - 1, True))
    previousYear = iYear - 1
    
    myMonth = UCase(Left(iMonth, 3))   'Get current month name in 3-letter format
    newYearMonth = Dependencies.Get_Year_Month("yyyy-mm", DateValue("01 " & iMonth & " " & iYear))
    
    On Error Resume Next
    Set previousMonthSh = Worksheets(myMonth & previousYear)
    Set newMonthSh = Worksheets(myMonth)
    ' Check to see if both sheet names already exist, and if so exit sub
    If Not previousMonthSh Is Nothing And Not newMonthSh Is Nothing Then
        previousMonthSh.Visible = xlSheetHidden
        MsgBox "Both sheets already exist!"
        Exit Sub
    Else
        If Not previousMonthSh Is Nothing Then
            previousMonthSh.Visible = xlSheetHidden
        Else
            With Worksheets(myMonth)
                .name = myMonth & previousYear
                .Visible = xlSheetHidden
            End With
        End If

        ' Add new month sheet and copy template
        Worksheets.Add(after:=Worksheets(previousMonth)).name = myMonth
        Worksheets(previousMonth).Range("A1:O5").Copy
        With Worksheets(myMonth)
            With .Range("A1")
                .PasteSpecial xlPasteColumnWidths
                .PasteSpecial xlPasteValues, , False, False
                .PasteSpecial xlPasteFormats, , False, False
                Application.CutCopyMode = False
            End With
            .Range("C2").Value = newYearMonth
            .Range("C2").Select
            ActiveWindow.Zoom = 85
        End With
    End If

    On Error GoTo 0
End Sub
Public Sub Shift_WO_Summary(ByVal iMonth As String, ByVal iYear As Integer)
    Dim lastYearMonth As String
    Dim newYearMonth As String
    
    Dim oldMonthCol As Range
    Dim hideCol As Integer
    
    Dim insertCol As Integer
    
    lastYearMonth = UCase(iMonth) + " " + CStr(iYear - 1) + "'"
    newYearMonth = UCase(iMonth) + " " + CStr(iYear) + "'"

    With Worksheets("WO Summary")
        ' Hide old month column
        Set oldMonthCol = .Rows(2).Find(what:=lastYearMonth, LookIn:=xlValues, lookat:=xlWhole)
        If Not oldMonthCol Is Nothing Then
            hideCol = oldMonthCol.Column
            .Columns(hideCol).Hidden = True
        End If

        ' Insert new column for new month
        insertCol = .Rows(2).Find(what:="CUM (ALL TIME)", LookIn:=xlValues, lookat:=xlWhole).Column
        .Columns(insertCol).Insert shift:=xlToRight
        
        .Cells(2, insertCol).Value = newYearMonth
        .Cells(3, insertCol).Formula = _
            "=COUNTIF(" & UCase(Left(iMonth, 3)) & "!B:B,'" & .name & "'!$B3)"
        .Cells(3, insertCol).AutoFill Destination:=.Range(.Cells(3, insertCol), .Cells(9, insertCol))
        
        .Range(.Cells(11, insertCol - 1), .Cells(15, insertCol - 1)).AutoFill _
            Destination:=.Range(.Cells(11, insertCol - 1), .Cells(15, insertCol))
            
        ' Adjust formulas in cumulative column to include new month
        Dim iCell As Range
        For Each iCell In _
            Union(.Range(.Cells(3, insertCol + 1), .Cells(9, insertCol + 1)), .Cells(13, insertCol + 1))
                iCell.Formula = Replace(iCell.Formula, "$" & Dependencies.Col_Number_To_Letter(insertCol - 1), _
                    "$" & Dependencies.Col_Number_To_Letter(insertCol))
        Next iCell
        
        .Cells(13, insertCol).Select
    End With

End Sub

Public Sub Shift_Site_Summary(ByVal iMonth As String, ByVal iYear As Integer)
    Dim previousYearMonth As String
    Dim newYearMonth As String
    
    Dim prevMonthCol As Integer
    Dim newMonthCol As Integer

    previousYearMonth = UCase(MonthName(month(DateValue("01 " & iMonth & " 2012")) - 1)) + " " + CStr(iYear) + "'"
    newYearMonth = UCase(iMonth) + " " + CStr(iYear) + "'"
    
    With Worksheets("Site Summary")
        prevMonthCol = .Rows(1).Find(what:=previousYearMonth, LookIn:=xlValues, lookat:=xlWhole).Column
        newMonthCol = prevMonthCol + 10
        
        ' Create new month section
        .Range(.Cells(1, prevMonthCol), .Cells(10, prevMonthCol + 9)).Copy
        With .Cells(1, newMonthCol)
'            .PasteSpecial xlPasteAll
            .PasteSpecial xlPasteColumnWidths
            .PasteSpecial xlPasteFormulas, , False, False
            .PasteSpecial xlPasteFormats, , False, False
            Application.CutCopyMode = False
            
            .Value = newYearMonth   ' Update month title
        End With
        
        ' Adjust formulas to point to new month
        Dim iCell As Range
        For Each iCell In .Range(.Cells(3, newMonthCol), .Cells(9, newMonthCol + 6))
            iCell.Formula = Replace(iCell.Formula, _
                Left(previousYearMonth, 3) & "!", _
                Left(newYearMonth, 3) & "!")
        Next iCell
        
        ' Update cumulative section to include new month
        If InStr(.Range("B15").Formula, Dependencies.Col_Number_To_Letter(newMonthCol)) <> 0 Then
            MsgBox "Cumulative formulas already updated!"
            Exit Sub
        Else
            .Range("B15").Formula = .Range("B15").Formula + _
                "+" + Dependencies.Col_Number_To_Letter(newMonthCol) + "3"
            .Range("B15").AutoFill Destination:=.Range("B15:B21"), Type:=xlFillValues
            .Range("B15:B21").AutoFill Destination:=Range("B15:H21"), Type:=xlFillValues
        End If
        
        .Cells(10, newMonthCol + 7).Select
    End With
End Sub

Public Sub Shift_Crew_Summary(ByVal iMonth As String, ByVal iYear As Integer)
    Dim previousYearMonth As String
    Dim newYearMonth As String
    
    Dim rowCount As Integer
    Dim prevMonthCol As Integer
    Dim newMonthCol As Integer
    Dim cumMonthCol As Integer

    previousYearMonth = UCase(MonthName(month(DateValue("01 " & iMonth & " 2012")) - 1)) + " " + CStr(iYear) + "'"
    newYearMonth = UCase(iMonth) + " " + CStr(iYear) + "'"
    
    With Worksheets("Crew Summary")
        Dim findMonth As Range
        Set findMonth = .Rows(1).Find(what:=newYearMonth, LookIn:=xlValues, lookat:=xlWhole)
        If Not findMonth Is Nothing Then
            MsgBox "Month-year already exists!"
            Exit Sub
        Else
            prevMonthCol = .Rows(1).Find(what:=previousYearMonth, LookIn:=xlValues, lookat:=xlWhole).Column
            newMonthCol = prevMonthCol + 10
            cumMonthCol = newMonthCol + 10
            rowCount = .Range("A3").End(xlDown).Row
            
            ' Insert new month section
            .Range(.Cells(1, prevMonthCol), _
                .Cells(rowCount + 1, prevMonthCol + 9)).Copy
            .Cells(1, newMonthCol).Insert shift:=xlToRight
            Application.CutCopyMode = False
            .Cells(1, newMonthCol).Value = newYearMonth    ' Update month title
            
            ' Adjust new month formulas to point to new month tab
            Dim iCell As Range
            For Each iCell In .Range(.Cells(3, newMonthCol), .Cells(rowCount, newMonthCol + 6))
                iCell.Formula = Replace(iCell.Formula, _
                    Left(previousYearMonth, 3) & "!", _
                    Left(newYearMonth, 3) & "!")
            Next iCell
            
            ' Update cumulative section formulas to include new month
            .Cells(3, cumMonthCol).Formula = .Cells(3, cumMonthCol).Formula + _
                "+" + Dependencies.Col_Number_To_Letter(newMonthCol) + "3"
            .Cells(3, cumMonthCol).AutoFill _
                Destination:=.Range(.Cells(3, cumMonthCol), .Cells(rowCount, cumMonthCol)), _
                Type:=xlFillValues
            .Range(.Cells(3, cumMonthCol), .Cells(rowCount, cumMonthCol)).AutoFill _
                Destination:=.Range(.Cells(3, cumMonthCol), .Cells(rowCount, cumMonthCol + 7)), _
                Type:=xlFillValues
                
            .Columns(cumMonthCol + 7).AutoFit
            .Cells(rowCount + 1, newMonthCol + 7).Select
        End If
    End With
End Sub


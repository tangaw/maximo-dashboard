Attribute VB_Name = "Chart"
Option Explicit

' Update completion percentage value
Public Sub Update_Overall_Chart()
    Dim currentComp As Double
    Dim previousComp As Double
    
    With Worksheets("OverallTracker").Range("OverallTracker[COMPLETION]")
        If .Rows.Count = 1 Then
            currentComp = .Value
            previousComp = 0
        Else
            currentComp = .End(xlDown).Value
            previousComp = .End(xlDown).Offset(-1, 0).Value
        End If
    End With
    
    With Worksheets("OverallChart")
        .Range("B2").Value = currentComp
        .Range("B5").Value = previousComp
    End With
End Sub

Public Sub Update_Category_Chart(ByVal category As String)
    Dim headersDict As New Dictionary
    Dim headerKey As Variant
    Dim compPerc As Double
    
    Dim trackerName As String
    Dim chartName As String
    Dim endRow As Integer
    Dim listCol As String
    
    Dim insertCol As Integer
    
    trackerName = category & "Tracker"
    chartName = category & "Chart"
    Set headersDict = Get_Headers(category)
    
    ' Store latest completion rate data in dictionary
    With Worksheets(trackerName)
        If .Range(trackerName).Rows.Count = 1 Then  ' When only one entry
            endRow = 1
        Else
            endRow = .Range(trackerName).End(xlDown).Row - 1    '-1 since listCol starts with data range
        End If
        For Each headerKey In headersDict.Keys
            listCol = trackerName & "[" & headerKey & "]"
            
            If endRow = 1 Then
                headersDict(headerKey) = .Range(listCol)(endRow).Value & "," & 0
            Else
                ' Store both current and previous completion rates with comma separator
                headersDict(headerKey) = .Range(listCol)(endRow).Value & "," & _
                    .Range(listCol)(endRow - 1).Value
            End If
        Next headerKey
    End With
    
    With Worksheets(chartName)
        For Each headerKey In headersDict.Keys
            insertCol = .Rows("1").Find(what:=headerKey, LookIn:=xlValues, lookat:=xlWhole).Column + 1
            .Cells(3, insertCol).Value = Split(headersDict(headerKey), ",")(0)
            .Cells(6, insertCol).Value = Split(headersDict(headerKey), ",")(1)
        Next headerKey
    End With
    
End Sub

' Update CrewPivot to filter to the current month and year
Public Sub Update_Pivot()
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim currentDate As String
    
    Dim sheets() As String
    Dim pivots() As String
    Dim i As Integer
    
    For Each pc In ThisWorkbook.PivotCaches
        pc.Refresh
    Next pc
    
    currentDate = Dependencies.Get_Year_Month("yyyy-mm")
    sheets = Split("DonutPivot,CrewChart", ",")
    pivots = Split("DonutPivot,CrewPivot", ",")
    
    For i = LBound(sheets) To UBound(sheets)
        Set pt = Worksheets(sheets(i)).PivotTables(pivots(i))
        Set pf = pt.PivotFields("year-month")
        
        ' Set "year-month" filter to current year & month if not already
        If pf.CurrentPage <> currentDate Then
            pf.CurrentPage = currentDate
        End If
    Next i

End Sub

Public Sub Update_Crew_Chart_Table()
    ' 1. Update chart table according to number of entries in pivot table
    ' 2. Update source data for chart
    
    Dim people() As Variant
    Dim person As Variant
    
    Dim startCol As Integer
    Dim startRow As Integer
    Dim endRow As Integer
    
    With Worksheets("CrewChart")
        startRow = .Columns("A").Find(what:="Row Labels", LookIn:=xlValues, lookat:=xlWhole).Row
        endRow = .Columns("A").Find(what:="Grand Total", LookIn:=xlValues, lookat:=xlWhole).Row - 1
        
        ' Clear out contents before filling in
        startCol = .Rows(startRow).Find(what:="ASSIGNEE", LookIn:=xlValues, lookat:=xlWhole).Column
        .Cells(startRow, startCol).CurrentRegion.Offset(1, 0).Resize(, 2).ClearContents
        
        people = .Range(.Cells(startRow + 1, 1), .Cells(endRow, 1)).Value
        
        ' Copy array values into chart table
        .Columns(startCol).Find(what:="ASSIGNEE", LookIn:=xlValues, lookat:=xlWhole) _
            .Offset(1, 0).Resize(Dependencies.Arr_Len(people), 1).Value = people
    End With

    ' Get the previous completion rate of all people
    Dim peopleDict As New Dictionary
    Dim personKey As Variant
    Dim listCol As String
    
    For Each person In people
        peopleDict.Add Key:=person, Item:=""
    Next person
    
    With Worksheets("CrewCompTracker")
        If WorksheetFunction.CountA(.Range("CrewCompTracker")) = 1 Then    ' When only one entry
            endRow = -1
        Else
            endRow = .Range("CrewCompTracker").End(xlDown).Row - 2    'Reuse endRow variable to get previous completion
        End If
        
        For Each personKey In peopleDict.Keys
            listCol = "CrewCompTracker[" & personKey & "]"
            If endRow = -1 Then
                peopleDict(personKey) = 0
            Else
                peopleDict(personKey) = .Range(listCol)(endRow).Value
            End If
        Next personKey
    End With
    
    Dim insertRow As Integer
    
    With Worksheets("CrewChart")
        For Each personKey In peopleDict.Keys
            insertRow = .Columns(startCol).Find(what:=personKey, LookIn:=xlValues, lookat:=xlWhole).Row
            .Cells(insertRow, 11).Value = peopleDict(personKey)
        Next personKey
    End With
End Sub

' Update crew chart references based on new data
Public Sub Update_Crew_Chart_Reference()
    Dim i As Long
    Dim startRow As Long
    Dim startCol As Long
    Dim startColLetter As String
    Dim targetCol As String
    Dim endRow As Long
    Dim seriesStr As String
    Dim xValuesRef As String
    
    With Worksheets("CrewChart")
        If .FilterMode Then .showAllData    'Clear all filters to ensure accurate data ranges
        ' Get letter representation of starting column for chart area
        startRow = .Columns("A").Find(what:="Row Labels", LookIn:=xlValues, lookat:=xlWhole).Row
        startCol = .Rows(startRow).Find(what:="ASSIGNEE", LookIn:=xlValues, lookat:=xlWhole).Column
        startColLetter = Col_Number_To_Letter(startCol)
        endRow = .Cells(startRow, startCol).End(xlDown).Row
    End With

    With Worksheets("Dashboard").ChartObjects("CrewChart").Chart
        ' Update reference range for each data series
        For i = 1 To 3
            'Columns are ordered backwards to series sequence: ("PREVIOUS","CURRENT","TOTAL")
            targetCol = Chr(Asc(startColLetter) + (4 - i))
            seriesStr = "=CrewChart!$" & targetCol & "$5:$" & targetCol & "$" & endRow
            xValuesRef = "=CrewChart!$" & startColLetter & "$5:$" & startColLetter & "$" & endRow
            .FullSeriesCollection(i).Values = seriesStr
            .FullSeriesCollection(i).XValues = xValuesRef   'Update x value refs (axis labels) for all series
        Next i
        
        ' Update data label reference for growth percentages
        targetCol = Chr(Asc(startColLetter) + 4)
        seriesStr = "=CrewChart!$" & targetCol & "$5:$" & targetCol & "$" & endRow
        .SeriesCollection(2).DataLabels.format.TextFrame2.TextRange. _
            InsertChartField msoChartFieldRange, seriesStr, 0
            
        ' Update data label reference for Current / Total
        targetCol = Chr(Asc(startColLetter) + 5)
        seriesStr = "=CrewChart!$" & targetCol & "$5:$" & targetCol & "$" & endRow
        .SeriesCollection(1).DataLabels.format.TextFrame2.TextRange. _
            InsertChartField msoChartFieldRange, seriesStr, 0
    End With

End Sub

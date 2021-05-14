Attribute VB_Name = "Tracker"
Option Explicit

' Update either the site or crew tracker
Public Sub Update_Completion_Tracker(ByVal category As String)
    Dim currentMonthYear As String
    Dim summarySheet As String
    Dim trackerName As String
    
    Dim headersDict As New Dictionary
    Dim headerKey As Variant
    
    Dim compRow As Integer
    Dim compCol As Integer
    
    Dim insertRow As Integer
    Dim listColStr As String

    currentMonthYear = Get_Year_Month("mmmm-yy") & "'"
    ' Assign sheet and table names to use based on the 'category' argument
    If category = "CrewComp" Then
        summarySheet = "Crew Summary"
    Else
        summarySheet = category & " Summary"
    End If
    trackerName = category & "Tracker"
    Set headersDict = Get_Headers(category)
    
    ' Get completion data from 'Summary' tab and store in dictionary
    For Each headerKey In headersDict.Keys
        With Worksheets(summarySheet)
            If category = "Site" Then
                compRow = .Columns("A").Find(what:=headerKey, LookIn:=xlValues, lookat:=xlWhole).Row
            Else 'category = "Crew"
                compRow = .Columns("C").Find(what:=headerKey, LookIn:=xlValues, lookat:=xlWhole).Row
            End If
            
            If category = "CrewComp" Then
                compCol = .Rows("1").Find(what:=currentMonthYear, LookIn:=xlValues, lookat:=xlWhole).Column
                headersDict(headerKey) = .Cells(compRow, compCol).Value + .Cells(compRow, compCol + 1).Value
            Else 'category = "Crew" || "Site"
                compCol = .Rows("1").Find(what:=currentMonthYear, LookIn:=xlValues, lookat:=xlWhole).Column + 8
                headersDict(headerKey) = .Cells(compRow, compCol).Value
            End If
        End With
    Next headerKey
    
    ' Determine row to enter new data
    insertRow = Get_Row_to_Insert(trackerName)
    
    ' Paste data from dictionary to appropriate row in table
    For Each headerKey In headersDict.Keys
        listColStr = trackerName & "[" & headerKey & "]"
        Worksheets(trackerName).Range(listColStr)(insertRow).Value = headersDict(headerKey)
    Next headerKey

End Sub

' Update the overall completion rate tracker
Public Sub Update_Overall_Tracker()
    Dim currentMonthYear As String
    Dim currentCompletion As Double
    
    Dim compRow As Integer
    Dim compCol As Integer
    
    Dim insertRow As Integer
    
    currentMonthYear = Get_Year_Month("mmmm-yy") & "'"
    
    With Worksheets("WO Summary")
        compRow = .Columns("B").Find(what:="Comp (%)", LookIn:=xlValues, lookat:=xlWhole).Row
        compCol = .Rows("2").Find(what:=currentMonthYear, LookIn:=xlValues, lookat:=xlWhole).Column
        currentCompletion = .Cells(compRow, compCol).Value
    End With
    
    ' Get row to insert new data
    insertRow = Get_Row_to_Insert("OverallTracker")
    
    ' Paste current completion into tracker
    Worksheets("OverallTracker").Range("OverallTracker[COMPLETION]")(insertRow).Value = currentCompletion
    
End Sub

'------------------------------------------------FUNCTIONS----------------------------------------------------------------------------

' Determine row (by date) to enter new data on tracker
Public Function Get_Row_to_Insert(ByVal trackerName As String) As Integer
    Dim findDate As Range
    Dim insertRow As Integer
    
    With Worksheets(trackerName)
        If WorksheetFunction.CountA(.Range(trackerName & "[DATE]")) = 0 Then    ' Check for empty table
            insertRow = 1
            .Range(trackerName & "[DATE]")(insertRow).Value = Date
        Else
            With .ListObjects(trackerName)
                Set findDate = .ListColumns(1).DataBodyRange _
                    .Find(what:=Date, LookIn:=xlValues, lookat:=xlWhole)
                If findDate Is Nothing Then
                    insertRow = .DataBodyRange.Rows.Count + 1   ' DataBodyRange counts below header
                Else
                    insertRow = findDate.Row - 1
                End If
                .ListColumns(1).DataBodyRange(insertRow).Value = Date
            End With
        End If
    End With
    
    Get_Row_to_Insert = insertRow
End Function
' Get header values (excluding 'Date' column) of requested tab
Public Function Get_Headers(ByVal category As String) As Dictionary
    Dim categoryTab As String
    Dim headers As Range
    Dim dict As New Dictionary
    Dim header As Range
    
    categoryTab = category & "Tracker"
    
    Set headers = Worksheets(categoryTab).ListObjects(categoryTab).HeaderRowRange
    
    ' Store each header as a key in dictionary with an empty item
    For Each header In headers.Offset(0, 1).Resize(1, headers.Count - 1)    ' Exclude 'DATE' header
        dict.Add Key:=header.Value, Item:=""
    Next header

    Set Get_Headers = dict

End Function

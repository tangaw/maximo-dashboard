Attribute VB_Name = "Active"
Option Explicit

Public Sub Refresh_Tabs()
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Dim month As String
    Dim months() As String
    
    months = Get_Months
    
    For Each ws In ThisWorkbook.Worksheets
        month = Left(ws.name, 3)
        If (Is_In_Array(month, months)) Then
            Call Refresh_Page(ws.name)
            Call Reset_Cursor(ws.name)
        End If
    Next ws
    
    ' Log completion rates onto trackers and update charts if it is a weekday
    If Is_Weekday(Date) Then
        Call tracker.Update_Overall_Tracker
        Call tracker.Update_Completion_Tracker("Site")
        Call tracker.Update_Completion_Tracker("Crew")
        Call tracker.Update_Completion_Tracker("CrewComp")
        
        Call Chart.Update_Overall_Chart
        Call Chart.Update_Category_Chart("Site")
        Call Chart.Update_Pivot
        Call Chart.Update_Crew_Chart_Table
        Call Chart.Update_Crew_Chart_Reference
    End If
    
    Worksheets("Dashboard").Activate
    
    Application.ScreenUpdating = True
End Sub

Public Sub Refresh_Page(ByVal wsName As String)
    Call Update_Page_Content(wsName)
    Call Sort(wsName)
    Call Filter(wsName)
End Sub
Public Sub Update_Page_Content(ByVal wsName As String)
Attribute Update_Page_Content.VB_ProcData.VB_Invoke_Func = " \n14"
    Call Clear_Tab_Data(wsName)
    
    With Worksheets(wsName)
        sheets("ALL").Range("Table_Maximo_Report_Import[#All]").AdvancedFilter _
            Action:=xlFilterCopy, CriteriaRange:=.Range("A1").CurrentRegion, _
            CopyToRange:=.Range("A5:O5"), Unique:=False
    End With
    
    Application.CutCopyMode = False
End Sub
Public Sub Sort(ByVal wsName As String)
    Dim ws As Worksheet
    Set ws = Worksheets(wsName)
    
    Dim findInprg As Range
    Dim findWappr As Range
    Dim findNc As Range
    
    With ws
        Set findInprg = .Range("B:B").Find(what:="INPRG", LookIn:=xlValues, lookat:=xlWhole)
        Set findWappr = .Range("B:B").Find(what:="WAPPR", LookIn:=xlValues, lookat:=xlWhole)
        Set findNc = .Range("B:B").Find(what:="NC", LookIn:=xlValues, lookat:=xlWhole)
    
        If Not .AutoFilterMode Then
            .Range("A5").AutoFilter
        End If
        
        With .AutoFilter.Sort
            With .SortFields
                .Clear
                If Not findInprg Is Nothing Or Not findWappr Is Nothing Then
                    .Add(ws.Range("E6"), xlSortOnCellColor, xlAscending, , xlSortNormal) _
                        .SortOnValue.Color = RGB(255, 255, 102)
                End If
                If Not findNc Is Nothing Then
                    .Add(ws.Range("E6"), xlSortOnCellColor, xlAscending, , xlSortNormal) _
                        .SortOnValue.Color = RGB(255, 153, 102)
                End If
                .Add2 Key:=Range("E6"), SortOn:=xlSortOnValues, Order:=xlAscending, _
                    DataOption:=xlSortNormal    ' Lowest sort level: alphabetically
            End With
            .header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
End Sub
Public Sub Filter(ByVal wsName As String)
Attribute Filter.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim findInprg As Range
    Dim findWappr As Range
    Dim findNc As Range
    
    With Worksheets(wsName)
        Set findInprg = .Range("B:B").Find(what:="INPRG", LookIn:=xlValues, lookat:=xlWhole)
        Set findWappr = .Range("B:B").Find(what:="WAPPR", LookIn:=xlValues, lookat:=xlWhole)
        Set findNc = .Range("B:B").Find(what:="NC", LookIn:=xlValues, lookat:=xlWhole)
        
        If findInprg Is Nothing And findWappr Is Nothing And findNc Is Nothing Then
            Exit Sub
        Else
            .Range("B6").CurrentRegion.AutoFilter field:=2, Criteria1:=Array("INPRG", "WAPPR", "NC") _
                , Operator:=xlFilterValues
        End If
    End With
End Sub
' Reset cursor on all month tabs to Cell "C2"
Public Sub Reset_Cursor(ByVal wsName As String)
    With Worksheets(wsName)
        If .Visible = xlSheetHidden Or .Visible = xlSheetVeryHidden Then
            .Visible = xlSheetVisible
            .Select
            .Range("C2").Activate
            .Visible = xlSheetHidden
        Else
            .Select
            .Range("C2").Activate
        End If
    End With
End Sub

Public Sub Clear_All_Month_Tabs()
Attribute Clear_All_Month_Tabs.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim ws As Worksheet
    Dim months() As String
    
    months = Get_Months
    
    For Each ws In ThisWorkbook.Worksheets
        If (Is_In_Array(Left(ws.name, 3), months)) Then
            Call Clear_Tab_Data(ws.name)
        End If
    Next ws

End Sub
Public Sub Clear_Tab_Data(ByVal wsName As String)
Attribute Clear_Tab_Data.VB_ProcData.VB_Invoke_Func = " \n14"
    With Worksheets(wsName)
        On Error Resume Next
        .showAllData
        On Error GoTo 0
        .Range("A5").CurrentRegion.Offset(1, 0).EntireRow.Delete    'Exclude headers
    End With
    
End Sub

' Toggle visibility of all design tabs, including stylesheet, trackers, charts, and pivots
Public Sub Toggle_Design_Tabs_Visibility()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "Stylesheet" Or _
            InStr(ws.name, "Tracker") <> 0 Or _
            InStr(ws.name, "Chart") <> 0 Or _
            InStr(ws.name, "Pivot") <> 0 Then
                If ws.Visible = xlSheetVeryHidden Or ws.Visible = xlSheetHidden Then
                    ws.Visible = xlSheetVisible
                Else
                    ws.Visible = xlSheetVeryHidden
                End If
        End If
    Next ws
End Sub

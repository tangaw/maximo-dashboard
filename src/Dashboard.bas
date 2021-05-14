Attribute VB_Name = "Dashboard"
Option Explicit

' Dynamically display assignee information based on site selection
Public Sub Change_Assignee_Site_Page(ByVal selection As String)
    Dim pages() As String
    Dim i As Long
    
    Dim pageName As String
    Dim buttonName As String
    
    Dim pageShape As Shape
    Dim buttonShape As Shape
    
    pages = Split("All,Arques,Bowers/Scott", ",")
    With Worksheets("Dashboard")
        For i = LBound(pages) To UBound(pages)
            pageName = pages(i) + " Page"
            buttonName = pages(i) + " Button"
            
            Set pageShape = .Shapes(pageName)
            Set buttonShape = .Shapes(buttonName)
            
            ' Change selection button color to white and bring to front
            If pages(i) = selection Then
                buttonShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
                pageShape.ZOrder msoBringToFront
            Else
                buttonShape.Fill.ForeColor.RGB = RGB(248, 247, 247)
                pageShape.ZOrder msoSendToBack
            End If
        Next i
    End With
    
    Call Filter_Crew_Chart(selection)
    
End Sub

Public Sub Filter_Crew_Chart(ByVal selection As String)
    Dim sites() As String
    Dim startRow As Long
    Dim startCol As Long
    
    ' Filter data
    With Worksheets("CrewChart")
        If selection = "All" Then
            If .FilterMode Then .showAllData
        Else
            sites = Split(selection, "/")
            startRow = .Columns("A").Find(what:="Row Labels", LookIn:=xlValues, lookat:=xlWhole).Row
            startCol = .Rows(startRow).Find(what:="ASSIGNEE", LookIn:=xlValues, lookat:=xlWhole).Column
            .Range(.Cells(startRow, startCol).Address).AutoFilter field:=7, Criteria1:=sites, Operator:=xlFilterValues
        End If
    End With
    
    ' Adjust axis tick mark and data label font sizes
    With Worksheets("Dashboard").ChartObjects("CrewChart").Chart
        If selection = "All" Then
            .FullSeriesCollection(1).DataLabels.format.TextFrame2.TextRange.Font.Size = 9
            .FullSeriesCollection(2).DataLabels.format.TextFrame2.TextRange.Font.Size = 9
            .Axes(xlCategory).TickLabels.Font.Size = 9
        Else
            .FullSeriesCollection(1).DataLabels.format.TextFrame2.TextRange.Font.Size = 11
            .FullSeriesCollection(2).DataLabels.format.TextFrame2.TextRange.Font.Size = 10
            .Axes(xlCategory).TickLabels.Font.Size = 10
        End If
    End With
End Sub

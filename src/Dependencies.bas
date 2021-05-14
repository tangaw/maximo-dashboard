Attribute VB_Name = "Dependencies"
Option Explicit

' Return current month and year in "mmmm-yy" format (e.g. DECEMBER 21)
Public Function Get_Year_Month(ByVal format As String, Optional iDate As Date) As String
    Dim returnMonth As String
    Dim returnYear As Integer
    
    ' Default to current date if no argument is supplied
    If iDate = 0 Then iDate = Date
    
    'e.g. DECEMBER 21
    If format = "mmmm-yy" Then
        returnMonth = UCase(MonthName(month(iDate)))
        returnYear = Right(Year(iDate), 2)
    
        Get_Year_Month = returnMonth & " " & returnYear
    'e.g. 2021-04
    ElseIf format = "yyyy-mm" Then
        returnMonth = month(iDate)
        If Len(returnMonth) = 1 Then
            returnMonth = "0" & returnMonth
        End If
        returnYear = Year(iDate)
        
        Get_Year_Month = returnYear & "-" & returnMonth
    'e.g. APR21
    ElseIf format = "mmm-yy" Then
        returnYear = Right(Year(iDate), 2)
        returnMonth = UCase(MonthName(month(iDate), True))
        
        Get_Year_Month = returnMonth & returnYear
    End If

End Function

Public Function Is_Weekday(ByVal inputDate As Date) As Boolean
    Select Case Weekday(inputDate)
        Case 2 To 6
            Is_Weekday = True
        Case Else
            Is_Weekday = False
    End Select
End Function
Public Function Get_Months() As String()
    Dim months As String
    
    months = "JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC"
    Get_Months = Split(months, ",")
End Function
Public Function Is_In_Array(stringToBeFound As String, arr As Variant) As Boolean
    Dim i As Byte
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            Is_In_Array = True
            Exit Function
        End If
    Next i
    Is_In_Array = False
End Function

' Get length of an array
Public Function Arr_Len(ByVal arr As Variant) As Integer
    Arr_Len = UBound(arr) - LBound(arr) + 1
End Function

' Get letter representation of column number
Public Function Col_Number_To_Letter(ByVal lngCol As Long) As String
    Dim vArr As Variant
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Number_To_Letter = vArr(0)
End Function

Public Function VBA_Long_To_RGB(lColor As Long) As String
    Dim iRed As Byte
    Dim iGreen As Byte
    Dim iBlue As Byte
    
    'Convert Decimal Color Code to RGB
    iRed = (lColor Mod 256)
    iGreen = (lColor \ 256) Mod 256
    iBlue = (lColor \ 65536) Mod 256
    
    'Return RGB Code
    VBA_Long_To_RGB = "RGB(" & iRed & "," & iGreen & "," & iBlue & ")"
End Function

Public Sub Add_Conditional_Format_All()
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Dim month As String
    Dim months() As String

    months = Get_Months

    For Each ws In Worksheets
        month = Left(ws.name, 3)
        If Is_In_Array(month, months) Then
            Call Conditional_Format(ws.name)
        End If
    Next ws
    
    Application.ScreenUpdating = True
End Sub
Public Sub Conditional_Format(ByVal wsName As String)
    With Worksheets(wsName)
        .Cells.FormatConditions.Delete
        With .Columns("E:E")
        
            .FormatConditions.Add Type:=xlExpression, Formula1:="=($B1=""NC"")"
            .FormatConditions(selection.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 6724095
                .TintAndShade = 0
            End With
            .FormatConditions(1).StopIfTrue = False
        
            .FormatConditions.Add Type:=xlExpression, Formula1:="=OR($B1=""INPRG"", $B1=""WAPPR"")"
            .FormatConditions(selection.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 6750207
                .TintAndShade = 0
            End With
            .FormatConditions(1).StopIfTrue = False
        End With

    End With
End Sub

Attribute VB_Name = "Shortcuts"
Option Explicit

Public Sub CreateShortcuts()
    Application.OnKey "+^{u}", "Active.Refresh_Tabs"
    Application.OnKey "^{u}", "Active.Update_Page_Content"
    Application.OnKey "^{y}", "Active.Sort"
    Application.OnKey "^{i}", "Active.Filter"
    Application.OnKey "+^{p}", "Active.Clear_All_Month_Tabs"
    Application.OnKey "^{p}", "Active.Clear_Tab_Data"
End Sub
Public Sub DeleteShortcuts()
    Application.OnKey "+^{u}"
    Application.OnKey "^{u}"
    Application.OnKey "^{y}"
    Application.OnKey "^{i}"
    Application.OnKey "+^{p}"
    Application.OnKey "^{p}"
End Sub

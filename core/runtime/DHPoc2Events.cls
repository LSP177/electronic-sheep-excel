VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "DHPoc2Events"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents App As Excel.Application
Private Sub App_SheetActivate(ByVal Sh As Object)
    DHPoc2_Invalidate "sheet activated"
End Sub
Private Sub App_SheetDeactivate(ByVal Sh As Object)
    DHPoc2_Invalidate "sheet deactivated"
End Sub
Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    DHPoc2_Invalidate "workbook activated"
End Sub
Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook)
    DHPoc2_Invalidate "workbook deactivated"
End Sub
Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    DHPoc2_SelectionEvent Sh, Target
End Sub
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    DHPoc2_DataChanged Sh
End Sub
Private Sub App_SheetCalculate(ByVal Sh As Object)
    DHPoc2_DataChanged Sh
End Sub
Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    DHPoc2_Invalidate "workbook closing"
End Sub

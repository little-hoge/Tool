VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents app As Application
Attribute app.VB_VarHelpID = -1

Private Sub app_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)

    Application.ScreenUpdating = True
End Sub

Private Sub app_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)

    Call クローズハイライトオフ
End Sub

Private Sub Workbook_Open()

    Set app = Application
End Sub

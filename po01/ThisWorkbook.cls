﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
Call ExportModules
End Sub



Private Sub Workbook_Open()
'Git検証～
Range("B1") = "poDetl01"
Range("B2") = "poDetl02"


End Sub

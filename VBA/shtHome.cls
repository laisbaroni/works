VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
'barra de scroll horizontal
ActiveWindow.DisplayHorizontalScrollBar = False
'barra de scroll vertical
ActiveWindow.DisplayVerticalScrollBar = False

'cor padrao
    ActiveSheet.Shapes.Range(Array("MenuTopo")).Fill.ForeColor.RGB = RGB(217, 217, 217)
    ActiveSheet.Shapes.Range(Array("MenuTopo")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(128, 128, 128)

End Sub

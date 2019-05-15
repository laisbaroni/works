VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
'barra de scroll horizontal
ActiveWindow.DisplayHorizontalScrollBar = True
'barra de scroll vertical
ActiveWindow.DisplayVerticalScrollBar = True
End Sub


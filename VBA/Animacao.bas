Attribute VB_Name = "Animacao"
Sub animacaoHome()
If shtHome.Shapes("home2").Visible = True Then
shtHome.Shapes("home2").Visible = False
Else
shtHome.Shapes("home2").Visible = True
End If
End Sub

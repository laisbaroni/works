Attribute VB_Name = "mdTela"
' OCULTAR OU MOSTRAR COMANDOS DO EXCEL
Public Sub TelaMaxima()
'menu superior do excel
Application.ExecuteExcel4Macro "show.toolbar(""ribbon"",false)"
'barra de formula
Application.DisplayFormulaBar = False
'barra de status
Application.DisplayStatusBar = False
'títulos ou cabeçalho
ActiveWindow.DisplayHeadings = False
'linhas de grade
ActiveWindow.DisplayGridlines = False
'planilha tabs
ActiveWindow.DisplayWorkbookTabs = False
'barra e scroll horizontal
ActiveWindow.DisplayHorizontalScrollBar = False
'barra e scroll vertical
ActiveWindow.DisplayVerticalScrollBar = False
End Sub

Public Sub TelaNormal()
'menu superior do excel
Application.ExecuteExcel4Macro "show.toolbar(""ribbon"",true)"
'barra de formula
Application.DisplayFormulaBar = True
'barra de status
Application.DisplayStatusBar = True
'títulos ou cabeçalho
ActiveWindow.DisplayHeadings = True
'linhas de grade
ActiveWindow.DisplayGridlines = True
'planilha tabs
ActiveWindow.DisplayWorkbookTabs = True
'barra e scroll horizontal
ActiveWindow.DisplayHorizontalScrollBar = True
'barra e scroll vertical
ActiveWindow.DisplayVerticalScrollBar = True
End Sub

'ATIVAR E DESATIVAR TELA CHEIA
Public Sub AtivarDesativar()
Application.ScreenUpdating = False
If shtCadastro.Range("OpTela") = "" Then
shtCadastro.Range("OpTela") = 2
Call TelaMaxima
shtCadastro.Range("OpTela") = 1

ElseIf shtCadastro.Range("OpTela") = 2 Then
Call TelaMaxima
shtCadastro.Range("OpTela") = 1

Else
Call TelaNormal
shtCadastro.Range("OpTela") = 2
End If
Application.ScreenUpdating = True
End Sub

'ENTRAR
Public Sub Entrar()
Application.ScreenUpdating = False
linha = 12
Do Until shtConfig.Cells(linha, "M") = ""

If shtLogin.Range("codigo") = shtConfig.Cells(linha, "M") Then
MsgBox "Olá " & shtConfig.Cells(linha, "L"), vbInformation, "ACESSO LIBERADO"
Range("codigo") = ""

ActiveWorkbook.Unprotect Password:="1702"
shtCadastro.Visible = xlSheetVisible
shtConfig.Visible = xlSheetVisible
shtDados.Visible = xlSheetVisible
shtHome.Visible = xlSheetVisible
shtLogin.Visible = xlSheetVisible
shtHome.Select
Call TelaMaxima
Exit Sub
End If
linha = linha + 1
Loop

MsgBox "Código incorreto", vbCritical, "ACESSO NEGADO"
Range("codigo") = ""
Application.ScreenUpdating = True
End Sub


'SAIR DO SISTEMA
Public Sub SairSistema()
Application.ScreenUpdating = False
ActiveWorkbook.Save
Application.Quit
Application.ScreenUpdating = True
End Sub


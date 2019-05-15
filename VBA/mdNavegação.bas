Attribute VB_Name = "mdNavegação"
Sub EfeitoMenuTopo()
Attribute EfeitoMenuTopo.VB_ProcData.VB_Invoke_Func = " \n14"
Dim sMenuTopo As String

sMenuTopo = ActiveSheet.Shapes(Application.Caller).Name
    
    'cor padrao
    ActiveSheet.Shapes.Range(Array("MenuTopo")).Fill.ForeColor.RGB = RGB(217, 217, 217)
    ActiveSheet.Shapes.Range(Array("MenuTopo")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(128, 128, 128)

    'cor selecionado
    ActiveSheet.Shapes.Range(Array(sMenuTopo)).Fill.ForeColor.RGB = RGB(51, 141, 143)
    ActiveSheet.Shapes.Range(Array(sMenuTopo)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
      
End Sub

Public Sub EfeitoMenuCadastro()
Dim sMenuCadastro As String

sMenuCadastro = ActiveSheet.Shapes(Application.Caller).Name
    
    'cor padrao
    ActiveSheet.Shapes.Range(Array("MenuCadastro")).Fill.ForeColor.RGB = RGB(217, 217, 217)
    ActiveSheet.Shapes.Range(Array("MenuCadastro")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(128, 128, 128)

    'cor selecionado
    ActiveSheet.Shapes.Range(Array(sMenuCadastro)).Fill.ForeColor.RGB = RGB(51, 141, 143)
    ActiveSheet.Shapes.Range(Array(sMenuCadastro)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
      
End Sub

'Menu de informações pessoais
Public Sub MenuIP()
Application.ScreenUpdating = False
shtCadastro.Unprotect Password:=""
Call EfeitoMenuCadastro
shtCadastro.Range("Guia") = 1
shtCadastro.Rows("9:28").EntireRow.Hidden = False
shtCadastro.Rows("29:31").EntireRow.Hidden = True
shtCadastro.Rows("32:45").EntireRow.Hidden = True
shtCadastro.Shapes("Rect01").Visible = msoTrue 'isso pq o retangulo estava aparecendo nas duas abas
shtCadastro.Shapes("Rect02").Visible = msoFalse 'isso pq o retangulo estava aparecendo nas duas abas
shtCadastro.Protect Password:=""
shtCadastro.Range("Cad_1").Select
Application.ScreenUpdating = True
End Sub

'Menu de informções do curso
Public Sub MenuIC()
Application.ScreenUpdating = False
shtCadastro.Unprotect Password:=""
Call EfeitoMenuCadastro
shtCadastro.Range("Guia") = 2
shtCadastro.Rows("9:28").EntireRow.Hidden = True
shtCadastro.Rows("29:31").EntireRow.Hidden = True
shtCadastro.Rows("32:45").EntireRow.Hidden = False
shtCadastro.Shapes("Rect01").Visible = msoFalse 'isso pq o retangulo estava aparecendo nas duas abas
shtCadastro.Shapes("Rect02").Visible = msoTrue 'isso pq o retangulo estava aparecendo nas duas abas
shtCadastro.Protect Password:=""
shtCadastro.Range("Cad_16").Select
Application.ScreenUpdating = True
End Sub

'Efeito menu Ação
Sub EfeitoMenuAcao()
Dim sMenuAcao As String

sMenuAcao = ActiveSheet.Shapes(Application.Caller).Name
    
    'cor padrao
    ActiveSheet.Shapes.Range(Array("MenuAcao")).Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("MenuAcao")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(128, 128, 128)

    'cor selecionado
    ActiveSheet.Shapes.Range(Array(sMenuAcao)).Fill.ForeColor.RGB = RGB(51, 141, 143)
    ActiveSheet.Shapes.Range(Array(sMenuAcao)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
      
End Sub

Private Sub OpenFolder(strDirectory As String)
'DESCRIPTION: Open folder if not already open. Otherwise, activate the already opened window
'DEVELOPER: Ryan Wells (wellsr.com)
'INPUT: Pass the procedure a string representing the directory you want to open
Dim pID As Variant
Dim sh As Variant
On Error GoTo 102:
Set sh = CreateObject("shell.application")
For Each w In sh.Windows
    If w.Name = "Windows Explorer" Or w.Name = "File Explorer" Then
        If w.document.folder.self.Path = strDirectory Then
            'if already open, bring it front
            w.Visible = False
            w.Visible = True
            Exit Sub
        End If
    End If
Next
'if you get here, the folder isn't open so open it
pID = Shell("explorer.exe " & strDirectory, vbNormalFocus)
102:
End Sub

'Continuação do anterior para abrir pasta de documentos
Sub AbrirPastaDocs()
'Demo - opens the folder location saved to the variable strPath
Dim strPath As String
strPath = ThisWorkbook.Path & "\Documentos\"
Call OpenFolder(strPath)
End Sub


'COMANDOS PARA MENUS DE NAVEGAÇÃO

Public Sub mCadastro()
Application.ScreenUpdating = False
shtCadastro.Select
Range("Cad_1").Select
Call EfeitoMenuTopo
Application.ScreenUpdating = True
End Sub

Public Sub mConfig()
Application.ScreenUpdating = False
shtConfig.Select
Range("A1").Select
Call EfeitoMenuTopo
Application.ScreenUpdating = True
End Sub

Public Sub mDados()
Application.ScreenUpdating = False
shtDados.Select
Range("A10").Select
Call EfeitoMenuTopo
Application.ScreenUpdating = True
End Sub

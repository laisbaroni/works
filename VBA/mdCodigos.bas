Attribute VB_Name = "mdCodigos"
Public fotopasta As String
Public fotobd As String
Public linha As Integer
Public linhalistbox As Integer
Public conte As Integer
Public resposta As String
Public UltimaLinha As Integer
Public cliquepesquisa As String
Public valorpesquisado As String
Public valorcelula As String
Public coluna As Integer

'Buscar foto do cliente na pasta do pc
Public Sub BuscarFoto()
On Error GoTo erro
fotopasta = Application.GetOpenFilename(FileFilter:="Image Files(*.jpg),*.jpg")
shtCadastro.FotoCliente.Picture = LoadPicture(fotopasta)
erro:
End Sub

'Cancelar Foto
Public Sub CancelarFoto()
On Error GoTo erro
fotopasta = ThisWorkbook.Path & "\Fotos\" & "padrao.jpg"
shtCadastro.FotoCliente.Picture = LoadPicture(fotopasta)
fotopasta = ""
erro:
End Sub

'Cadastro de Informações
Public Sub Cadastrar()
shtDados.Unprotect Password:=""
Call EfeitoMenuAcao
linha = 10

If shtCadastro.Range("Guia") = 1 Then

'1a restrição - código ou nome em branco, não cadastrar
If shtCadastro.Range("Cad_0") = "" Or shtCadastro.Range("Cad_1") = "" Then
MsgBox "Verifique o preenchimento do código e do nome do aluno", vbInformation, "Cadastro Negado"
Exit Sub
End If

'2a restrição - se o código já estiver cadastrado, não cadastrar
Do Until shtDados.Cells(linha, "A") = ""
If shtDados.Cells(linha, "A") = shtCadastro.Range("Cad_0") Then
MsgBox "Este aluno já está cadastrado", vbInformation, "Cadastro Negado"
Exit Sub
End If
linha = linha + 1
Loop

'3a restrição - se a matrícula já estiver cadastrada, não cadastrar
Do Until shtDados.Cells(linha, "A") = ""
If shtDados.Cells(linha, "Z") = shtCadastro.Range("Cad_25") Then
MsgBox "Esta matrícula já existe", vbInformation, "Cadastro Negado"
Exit Sub
End If
linha = linha + 1
Loop

'não havendo mais restrições então, cadastrar
Do Until shtDados.Cells(linha, "A") = ""
linha = linha + 1
Loop

For a = 0 To 16
shtDados.Cells(linha, "A").Offset(0, a) = shtCadastro.Range("Cad_" & a).Value
Next
shtDados.Cells(linha, "A").Offset(0, 26) = fotopasta
fotopasta = ""
MsgBox "Dados cadastrados com sucesso", vbInformation, "Cadastro de Informações"
Exit Sub

'cadastrar guia de informações de curso
Else
Do Until shtDados.Cells(linha, "A") = ""
If shtDados.Cells(linha, "A") = shtCadastro.Range("Cad_0") Then

'verificar se já existe dados cadastrados
For a = 17 To 25
If shtDados.Cells(linha, "A").Offset(0, a) <> "" Then
MsgBox "Já cadastrado", vbInformation, "Cadastro Existente"
Exit Sub
End If
Next

'se não existir, cadastrar
For a = 17 To 25
shtDados.Cells(linha, "A").Offset(0, a) = shtCadastro.Range("Cad_" & a).Value
Next
MsgBox "Dados cadastrados com sucesso", vbInformation, "Cadastro de Informações"
ActiveWorkbook.Save
Exit Sub
End If
linha = linha + 1
Loop
MsgBox "Este aluno ainda não foi cadastrado", vbInformation, "Cadastro Negado"

End If
shtDados.Protect Password:=""
End Sub

'Buscar dados para o formulário pela matrícula
Public Sub BuscarMatricula()
linha = 10
Do Until shtDados.Cells(linha, "Z") = ""
If shtDados.Cells(linha, "Z") = shtCadastro.Range("Cad_25") Then

'antes da Matrícula
For a = -25 To -1
shtCadastro.Range("Cad_" & a + 25) = shtDados.Cells(linha, "Z").Offset(0, a).Value
Next
'Depois da Matrícula não tem mais nada
fotobd = shtDados.Cells(linha, "Z").Offset(0, 1).Value
shtCadastro.FotoCliente.Picture = LoadPicture(fotobd)

End If
linha = linha + 1
Loop
End Sub

'Buscar dados para o formulário pelo código
Public Sub BuscarCodigo()
linha = 10
Do Until shtDados.Cells(linha, "A") = ""
If shtDados.Cells(linha, "A") = shtCadastro.Range("Cad_0") Then

For a = 1 To 25
shtCadastro.Range("Cad_" & a) = shtDados.Cells(linha, "A").Offset(0, a).Value
Next
fotobd = shtDados.Cells(linha, "A").Offset(0, 26).Value
shtCadastro.FotoCliente.Picture = LoadPicture(fotobd)

End If
linha = linha + 1
Loop
End Sub

'Buscar dados para o formulário pelo CPF
Public Sub BuscarCPF()
linha = 10
Do Until shtDados.Cells(linha, "E") = ""
If shtDados.Cells(linha, "E") = shtCadastro.Range("Cad_4") Then

'antes do CPF
For a = -4 To -1
shtCadastro.Range("Cad_" & a + 4) = shtDados.Cells(linha, "E").Offset(0, a).Value
Next
'Depois do CPF
For a = 1 To 21
shtCadastro.Range("Cad_" & a + 4) = shtDados.Cells(linha, "E").Offset(0, a).Value
Next

fotobd = shtDados.Cells(linha, "E").Offset(0, 22).Value
shtCadastro.FotoCliente.Picture = LoadPicture(fotobd)

End If
linha = linha + 1
Loop
End Sub

'LIMPAR FORMULÁRIO
Public Sub Cancelar()
Call EfeitoMenuAcao
For a = 0 To 25
shtCadastro.Range("Cad_" & a) = ""
Next
Call CancelarFoto
End Sub

'NOVO ALUNO
Public Sub NovoCliente()
Call EfeitoMenuAcao
linha = 10
conte = 1
Do Until shtDados.Cells(linha, "A") = ""
linha = linha + 1
conte = conte + 1
Loop
Call Cancelar
Call MenuIP
shtCadastro.Range("Cad_0") = conte
End Sub

'ALTERAR CADASTRO
Public Sub Alterar()
Call EfeitoMenuAcao
shtDados.Unprotect Password:=""
linha = 10
Do Until shtDados.Cells(linha, "A") = ""
If shtDados.Cells(linha, "A") = shtCadastro.Range("Cad_0") Then

'Informações pessoais
If shtCadastro.Range("Guia") = 1 Then
For a = 1 To 16
shtDados.Cells(linha, "A").Offset(0, a) = shtCadastro.Range("Cad_" & a).Value
Next
If fotopasta = "" Then
shtDados.Cells(linha, "A").Offset(0, 26) = fotobd
fotobd = ""
Else
shtDados.Cells(linha, "A").Offset(0, 26) = fotopasta
fotopasta = ""
End If
MsgBox "Dados alterados com sucesso", vbInformation, "Alteração do Cadastro"
ActiveWorkbook.Save
Exit Sub

'Informações de curso
Else
For a = 17 To 25
shtDados.Cells(linha, "A").Offset(0, a) = shtCadastro.Range("Cad_" & a).Value
Next
MsgBox "Dados alterados com sucesso", vbInformation, "Alteração do Cadastro"
ActiveWorkbook.Save
Exit Sub

End If
End If
linha = linha + 1
Loop
End Sub

'EXCLUIR
Public Sub Excluir()
Call EfeitoMenuAcao
linha = 10
Do Until shtDados.Cells(linha, "A") = ""
If shtDados.Cells(linha, "A") = shtCadastro.Range("Cad_0") Then
resposta = MsgBox("Tem certeza que deseja excluir este cadastro?", vbYesNo, "Excluir")
If resposta = vbYes Then
shtDados.Cells(linha, "A").EntireRow.Delete
MsgBox ("Cadastro excluido com sucesso"), cbInformation, "Excluido"
ActiveWorkbook.Save
Call Cancelar
Call OrganizarCodigoBD
Else: Exit Sub
End If
End If
linha = linha + 1
Loop
End Sub

'Organizar código no BD
Public Sub OrganizarCodigoBD()
linha = 10
conte = 0
Do Until shtDados.Cells(linha, "B") = ""
conte = conte + 1
shtDados.Cells(linha, "A") = conte
linha = linha + 1
Loop
End Sub

'BOTÕES ANTERIOR e PRÓXIMO
'Anterior
Public Sub Anterior()
Call EfeitoMenuAcao
shtCadastro.Range("Cad_0") = shtCadastro.Range("Cad_0") - 1
Call BuscarCodigo
If shtCadastro.Range("Cad_0") <= 0 Then
shtCadastro.Range("Cad_0") = 1
Call BuscarCodigo
MsgBox "Este é o primeiro cadastro", vbInformation, "Impossível Continuar"
Exit Sub
End If
End Sub

'Próximo
Public Sub Proximo()
Call EfeitoMenuAcao
shtCadastro.Range("Cad_0") = shtCadastro.Range("Cad_0") + 1
Call BuscarCodigo
UltimaLinha = shtDados.Cells(Rows.Count, "A").End(xlUp).Row - 9
If shtCadastro.Range("Cad_0") >= UltimaLinha Then
shtCadastro.Range("Cad_0") = UltimaLinha
Call BuscarCodigo
MsgBox "Este é o último cadastro", vbInformation, "Impossível Continuar"
Exit Sub
End If
End Sub

'ABRIR FORMULARIO DE PESQUISA
Public Sub AbrirPesquisa()
frmPesquisa.Show
End Sub

'PREENCHER LISTBOX
Public Sub PreencherListBox()
frmPesquisa.ListBox1.Clear
linha = 10
linhalistbox = 0
conte = 0

Do Until shtDados.Cells(linha, "A") = ""

    With frmPesquisa.ListBox1
    
        .AddItem
        .List(linhalistbox, 0) = shtDados.Cells(linha, "B").Value 'Nome
        .List(linhalistbox, 1) = Format(shtDados.Cells(linha, "E").Value, "000"".""000"".""000""-""00") 'CPF
        .List(linhalistbox, 2) = Format(shtDados.Cells(linha, "N").Value, "(00)"" ""00000""-""0000") 'Celular
    linhalistbox = linhalistbox + 1
    End With

conte = conte + 1
linha = linha + 1
Loop

frmPesquisa.lblTotalRegistros = "Total de registros localizados: " & conte

End Sub

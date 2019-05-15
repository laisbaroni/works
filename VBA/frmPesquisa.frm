VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPesquisa 
   Caption         =   "Pesquisa"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5520
   OleObjectBlob   =   "frmPesquisa.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CARREGAR OS DADOS DO BD PARA O FORMULARIO QUANDO CLICAR NO NOME DO CLIENTE NA LIST BOX
Private Sub ListBox1_Click()

cliquepesquisa = Me.ListBox1
linha = 10
Do Until shtDados.Cells(linha, "B") = cliquepesquisa
linha = linha + 1
Loop
'On Error Resume Next
For a = 0 To 25
shtCadastro.Range("Cad_" & a) = shtDados.Cells(linha, "A").Offset(0, a).Value
Next
fotobd = shtDados.Cells(linha, "A").Offset(0, 27).Value
shtCadastro.FotoCliente.Picture = LoadPicture(fotobd)

End Sub

'FAZER UM FILTRO EM TEMPO REAL NA CAIXA DE TEXTO DA PESQUISA
Private Sub TextBox1_Change()

Me.ListBox1.Clear
valorpesquisado = Me.TextBox1

If Me.OptNome = True Then
coluna = 2

ElseIf Me.OptCPF = True Then
coluna = 5

Else
coluna = 14
End If

linha = 10
linhalistbox = 0
conte = 0

With shtDados
    While .Cells(linha, "A").Value <> Empty
    valorcelula = .Cells(linha, coluna).Value
    
    If InStr(1, UCase(valorcelula), Trim$(UCase(valorpesquisado))) > 0 Then
        
        With Me.ListBox1
            
            .AddItem
            .List(linhalistbox, 0) = shtDados.Cells(linha, "B").Value 'Nome
            .List(linhalistbox, 1) = Format(shtDados.Cells(linha, "E").Value, "000"".""000"".""000""-""00") 'CPF
            .List(linhalistbox, 2) = Format(shtDados.Cells(linha, "N").Value, "(00)"" ""00000""-""0000") 'Celular
        
        linhalistbox = linhalistbox + 1
        End With
    conte = conte + 1
    End If
    
    linha = linha + 1
    Wend
End With

Me.lblTotalRegistros = "Total de registros localizados: " & conte

End Sub

'PREENCHER LISTBOX QUANDO INICIAR O FORMULARIO DE PESQUISA
Private Sub UserForm_Initialize()
Call PreencherListBox
Call removeCaption(Me)
End Sub

'POSIÇÃO DO FORMULARIO NA PLANILHA
Private Sub UserForm_Layout()
Me.Move 730, 150
End Sub

Private Sub cmdSair_Click()
Unload frmPesquisa
End Sub

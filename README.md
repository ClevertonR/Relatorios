Codigo feito para informar relatorio de falhas do widnows via VB:


Option Explicit

Private Sub cmdLISTARRELATORIO_Click()

On Error GoTo ErrorHandler

    Shell "perfmon /rel", vbNormalFocus

    Exit Sub

ErrorHandler:

    MsgBox "Ocorreu erro ao abrir o relatorio do windows, sistema sera abortado." & vbCrLf & _
           "Erro: " & Err.Description, vbCritical, "Erro"




End Sub


![image](https://github.com/ClevertonR/Relatorios/assets/51756371/01d4ff92-4a12-4a22-aebb-7374b25bd800)

Após, realizado melhoria onde o codigo nao apresente erros ou travamentos solicitando acesso como admionistrador foi utilizado Api do windows.
Para fazer isso, você pode usar o comando da API do Windows. Aqui está o código atualizado:ShellExecute...




Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Sub cmdLISTARRELATORIO_Click()

On Error GoTo ErrorHandler

    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "runas", "perfmon", "/rel", "", 1)

    If lSuccess = 0 Then
        MsgBox "Não foi possível executar o comando como administrador." & vbCrLf & _
               "Código de erro: " & Err.LastDllError, vbCritical, "Erro"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Ocorreu um erro ao listar os arquivos." & vbCrLf & _
           "Erro: " & Err.Description, vbCritical, "Erro"
End Sub


![image](https://github.com/ClevertonR/Relatorios/assets/51756371/f4133c6e-7252-4582-bafc-6a519717b139)


![image](https://github.com/ClevertonR/Relatorios/assets/51756371/9e05be96-d386-4169-a1ba-20466d61603a)









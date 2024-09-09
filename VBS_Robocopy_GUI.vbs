'---------------------------------------------------------------------------------
' CITRA TI - EXCELÊNCIA EM TI
' SCRIPT PARA CÓPIA DE ARQUIVOS USANDO ROBOCOPY COM INTERFACE GRÁFICA E PROGRESSO
' Autor: luciano@citrait.com.br
' Date: 30/10/2021
' Version: 1.0
'---------------------------------------------------------------------------------
' On Error Resume Next
' Option Explicit




' Possiveis Status de execução wshShell.Exec
STATUS_RUNNING  = 0
STATUS_FINISHED = 1
STATUS_ERROR    = 2

' criando o objeto shell
Set wshShell = CreateObject("WScript.Shell")
Set oFS = CreateObject("Scripting.FileSystemObject")


' Verificando se o script foi executado como privil. de adm.
' Caso não seja, então re-executa o script como adm.
Set command = wshShell.Exec("whoami.exe /groups")
hasAdmin = False
While command.Status = STATUS_RUNNING
    Do While Not command.stdout.AtEndOfStream 
        output = command.stdout.ReadAll()
        If InStr(output, "S-1-16-12288") <> 0 Then
			hasAdmin = True
        End If
    Loop
Wend

If Not hasAdmin Then
	Set wshApp = CreateObject("Shell.Application")
	wshApp.ShellExecute "C:\windows\system32\wscript.exe", wscript.scriptfullname, oFS.GetParentFolderName(wscript.scriptfullname), "runas"
	wscript.Quit
End If



' Obtendo do usuário a origem e destino da copia
DIR_SOURCE = oFS.GetAbsolutePathName( InputBox("Caminho Origem:") )
DIR_DEST  = oFS.GetAbsolutePathName( InputBox("Caminho Destino:") )


' Verificando se os caminhos informados são válidos.
If DIR_SOURCE = "" Then
	wscript.echo "Erro. Diretorio Origem Vazio." & vbCrLf & "Execute novamente o script e informe um caminho válido!"
	wscript.Quit
ElseIf DIR_DEST = "" Then
	wscript.echo "Erro. Diretorio Destino Vazio." & vbCrLf & "Execute novamente o script e informe um caminho válido!"
	wscript.Quit
ElseIf Not oFS.FolderExists(DIR_SOURCE) Then
	wscript.echo "Erro. Diretorio Origem Inválido." & vbCrLf & "Execute novamente o script e informe um caminho válido!"
	wscript.Quit
ElseIf Not oFS.FolderExists(DIR_DEST) Then
	' Tentar criar a pasta destino
	result = oFS.CreateFolder(DIR_DEST)
	If Not oFS.FolderExists(DIR_DEST) Then
		wscript.echo "Erro. Diretorio Destino Inválido." & vbCrLf & "Execute novamente o script e informe um caminho válido!"
		wscript.Quit
	End If
	
	
End If


' Usando o Internet Explorer para exibir o progresso da cópia.
Set IE = CreateObject("InternetExplorer.Application")
IE.Offline = True
IE.StatusBar = False
IE.Addressbar = False
IE.MenuBar = False
IE.ToolBar = False
IE.Visible = True
IE.Width = 760
IE.Height = 420
IE.Navigate("about:blank")

' Mostrando origem e destino para o usuário.
HTMLBody = "<p>Origem: " & DIR_SOURCE & "<br>Destino: " & DIR_DEST & "</p>"
IE.Document.Body.InnerHTML = HTMLBody


' ////////////////////////////////////////////////////////////////////////////////////////////
'                         Obtendo as estatisticas iniciais de cópia.
' ////////////////////////////////////////////////////////////////////////////////////////////
IE.Document.Body.InnerHTML = IE.Document.Body.InnerHTML & "<p>Coletando Estatisticas...</p>"
robocopy_command = "robocopy.exe " & chr(34) & DIR_SOURCE & chr(34) & " " & chr(34) & DIR_DEST & chr(34) & " " & "/L /MIR /NDL /NFL /NC /NP /BYTES /R:0 /W:0"
Set robocopy = wshShell.Exec(robocopy_command)
wshShell.AppActivate "Página em Branco - Internet Explorer"

robocopy_stdoutAll = ""
Do While robocopy.Status = STATUS_RUNNING
	robocopy_stdoutAll = robocopy_stdoutAll & robocopy.stdout.ReadAll()
	wscript.sleep 1000
Loop



' ----------------------------------------
' Parsing da saída do robocopy
' ----------------------------------------
Set regex = New RegExp
TotalBytes = 0

regex.Pattern = "Bytes\s?:\s+\d+\s+(\d+)"
Set matches   = regex.Execute(robocopy_stdoutAll)

For Each match in matches
	TotalBytes = TotalBytes + match.SubMatches(0)
Next

' Mostrando ao usuário a qtde de bytes para copiar.
IE.Document.Body.InnerHTML = IE.Document.Body.InnerHTML & "<p>Bytes para Copiar: " & TotalBytes & "</p>"


' Verificando se a qtde de arquivos para copiar é 0
If TotalBytes = 0 Then
	wscript.echo "Nenhum dado novo para copiar..."
	IE.Quit
	Set IE = Nothing
	Set wshShell = Nothing
	Set robocopy = Nothing
	wscript.Quit
End If




' ////////////////////////////////////////////////////////////////////////////////////////////
'                                     Realizando A cópia
' ////////////////////////////////////////////////////////////////////////////////////////////

Set regex = New RegExp
Set regex2 = New RegExp
CopiedBytes = 0

regex.Pattern = "(Novo Arquivo|New File|Mais Antigo|Older|Mais Recente|Newer)\s+(\d+)\s+.*"
regex2.Pattern = "(Arquivos|Files)\s?:\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)"

' Executando o robocopy
robocopy_command = "robocopy " & chr(34) & DIR_SOURCE & chr(34) & " " & chr(34) & DIR_DEST & chr(34) & " " & "/MIR /NDL /BYTES /R:0 /W:0"
Set robocopy = wshShell.Exec(robocopy_command)
wshShell.AppActivate "Página em Branco - Internet Explorer"


While robocopy.Status <> STATUS_FINISHED  ' enquanto robocopy estiver executando...
	While Not robocopy.StdOut.AtEndOfStream  ' enquanto houver buffer de saida de texto
		output_line = robocopy.stdout.ReadLine()
		' Testando se a linha contem ação de copia.
		Set Matches  = regex.Execute(output_line)
		If Matches.Count > 0 Then
			For Each Match in Matches
				CopiedBytes = CopiedBytes + Match.SubMatches(1)
				progress = Round(CopiedBytes * 100 / TotalBytes)
				HTMLBody = "<p>Origem: " & DIR_SOURCE & "<br>Destino: " & DIR_DEST & "</p>"
				HTMLBody = HTMLBody &"<progress value='" & progress & "' max='100' style='width:100%'></progress>"
				HTMLBody = HTMLBody & "<p>Copiados " & CopiedBytes & " de " & TotalBytes & " Bytes</p>"
				HTMLBody = HTMLBody & "<p>Progresso: " & progress & "%</p>"
				IE.Document.Body.InnerHTML = HTMLBody
			Next
		Else
			' Achamos a linha de resumo. copia ja terminou.
			Set Matches = regex2.Execute(output_line)
			If Matches.Count > 0 Then
				For Each Match in Matches
					HTMLBody = "<p>Origem: " & DIR_SOURCE & "<br>Destino: " & DIR_DEST & "</p>"
					HTMLBody = HTMLBody &"<progress value='" & progress & "' max='100' style='width:100%'></progress>"
					HTMLBody = HTMLBody & "<p>Copiados " & CopiedBytes & " de " & TotalBytes & " Bytes</p>"
					HTMLBody = HTMLBody & "<p>Progresso: " & progress & "%</p>"
					HTMLBody = HTMLBody & "<p>Total de Arquivos: " & Match.SubMatches(1) & _
						"<br>Copiados: " & Match.Submatches(2) & _
						"<br>Ignorados: " & Match.Submatches(3) & _
						"<br>Mismatch: " & Match.Submatches(4) & _
						"<br>Erros: " & Match.Submatches(5) & _
						"<br>Extras: " & Match.Submatches(6) & "</p>"
					IE.Document.Body.InnerHTML = HTMLBody
				Next
			End If
		End If
	Wend
	wscript.sleep 1000
Wend


wshShell.PopUp("Copia Finalizada!")
IE.Quit
Set IE = Nothing
Set wshShell = Nothing
Set robocopy = Nothing

















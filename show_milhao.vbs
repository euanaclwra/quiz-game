Dim opcao, caminho, diretorio, db, rs, sql, n, resp, resp_user, resp_certa, aux, i, pergunta_encontrada, pontos 'Declaração de variáveis

call menu()

'Menu Principal
sub menu()
	opcao = InputBox("ESCOLHA UMA OPCAO: " + vbNewLine + vbnewline &_
			"[1] Jogar" + vbNewLine &_
			"[2] Tutorial" + vbnewline &_
			"[3] Sair", "MENU PRINCIPAL")
			
	if opcao <> "" then	'Só inicia o select case caso alguma opção seja escolhida
						'Isso evita que a mensagem de erro seja exibida mesmo sem o usuário digitar nada
		select case opcao
			case "1" 'Se a opção for 1, o jogo começa
				call jogar()
			case "2" 'Se a opção for 2, exibe o tutorial
				call tutorial()
			case "3" 'Se a opção for 3, o script finaliza
				wscript.quit
			case Else 'Se a opção não for 1 nem 2 nem 3, dá erro
				msgbox("DIGITE UMA OPCAO VALIDA!") 
				call menu()
		end select
	end if
end sub

sub tutorial()
	resp = msgbox("- O algoritmo vai sortear perguntas com 4 opcoes de resposta" + vbnewline &_
				  "- Cada pergunta pode valer 10, 30 ou 50 pontos" + vbnewline &_
				  "- Quando voce erra, perde todos os pontos" + vbnewline + vbnewline &_
				  "Deseja iniciar o jogo?", vbYesNo, "TUTORIAL")
				 
	if resp = vbYes then
		call jogar()
	Else
		call menu()
	end if
end Sub	

sub jogar()
	'Conectando ao banco
	set db=createobject("ADODB.Connection") 'Cria um objeto de conexão
	caminho = WScript.ScriptFullName 'Armazena o caminho completo do arquivo atual
	diretorio = Left(caminho, InStrRev(caminho, "\") - 1) 'Retira o nome do arquivo do caminho, deixando apenas o diretório em que ele está
	db.open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & diretorio & "\questoes.accdb") 'Abre a conexão com o caminho especificado
	
	'Sorteando um número
	Do While True	
		sql = "select count(*) as qtde from tb_questoes where ja_utilizada = 'N'" 'Query para ver quantas perguntas ainda não foram utilizadas
		set rs = db.execute(sql) 'Executa a query
		aux = rs.fields("qtde") 'Armazena a quantidade de perguntas disponíveis
	
		If aux = 0 then 'Se não houver nenhuma pergunta disponível, o loop encerra
			MsgBox("TODAS AS PERGUNTAS JA FORAM UTILIZADAS, O JOGO SE ENCERRARA!")
			Exit Do 
		End If
	
		randomize(second(time)) 'Inicializa a função de números aleatórios
		n = int(rnd * aux) + 1 'Sorteia um número entre 1 e o total de perguntas disponíveis
	
		'Verificando se a pergunta já foi utilizada
		pergunta_encontrada = False 'Define que a pergunta ainda não foi repetida

		Do While Not pergunta_encontrada 'Loop que vai ser executado até que a pergunta não seja repetida 
			sql = "select * from tb_questoes where ja_utilizada = 'N'" 'Query pra selecionar as perguntas disponíveis
			Set rs = db.Execute(sql) 'Executa a query
		
			If Not rs.EOF Then 'Se houverem perguntas disponíveis
				pergunta_encontrada = True 'Define que a pergunta foi repetida e sai do loop
				Exit Do
			End If
		Loop
		
		'Selecionando a pergunta
		sql = "select * from tb_questoes where numero="& n &"" 'Query para selecionar a pergunta correspondente ao número sorteado
		set rs = db.execute(sql) 'Executa a query
		resp_certa = rs.fields(6).value 'Armazena a resposta certa (campo 6 da tabela) na variável resp_certa
		sql = "UPDATE tb_questoes SET ja_utilizada='S' WHERE numero="& n &""
		db.execute(sql)

		'Verifica o retorno da query e exibe a pergunta
		if rs.EOF = false Then 
		   resp_user=inputbox(""& rs.fields(1).value &"" + vbnewline + vbnewline &_ 
									  "[1] "& rs.fields(2).value &"" + vbnewline &_
									  "[2] "& rs.fields(3).value &"" + vbnewline &_
									  "[3] "& rs.fields(4).value &"" + vbnewline &_
									  "[4] "& rs.fields(5).value &"" + vbnewline + vbnewline &_
									  "Escolha a resposta correta:","SHOW DO MILHAO")
			call validar_resposta
		end if	
	Loop
end sub 

sub validar_resposta()

	if resp_user <> "" then 'Só inicia a validação caso alguma opção seja escolhida
							'Isso evita que a mensagem de erro seja exibida mesmo sem o usuário digitar nada
		if resp_user = resp_certa Then	
			If rs.Fields("tipo") = "A" Then 'Se for do tipo A, soma 10 pontos
				pontos = pontos + 10
			ElseIf rs.Fields("tipo") = "B" Then 'Se for do tipo B, soma 30 pontos
				pontos = pontos + 30
			ElseIf rs.Fields("tipo") = "C" Then 'Se for do tipo C, soma 50 pontos
				pontos = pontos + 50
			End If
			
			'Mensagem de acerto
			msgbox("Certa Resposta :D" + vbnewline &_
			"Sua pontuacao eh: " & pontos), vbInformation + vbOKOnly, "UHUL!"
			
			call jogar()
		Else
			pontos = 0 'Zera os pontos
			'Mensagem de erro
			resp = msgbox("Voce errou :(" + vbnewline &_ 
			"A resposta correta era: " & rs.fields(CInt(resp_certa) + 1).value + vbnewline &_ 
			"Deseja jogar novamente?", vbCritical + vbYesNo, "QUE PENA!")
			
			if resp = vbYes then
				call jogar() 'Se o usuário quiser jogar de novo, o jogo reinicia
			Else
				wscript.quit 'Se não, o script finaliza
			end if				
		end if
		Else
			wscript.quit 'Caso o usuário não digite nenhuma resposta e faça outra ação (como clicar em cancelar), o script finaliza
	end if
end sub
			
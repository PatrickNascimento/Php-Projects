


<%
 
	PERCENTUAL_A_TRIBUTAR = 2 '0.70 '70% do faturamento icms


	Function PrimeiraFatura_Proporcional(CLIENTE)
	
		PrimeiraFatura_Proporcional = CDATE("01/01/2000")
		
		Set VER = Server.CreateObject("ADODB.Recordset")
		VER.ActiveConnection = MM_Conexao_STRING
		VER.Source = "SELECT  TOP 1 DAT_ORDEM FROM ORDENS WHERE COD_OTIPO IN (3,4) AND COD_CLIENTE="&CLIENTE&" AND COD_PROVEDOR = "& Session("cod_provedor") &" ORDER BY DAT_ORDEM DESC"
		VER.Open()
		IF NOT VER.EOF THEN
			PrimeiraFatura_Proporcional = CDATE(DAY(VER("DAT_ORDEM"))&"/"&MONTH(VER("DAT_ORDEM"))&"/"&YEAR(VER("DAT_ORDEM")))
		END IF
		VER.Close()

	End Function



	Function FechamentoCaixa_UltimaData()
	
		FechamentoCaixa_UltimaData = CDATE("01/01/2000")
		
		Set VER = Server.CreateObject("ADODB.Recordset")
		VER.ActiveConnection = MM_Conexao_STRING
		VER.Source = "SELECT  TOP 1 DAT_CAIXA FROM FECHAMENTO_CAIXA WHERE STA_CAIXA=1 AND COD_PROVEDOR = "& Session("cod_provedor") &" ORDER BY DAT_CAIXA DESC"
		VER.Open()
		IF NOT VER.EOF THEN
			FechamentoCaixa_UltimaData = CDATE(VER("DAT_CAIXA"))
		END IF
		VER.Close()
	
	End Function


	Function FechamentoCaixa_VerificarUltimoFoiFechado()
	
		
		FechamentoCaixa_VerificarUltimoFoiFechado = false
		Set VER = Server.CreateObject("ADODB.Recordset")
		VER.ActiveConnection = MM_Conexao_STRING
		VER.Source = "SELECT TOP 1 DAT_PAGO_MENSALIDADE FROM MENSALIDADES m "_
				   & "WHERE DAT_PAGO_MENSALIDADE>(GETDATE()-500) and DAT_PAGO_MENSALIDADE<CONVERT(datetime, CONVERT(varchar(10), GETDATE(), 103)) "_
				   & "AND not exists (SELECT DAT_CAIXA FROM FECHAMENTO_CAIXA WHERE m.DAT_PAGO_MENSALIDADE=DAT_CAIXA and STA_CAIXA=1 AND COD_PROVEDOR="& Session("cod_provedor") &") "_
				   & "AND COD_PROVEDOR=" & Session("cod_provedor") & " AND (DEST_MENSALIDADE='FRENTE DE CAIXA') AND (TIP_PAGO_MENSALIDADE='COBRANÇA') AND "_
			 	   & "(STA_MENSALIDADE=1)"
		'VER.Source = "SELECT TOP 1 DAT_PAGO_MENSALIDADE FROM MENSALIDADES "_
		'		   & "WHERE DAT_PAGO_MENSALIDADE>'21/07/2010' and DAT_PAGO_MENSALIDADE<CONVERT(datetime, CONVERT(varchar(10), GETDATE(), 103)) "_
		'		   & "and DAT_PAGO_MENSALIDADE not in (SELECT DAT_CAIXA FROM FECHAMENTO_CAIXA WHERE STA_CAIXA=1 AND COD_PROVEDOR="& Session("cod_provedor") &") "_
		'		   & "AND COD_PROVEDOR=" & Session("cod_provedor") & " AND (DEST_MENSALIDADE='FRENTE DE CAIXA') AND (TIP_PAGO_MENSALIDADE='COBRANÇA') AND "_
		'	 	   & "(STA_MENSALIDADE=1)"		   
		VER.Open()
		IF NOT VER.EOF THEN FechamentoCaixa_VerificarUltimoFoiFechado = VER("DAT_PAGO_MENSALIDADE")
		VER.Close()


		VER.Source = "SELECT TOP 1 DAT_VENC_PAGAMENTO FROM PAGAMENTOS p "_
				   & "WHERE DAT_VENC_PAGAMENTO>(GETDATE()-500) and DAT_VENC_PAGAMENTO<CONVERT(datetime, CONVERT(varchar(10), GETDATE(), 103)) "_
				   & "and not exists (SELECT DAT_CAIXA FROM FECHAMENTO_CAIXA WHERE p.DAT_VENC_PAGAMENTO=DAT_CAIXA AND STA_CAIXA=1 AND COD_PROVEDOR="& Session("cod_provedor") &") "_				   
				   & "AND ORIG_PAGAMENTO='FRENTE DE CAIXA' AND COD_PROVEDOR=" & Session("cod_provedor") 
		'VER.Source = "SELECT TOP 1 DAT_VENC_PAGAMENTO FROM PAGAMENTOS "_
		'		   & "WHERE DAT_VENC_PAGAMENTO>'21/07/2010' and DAT_VENC_PAGAMENTO<CONVERT(datetime, CONVERT(varchar(10), GETDATE(), 103)) "_
		'		   & "and DAT_VENC_PAGAMENTO not in (SELECT DAT_CAIXA FROM FECHAMENTO_CAIXA WHERE STA_CAIXA=1 AND COD_PROVEDOR="& Session("cod_provedor") &") "_				   
		'		   & "AND ORIG_PAGAMENTO='FRENTE DE CAIXA' AND COD_PROVEDOR=" & Session("cod_provedor") 				   
		VER.Open()
		IF NOT VER.EOF THEN FechamentoCaixa_VerificarUltimoFoiFechado = VER("DAT_VENC_PAGAMENTO")
		VER.Close()
	
	
	End Function



	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA PEGAR ULTIMA NOTA  *****************       ************************'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	Function Pegar_ProximaNota(cod_provedor)
	
		Set VER = Server.CreateObject("ADODB.Recordset")
		VER.ActiveConnection = MM_Conexao_STRING
		VER.Source = "SELECT  MAX(NO_NFISCAL) AS ULTIMA "_
				   & "FROM NOTAS_FISCAIS "_
				   & "WHERE SERIE_NFISCAL='"& Session("serie_nf") &"' AND (COD_CLIENTE IN  (SELECT  COD_CLIENTE FROM CLIENTES WHERE COD_PROVEDOR = "& cod_provedor &"))"
		VER.Open()
		IF NOT VER.EOF THEN
			if isnull(VER("ULTIMA")) Then 
				Pegar_ProximaNota = 1
			else
				Pegar_ProximaNota = CDBL(VER("ULTIMA"))+1
			end if
		ELSE
			Pegar_ProximaNota = 0
		END IF
		VER.Close()

	End Function





	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA PEGAR ULTIMA NOTA  *****************       ************************'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	Function Pegar_Operacao(cfop)
		
		x = "SERVIÇOS"
		
		SELECT CASE cfop
			CASE 5933
				x="SERVIÇOS"
			CASE 6933
				x="SERVIÇOS"
			CASE 5302
				x="PRESTAÇÃO SERV. COMUNIC."
			CASE 6302
				x="PRESTAÇÃO SERV. COMUNIC."
			CASE 5303
				x="PRESTAÇÃO SERV. COMUNIC."
			CASE 6303
				x="PRESTAÇÃO SERV. COMUNIC."
		end select
		
		Pegar_Operacao = x
		
	End Function





	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA SOMAR ICMS FATURADO EM NOTAS  ******       ************************'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	Function Valor_TributadoICMS(cod_provedor,mes,ano)
	
		Set VER = Server.CreateObject("ADODB.Recordset")
		VER.ActiveConnection = MM_Conexao_STRING
		VER.Source = "SELECT  SUM(ICMS_NFISCAL) AS TOTAL "_
				   & "FROM NOTAS_FISCAIS "_
				   & "WHERE  STA_NFISCAL=1 AND (MONTH(EMI_NFISCAL) = "& mes &") AND (YEAR(EMI_NFISCAL) = "& ano &") "_
				   & "AND (COD_CLIENTE IN  (SELECT  COD_CLIENTE FROM CLIENTES WHERE COD_PROVEDOR = "& cod_provedor &"))"
		VER.Open()
		IF NOT VER.EOF THEN
			Valor_TributadoICMS = VER("TOTAL")
			if trim(VER("TOTAL"))="" or isNull(VER("TOTAL")) Then Valor_TributadoICMS = 0
		ELSE
			Valor_TributadoICMS = 0
		END IF
		VER.Close()

	End Function



	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA CALCULAR FATURAMENTO ICMS     ******       ************************'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	Function Valor_TributacaoICMS(cod_provedor)
	
		Set VER = Server.CreateObject("ADODB.Recordset")
		VER.ActiveConnection = MM_Conexao_STRING
		VER.Source = "SELECT  SUM(CS.VAL_COBR_SERV) AS TOTAL "_
				   & "FROM  CLIENTES C INNER JOIN "_
				   & "COBRANCAS O ON C.COD_CLIENTE = O.COD_CLIENTE INNER JOIN "_
                   & "COBRANCAS_SERVICOS CS ON O.COD_COBRANCA = CS.COD_COBRANCA INNER JOIN "_
                   & "SERVICOS S ON CS.COD_SERVICO = S.COD_SERVICO "_
				   & "WHERE  (S.NF_SERVICO = 'icms') AND (C.COD_PROVEDOR = "& cod_provedor &") AND (O.END_NFI_COBRANCA <> 0) "_
				   & "AND (O.STA_COBRANCA = 'ATIVO') AND (O.COD_TIPO_COBRANCA > 0)"
		VER.Open()
		IF NOT VER.EOF THEN
			Valor_TributacaoICMS = VER("TOTAL")*Session("icms")
		ELSE
			Valor_TributacaoICMS = 0
		END IF
		VER.Close()

	End Function




	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA FORMATAR DATA VENCIMENTO ***************************************'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	FUNCTION DATA_VENCIMENTO(DIA)
		MES = MONTH(Date)+1
		If MES=13 Then MES = 1
		ANO = YEAR(Date)		
		If (DIA-Day(date))>=6 Then MES = MONTH(Date) 'MES ATUAL
		If Month(Date)=12 and MES=1 Then ANO = ANO + 1 'adiciona ANO
		
 		DATA_VENCIMENTO =  DIA &"/"& MES &"/"& ANO
	END FUNCTION




	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   ORDENACAO DE CARACTERES A ESQ E DIR   *****************************************'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	FUNCTION ORD_DIREITA(N,Digitos,Caracter)
	  aux = ""
	  For i = (len(N)+1) to Digitos
	  	aux = aux & Caracter
	  Next
	  ORD_DIREITA = aux & N
	  
	END FUNCTION
	
	FUNCTION ORD_ESQUERDA(N,Digitos,Caracter)
	  aux = ""
	  For i = (len(N)+1) to Digitos
	  	aux = aux & Caracter
	  Next
	  ORD_ESQUERDA = N & aux 
	  
	END FUNCTION





	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA VERIFICAR ATRASO DE MENSALIDADE            ************************'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	Function Verificar_Multa(cod_cobranca,data_atrasada,VM_atualiza)
	
		Set VER = Server.CreateObject("ADODB.Recordset")
		VER.ActiveConnection = MM_Conexao_STRING
		VER.Source = "SELECT TOP 1 DAT_VENC_MENSALIDADE, MULTA_MENSALIDADE, VAL_COBR_MENSALIDADE, (DATEDIFF(dd, DAT_VENC_MENSALIDADE, DAT_PAGO_MENSALIDADE)-2) AS DIAS "_
					 & " FROM MENSALIDADES WHERE (COD_COBRANCA = "& cod_cobranca &") AND (COD_PROVEDOR="& Session("cod_provedor") & ") "_
					 & " AND DATEDIFF(dd, DAT_VENC_MENSALIDADE, DAT_PAGO_MENSALIDADE)>2 AND (STA_MENSALIDADE=1) AND (MULTA_MENSALIDADE=0) AND (YEAR(DAT_VENC_MENSALIDADE)>=2006) ORDER BY COD_MENSALIDADE"
		VER.CursorType = 2 : VER.CursorLocation = 2 : VER.LockType = 2
		VER.Open()
		IF NOT VER.EOF THEN

			Diario = (VER("VAL_COBR_MENSALIDADE")*0.01) / 30 '1% ao mes
			Multa  = VER("VAL_COBR_MENSALIDADE")*0.02 ' 2% de multa
			data_atrasada   = VER("DAT_VENC_MENSALIDADE")
				
			'ExceÇÃo Mes 10/2008 (10 dias de tolerancia para 10/10/2008 e 15/10/2008)
			tolerancia = 0
			If Session("cod_provedor")=1 and _
			   month(data_atrasada)=10 and year(data_atrasada)=2008 and _
			   day(data_atrasada)<=15 and day(data_atrasada)>=10 Then
			   tolerancia  = 10
			End If


			Juros  = Diario * (VER("DIAS")-tolerancia)
			if Juros<0 Then 
				Juros = 0
				Multa = 0
			End if
			

			if VM_atualiza then 
				VER("MULTA_MENSALIDADE") = 1
				VER.UpDate
			end if
			
		   Verificar_Multa = (Multa+Juros)

			
		ELSE
			Verificar_Multa = 0
		END IF

	End Function



	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA VER SE PLANO J� FATUROU NO PASSADO    *****************************'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	Function Verificar_PlanoFaturadoPassado(SERVICO,CLIENTE,VENCIMENTO)
		
		xx = false
		Set VerP = Server.CreateObject("ADODB.Recordset")
		VerP.ActiveConnection = MM_Conexao_STRING
		VerP.Source = "SELECT TOP 1 S.COD_SERVICO "_
					   & "FROM COBRANCAS O, MENSALIDADES M, SERVICOS S, MENSALIDADES_SERVICOS MS "_
					   & "WHERE O.COD_COBRANCA=M.COD_COBRANCA AND M.COD_MENSALIDADE=MS.COD_MENSALIDADE "_
					   & "	AND S.COD_SERVICO=MS.COD_SERVICO AND O.COD_CLIENTE="& CLIENTE &" "_
					   & "	AND (S.TIP_SERVICO=(SELECT TIP_SERVICO FROM SERVICOS WHERE COD_SERVICO="& SERVICO &") "_
					   & "        OR LEFT(S.DES_SERVICO,10)=(SELECT LEFT(DES_SERVICO,10) FROM SERVICOS WHERE COD_SERVICO="& SERVICO &")) "_
					   & "   AND S.TIP_SERVICO<>'' "_
					   & "	AND DATEDIFF(m, M.DAT_VENC_MENSALIDADE, getdate())=0 "_
					   & "	AND LEFT(S.DES_SERVICO,10)=(SELECT LEFT(DES_SERVICO,10) FROM SERVICOS WHERE COD_SERVICO="& SERVICO &")  "
					   'AND M.DAT_VENC_MENSALIDADE<'"& VENCIMENTO &"' 
		'response.Write VerP.Source
		VerP.Open()
		If Not VerP.EOF Then xx = true
		VerP.Close()
		
		Verificar_PlanoFaturadoPassado = xx 'false 'xx
	
	End Function



	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA LER VALOR TOTAL PARA FATURAR CLIENTE  *****************************'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	Function Calcular_ValorFaturamento(COD_COBRANCA,COD_CLIENTE,DATA)

		if DATA=now() or DATA=date() THen DATA = CDATE("01/"&MONTH(DATA+31)&"/"&year(DATA+31))

		' Pegar Desconto da Cobranca
		DESCONTO = 0
		Set Cob = Server.CreateObject("ADODB.Recordset")
		Cob.ActiveConnection = MM_Conexao_STRING
		Cob.Source = "SELECT DESCONTO, RETENCAO_VALOR, STA_COBRANCA, CONTROLE_HORARIO FROM COBRANCAS WHERE COD_COBRANCA="& COD_COBRANCA '&" AND COD_CLIENTE="& COD_CLIENTE
		Cob.Open()
		IF NOT Cob.EOF THEN 
			CONTROLE_HORARIO = Cob("CONTROLE_HORARIO")
			DESCONTO = Cob("DESCONTO")
			RETENCAO = Cob("RETENCAO_VALOR")
			STATUS	 = Cob("STA_COBRANCA")
		END IF
		Cob.Close()


		' PEGAR VALOR_TOTAL 
		 Possui_Dedicado = false
		 VALOR_TOTAL = FormatNumber(0)
		 Set CONTRATADO = Server.CreateObject("ADODB.Recordset")
		 CONTRATADO.ActiveConnection = MM_Conexao_STRING
		 CONTRATADO.Source = "SELECT C.*, DES_SERVICO FROM COBRANCAS_SERVICOS C, SERVICOS S  WHERE C.COD_SERVICO=S.COD_SERVICO AND COD_COBRANCA=" & COD_COBRANCA
		 CONTRATADO.Open()
		 While NOT CONTRATADO.EOF
			If inStr(CONTRATADO("DES_SERVICO"),"DEDICADO CONDOMINIO") or inStr(CONTRATADO("DES_SERVICO"),"DEDICADO RESIDENCIAL") Then Possui_Dedicado = true
			'If CONTRATADO("DAT_FATURA_COBR_SERV")<=DATA AND CONTRATADO("DAT_NAOFATURA_COBR_SERV")>DATA Then VALOR_TOTAL = (VALOR_TOTAL + CONTRATADO("VAL_COBR_SERV"))
			
			' Promocional
			If CONTRATADO("DAT_FIM_PROMO_COBR_SERV")>=cdate("1/"&month(DATA)&"/"&year(DATA)) _
			   and cdate("1/"&month(CONTRATADO("DAT_FIM_PROMO_COBR_SERV"))&"/"&year(CONTRATADO("DAT_FIM_PROMO_COBR_SERV")))>=cdate("1/"&month(DATA)&"/"&year(DATA)) Then 
			   'CONTRATADO("VAL_PROMO_COBR_SERV")>0 Then 

				VALOR_TOTAL = (VALOR_TOTAL + CONTRATADO("VAL_PROMO_COBR_SERV"))

			' Dentro de data de faturamento ou proporcional
			ElseIf (CONTRATADO("DAT_FATURA_COBR_SERV")<=DATA AND CONTRATADO("DAT_NAOFATURA_COBR_SERV")>DATA) OR datediff("m",CONTRATADO("DAT_FATURA_COBR_SERV"),DATA)=1 Then
				'Or day(CONTRATADO("DAT_FATURA_COBR_SERV"))>1 and month(CONTRATADO("DAT_FATURA_COBR_SERV"))=month(DATA) and year(CONTRATADO("DAT_FATURA_COBR_SERV"))=year(DATA)  Then
				
				' PROPORCIONAL, se ativou mes passado ao vencimento
				If datediff("m",CONTRATADO("DAT_FATURA_COBR_SERV"),DATA)=1 and int(DAY(CONTRATADO("DAT_FATURA_COBR_SERV")))>1 THen
					
					DIAS	= 30 - int(DAY(CONTRATADO("DAT_FATURA_COBR_SERV")))
					if DIAS<=0 Then DIAS = 1 
					VALOR_PROPORCIONAL	= ((CONTRATADO("VAL_COBR_SERV")/30) * DIAS)
					VALOR_TOTAL = formatnumber(VALOR_TOTAL + VALOR_PROPORCIONAL)

				' INTEGRAL
				Elseif datediff("m",CONTRATADO("DAT_FATURA_COBR_SERV"),DATA)<>0 Then	
					VALOR_TOTAL = (VALOR_TOTAL + CONTRATADO("VAL_COBR_SERV"))
				End If
				
			End if

			
	  		CONTRATADO.MoveNext()
		 Wend
		 CONTRATADO.Close()

		' Controle de Horario
		 If CONTROLE_HORARIO AND Possui_Dedicado Then VALOR_TOTAL = FormatNumber(VALOR_TOTAL + Session("val_controle_hor"))

		 ' Desconto
		 If DESCONTO>0 Then VALOR_TOTAL = FormatNumber(VALOR_TOTAL - DESCONTO)
		 
		 ' Retencao imposto - Desconto
		 If RETENCAO>0 Then VALOR_TOTAL = FormatNumber(VALOR_TOTAL - RETENCAO)

		 If VALOR_TOTAL<0 Then VALOR_TOTAL = FormatNumber(0)
		 
		 Calcular_ValorFaturamento = FormatNumber(VALOR_TOTAL)

	End Function



	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA PEGAR TITULO SEQUENCIAL 				      ************************'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	Function Pegar_TituloSeq(cod_provedor,conexao)
		
		 x = "????????"
		
		 ' PEGAR ATUAL
		 Set Pegar = Server.CreateObject("ADODB.Recordset")
		 Pegar.ActiveConnection = conexao
		 Pegar.Source = "SELECT TIT_SEQ_PROVEDOR FROM PROVEDORES WHERE COD_PROVEDOR="& cod_provedor
		 Pegar.CursorType = 0 : Pegar.CursorLocation = 2 : Pegar.LockType = 1 
		 Pegar.Open()
 		 If Not Pegar.EOF Then

	 		 x = Pegar("TIT_SEQ_PROVEDOR")

			 ' ATUALIZAR 
			 set Executar = Server.CreateObject("ADODB.Command")
			 Executar.CommandType = 1 : Executar.CommandTimeout = 0 : Executar.Prepared = true
			 Executar.ActiveConnection = conexao
			 Executar.CommandText = "UPDATE PROVEDORES SET TIT_SEQ_PROVEDOR="& (cdbl(x)+cdbl(1)) &" WHERE COD_PROVEDOR="& cod_provedor
			 Executar.Execute()  
		 
		 End If
		 Pegar.Close()
		 Set Pegar = nothing

		 Pegar_TituloSeq = x

	End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------

	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA PEGAR SERVICOS CONTRATADOS            *****************************'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	Function Pegar_ServicosBoleto(COD_CLIENTE,DATA)
	
		Servicos = ""
		
		' Pegar Controle de Horario
		Set Cob = Server.CreateObject("ADODB.Recordset")
		Cob.ActiveConnection = MM_Conexao_STRING
		Cob.Source = "SELECT DESCONTO, STA_COBRANCA, CONTROLE_HORARIO FROM COBRANCAS WHERE COD_COBRANCA="& COD_COBRANCA '&" AND COD_CLIENTE="& COD_CLIENTE
		Cob.Open()
		IF NOT Cob.EOF THEN 
			if Cob("CONTROLE_HORARIO") Then Servicos = Servicos &"; CONTROLE DE HORARIO DE ACESSO" 
		END IF
		Cob.Close()


		' PEGAR VALOR_TOTAL 
		 Possui_Dedicado = false
		 Set CONTRATADO = Server.CreateObject("ADODB.Recordset")
		 CONTRATADO.ActiveConnection = MM_Conexao_STRING
		 CONTRATADO.Source = "SELECT QTD_COBR_SERV, VAL_COBR_SERV, DAT_FATURA_COBR_SERV, DAT_NAOFATURA_COBR_SERV, DES_SERVICO FROM COBRANCAS_SERVICOS C, SERVICOS S  WHERE C.COD_SERVICO=S.COD_SERVICO AND COD_COBRANCA=" & COD_COBRANCA &" ORDER BY DES_SERVICO" 
		 CONTRATADO.Open()
		 While NOT CONTRATADO.EOF
			If CONTRATADO("DAT_FATURA_COBR_SERV")<=DATA AND CONTRATADO("DAT_NAOFATURA_COBR_SERV")>DATA Then Servicos = Servicos &"; "& CONTRATADO("DES_SERVICO") &" ("& CONTRATADO("QTD_COBR_SERV") &")"
	  		CONTRATADO.MoveNext()
		 Wend
		 CONTRATADO.Close()
		 
		 if Servicos<>"" Then Pegar_ServicosBoleto = mid(Servicos,2)


	End Function




	'/////////////////////////////////////////////////////////////////////////////////////////'
	'		*******   FUNCAO PARA PEGAR SERVICOS DE UMA COBRANCA PARA NOTA FISCAL  *****'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	SUB Pegar_ServicosNF2(COD_MENSALIDADE)
	

			REF_ICMS = 0
			REF_ISS  = 0
				
			' SERVICOS
			Set MSERVICOS = Server.CreateObject("ADODB.Recordset")
			MSERVICOS.ActiveConnection = MM_Conexao_STRING
			MSERVICOS.Source = "SELECT VALOR_MS, QTD_MS, DES_SERVICO, NF_SERVICO, S.COD_SERVICO FROM MENSALIDADES_SERVICOS M, SERVICOS S  WHERE S.COD_SERVICO=M.COD_SERVICO AND COD_MENSALIDADE=" & COD_MENSALIDADE
			MSERVICOS.Open()	 
	 
			 While NOT MSERVICOS.EOF
				
						Valor_Servico = MSERVICOS("VALOR_MS")
						
						Valor_Servico = FormatNumber(Valor_Servico)
						Qtde_Servico  = MSERVICOS("QTD_MS")
						Valor_Unit 	  = Valor_Servico / Qtde_Servico
						Valor_Unit 	  = FormatNumber(Valor_Unit)
						
						IF MSERVICOS("NF_SERVICO")="icms" Then 
							REF_ICMS = REF_ICMS + Valor_Servico
							DESCRICAO_PRODUTOS  = DESCRICAO_PRODUTOS   & MSERVICOS("DES_SERVICO") & CHR(13)
							QUANTIDADE_PRODUTOS = QUANTIDADE_PRODUTOS  & Qtde_Servico & CHR(13)
							VALORUNIT_PRODUTOS  = VALORUNIT_PRODUTOS   & Valor_Unit & CHR(13)
							VALORTOTAL_PRODUTOS  = VALORTOTAL_PRODUTOS & Valor_Servico & CHR(13)
							ALIQUOTA_PRODUTOS  = ALIQUOTA_PRODUTOS & (100*Session("icms"))&"%" & CHR(13)

						ELSEIF MSERVICOS("NF_SERVICO")="iss" Then 
							REF_ISS  = REF_ISS + Valor_Servico
							DESCRICAO_SERVICOS  = DESCRICAO_SERVICOS   & MSERVICOS("DES_SERVICO") & CHR(13)
							QUANTIDADE_SERVICOS = QUANTIDADE_SERVICOS  & Qtde_Servico & CHR(13)
							VALORUNIT_SERVICOS  = VALORUNIT_SERVICOS   & Valor_Unit & CHR(13)
							VALORTOTAL_SERVICOS  = VALORTOTAL_SERVICOS & Valor_Servico & CHR(13)
						
						END IF
							
						'response.Write Valor_Servico &" - "& REF_ICMS &"<br>"
						'DESCRICAO_SERVICOS  = DESCRICAO_SERVICOS   & MSERVICOS("DES_SERVICO") & CHR(13)
						'QUANTIDADE_SERVICOS = QUANTIDADE_SERVICOS  & Qtde_Servico & CHR(13)
						'VALORUNIT_SERVICOS  = VALORUNIT_SERVICOS   & Valor_Unit & CHR(13)
						'VALORTOTAL_SERVICOS  = VALORTOTAL_SERVICOS & Valor_Servico & CHR(13)
					
					
				MSERVICOS.MoveNext()
				
			 Wend
			 
			 MSERVICOS.Close()
			 
			 
			 
			' OUTROS
			Set MOUTROS = Server.CreateObject("ADODB.Recordset")
			MOUTROS.ActiveConnection = MM_Conexao_STRING
			MOUTROS.Source = "SELECT * FROM MENSALIDADES_OUTROS  WHERE COD_MENSALIDADE=" & COD_MENSALIDADE
			MOUTROS.Open()
	 
			 While NOT MOUTROS.EOF
				
						Valor_Servico = MOUTROS("VALOR_MO")
						
						Valor_Servico = FormatNumber(Valor_Servico)
						Qtde_Servico  = 1
						Valor_Unit 	  = Valor_Servico / Qtde_Servico
						Valor_Unit 	  = FormatNumber(Valor_Unit)
						

						IF MOUTROS("NF_MO")="icms" Then 
							REF_ICMS = REF_ICMS + Valor_Servico
							IF INSTR(MOUTROS("DESCRICAO_MO"),"DESCONTO") THEN REF_ICMS  = REF_ICMS - Valor_Servico - Valor_Servico
							DESCRICAO_PRODUTOS  = DESCRICAO_PRODUTOS   & UCASE(MOUTROS("DESCRICAO_MO")) & CHR(13)
							QUANTIDADE_PRODUTOS = QUANTIDADE_PRODUTOS  & Qtde_Servico & CHR(13)
							VALORUNIT_PRODUTOS  = VALORUNIT_PRODUTOS   & Valor_Unit & CHR(13)
							VALORTOTAL_PRODUTOS  = VALORTOTAL_PRODUTOS & Valor_Servico & CHR(13)
							ALIQUOTA_PRODUTOS  = ALIQUOTA_PRODUTOS & (100*Session("icms"))&"%" & CHR(13)
							
						ELSEIF MOUTROS("NF_MO")="iss" or MOUTROS("NF_MO")="" Then 
							REF_ISS  = REF_ISS + Valor_Servico
							IF INSTR(MOUTROS("DESCRICAO_MO"),"DESCONTO") THEN REF_ISS  = REF_ISS - Valor_Servico - Valor_Servico
							DESCRICAO_SERVICOS  = DESCRICAO_SERVICOS   & UCASE(MOUTROS("DESCRICAO_MO")) & CHR(13)
							QUANTIDADE_SERVICOS = QUANTIDADE_SERVICOS  & Qtde_Servico & CHR(13)
							VALORUNIT_SERVICOS  = VALORUNIT_SERVICOS   & Valor_Unit & CHR(13)
							VALORTOTAL_SERVICOS  = VALORTOTAL_SERVICOS & Valor_Servico & CHR(13)
						
						END IF					
					
						'DESCRICAO_SERVICOS  = DESCRICAO_SERVICOS   & UCASE(MOUTROS("DESCRICAO_MO")) & CHR(13)
						'QUANTIDADE_SERVICOS = QUANTIDADE_SERVICOS  & Qtde_Servico & CHR(13)
						'VALORUNIT_SERVICOS  = VALORUNIT_SERVICOS   & Valor_Unit & CHR(13)
						'VALORTOTAL_SERVICOS = VALORTOTAL_SERVICOS  & Valor_Servico & CHR(13)
					
					
				MOUTROS.MoveNext()
				
			 Wend
			 
			 MOUTROS.Close()
	
		

	End Sub






	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA PEGAR SERVICOS DE UMA COBRANCA PARA NOTA FISCAL  *****'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	SUB Pegar_ServicosNF(COD_COBRANCA)
	

			REF_ICMS = 0
			REF_ISS  = 0
			
			' Pegar Desconto da Cobranca
			DESCONTO   			 = 0
			Ja_descontado 		 = false
			Pode_descontar_de_um = false
			
			Set Cob = Server.CreateObject("ADODB.Recordset")
			Cob.ActiveConnection = MM_Conexao_STRING
			Cob.Source = "SELECT DESCONTO, STA_COBRANCA, CONTROLE_HORARIO FROM COBRANCAS WHERE COD_COBRANCA="& COD_COBRANCA '&" AND COD_CLIENTE="& COD_CLIENTE
			Cob.Open()
			IF NOT Cob.EOF THEN 
				CONTROLE_HORARIO = Cob("CONTROLE_HORARIO")
				DESCONTO = Cob("DESCONTO")
				STATUS	 = Cob("STA_COBRANCA")
			END IF
			Cob.Close()
			

	
	
			' PEGAR VALOR_TOTAL 
			 VALOR_TOTAL = Calcular_ValorFaturamento(COD_COBRANCA,0,NOW())

		
			 Set CONTRATADO = Server.CreateObject("ADODB.Recordset")
			 CONTRATADO.ActiveConnection = MM_Conexao_STRING
			 CONTRATADO.Source = "SELECT VAL_COBR_SERV, QTD_COBR_SERV, DAT_FATURA_COBR_SERV, DAT_NAOFATURA_COBR_SERV, DES_SERVICO, NF_SERVICO, C.COD_SERVICO FROM COBRANCAS_SERVICOS C, SERVICOS S  WHERE S.COD_SERVICO=C.COD_SERVICO AND COD_COBRANCA=" & COD_COBRANCA
			 CONTRATADO.CursorType = 2
			 CONTRATADO.CursorLocation = 2
			 CONTRATADO.LockType = 1
			 CONTRATADO.Open()

		 
			 IF DESCONTO>0 THEN
			 	While NOT CONTRATADO.EOF
					if CONTRATADO("VAL_COBR_SERV")>=DESCONTO THEN Pode_descontar_de_um = true ': response.Write CONTRATADO("VAL_COBR_SERV")& " "
					CONTRATADO.MoveNext
				Wend
				CONTRATADO.Requery()
			 END IF


			' Controle de Horario
			 If CONTROLE_HORARIO Then 
			
				Percentual = 0
				If DESCONTO>0 and Not Pode_descontar_de_um Then 
					Percentual 	  = Session("val_controle_hor") / (VALOR_TOTAL+DESCONTO)
					Valor_Servico = VALOR_TOTAL * Percentual
				Else 
					Valor_Servico = Session("val_controle_hor") 
				End If
				Valor_Servico = FormatNumber(Valor_Servico)
				Qtde_Servico  = 1
				Valor_Unit 	  = Valor_Servico

				'Contabiliza Controle de Horario como ISS
				REF_ISS  = REF_ISS + Valor_Servico
				
				DESCRICAO_SERVICOS  = DESCRICAO_SERVICOS   & "CONTROLE DE HORÁRIO DE ACESSO" & CHR(13)
				QUANTIDADE_SERVICOS = QUANTIDADE_SERVICOS  & Qtde_Servico & CHR(13)
				VALORUNIT_SERVICOS  = VALORUNIT_SERVICOS   & Valor_Unit & CHR(13)
				VALORTOTAL_SERVICOS  = VALORTOTAL_SERVICOS & Valor_Servico & CHR(13)

				
			 End if

			 
			 While NOT CONTRATADO.EOF
				
				if CONTRATADO("DAT_NAOFATURA_COBR_SERV")>NOW() Then  
				
					If CONTRATADO("DAT_FATURA_COBR_SERV")>NOW()  Then

						Valor_Servico = FormatNumber(0)
						Qtde_Servico  = CONTRATADO("QTD_COBR_SERV")
						Valor_Unit 	  = Valor_Servico
							
					Else

						' Se pode descontar o desconto de apenas um servico e se nenhum ainda foi descontado
						If DESCONTO>0 and CONTRATADO("VAL_COBR_SERV")>=DESCONTO and Pode_descontar_de_um and Not Ja_descontado Then
							Valor_Servico = CONTRATADO("VAL_COBR_SERV") - DESCONTO
							Ja_descontado = true

						' Se o desconto for fragmentado nos servicos
						ElseIf DESCONTO>0 and Not Pode_descontar_de_um Then
							Percentual 	  = CONTRATADO("VAL_COBR_SERV") / (VALOR_TOTAL+DESCONTO)
							Valor_Servico = VALOR_TOTAL * Percentual

						' Se nao houver desconto ou o desconto ja foi dado em um servico
						Else
							Valor_Servico = CONTRATADO("VAL_COBR_SERV")
						End If
						
						Valor_Servico = FormatNumber(Valor_Servico)
						Qtde_Servico  = CONTRATADO("QTD_COBR_SERV")
						Valor_Unit 	  = Valor_Servico / Qtde_Servico
						Valor_Unit 	  = FormatNumber(Valor_Unit)
						
						IF CONTRATADO("NF_SERVICO")="icms" Then REF_ICMS = REF_ICMS + Valor_Servico
						IF CONTRATADO("NF_SERVICO")="iss" Then REF_ISS  = REF_ISS + Valor_Servico
						'response.Write Valor_Servico &" - "& REF_ICMS &"<br>"
					
					
					End If
					
					DESCRICAO_SERVICOS  = DESCRICAO_SERVICOS   & CONTRATADO("DES_SERVICO") & CHR(13)
					QUANTIDADE_SERVICOS = QUANTIDADE_SERVICOS  & Qtde_Servico & CHR(13)
					VALORUNIT_SERVICOS  = VALORUNIT_SERVICOS   & Valor_Unit & CHR(13)
					VALORTOTAL_SERVICOS  = VALORTOTAL_SERVICOS & Valor_Servico & CHR(13)
					
				End if	
				
				
				CONTRATADO.MoveNext()
				
			 Wend
			 
			 CONTRATADO.Close()
	
		

	End Sub





	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA INSERIR FATURA    *************************************************'
	'/////////////////////////////////////////////////////////////////////////////////////////'
	
	Function Inserir_Mensalidade(TIPO,COD_CLIENTE,VALOR,VENCIMENTO,NOSSONUMERO)
	
		VALOR_SQL  = REPLACE(REPLACE(FORMATNUMBER(VALOR),".","") ,",",".")
		
		
		' PEGAR COD_COBRANCA
		Set COBRANCA = Server.CreateObject("ADODB.Recordset")
		COBRANCA.ActiveConnection = MM_Conexao_STRING
		COBRANCA.Source = "SELECT NOM_CLIENTE, COD_COBRANCA FROM COBRANCAS O, CLIENTES C  WHERE C.COD_CLIENTE=O.COD_CLIENTE AND C.COD_CLIENTE="& COD_CLIENTE 
		COBRANCA.Open()
		COD_COBRANCA = COBRANCA("COD_COBRANCA")
		NOM_CLIENTE = COBRANCA("NOM_CLIENTE")
		COBRANCA.Close()

		
		' INSERIR TITULO
		set INSERE_TITULO = Server.CreateObject("ADODB.Command")
		INSERE_TITULO.ActiveConnection = MM_Conexao_STRING
		INSERE_TITULO.CommandText = "INSERT INTO MENSALIDADES (COD_TIPO_COBRANCA,COD_COBRANCA, REG_VND_COBRANCA, COD_CLIENTE, TIPO_MENSALIDADE, DE_MENSALIDADE, NUM_TITULO_MENSALIDADE, NOSSO_NUM_MENSALIDADE, DAT_VENC_MENSALIDADE, VAL_COBR_MENSALIDADE, COD_PROVEDOR) "_
								  & "VALUES ("& TIPO &", "& COD_COBRANCA &", "& COD_COBRANCA &", "& COD_CLIENTE &", '"& TIPO &"', '"& NOM_CLIENTE &"', "& NOSSONUMERO &", "& NOSSONUMERO &", '"& VENCIMENTO &"', "& VALOR_SQL &", "& Session("cod_provedor") &") "
		INSERE_TITULO.Execute()
		
		
		' PEGAR COD_MENSALIDADE
		Set VerTit = Server.CreateObject("ADODB.Recordset")
		VerTit.ActiveConnection = MM_Conexao_STRING
		VerTit.Source = "SELECT MAX(COD_MENSALIDADE) AS COD_MENSALIDADE FROM MENSALIDADES WHERE COD_COBRANCA="& COD_COBRANCA &" AND COD_PROVEDOR="& Session("cod_provedor") 
		VerTit.Open()
		If Not VerTit.EOF Then COD_MENSALIDADE = VerTit("COD_MENSALIDADE")
		VerTit.Close()
		
		' RETORNAR CODIGO
		Inserir_Mensalidade = COD_MENSALIDADE
		

	End Function
	
	
    '/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA INSERIR FATURA ORDEM   *************************************************'
	'/////////////////////////////////////////////////////////////////////////////////////////'
	
	Function Inserir_Mensalidade_Ordem(TIPO,COD_CLIENTE,VALOR,VENCIMENTO,NOSSONUMERO,COD_ORDEM)
	
		VALOR_SQL  = REPLACE(REPLACE(FORMATNUMBER(VALOR),".","") ,",",".")
		
		
		' PEGAR COD_COBRANCA
		Set COBRANCA = Server.CreateObject("ADODB.Recordset")
		COBRANCA.ActiveConnection = MM_Conexao_STRING
		COBRANCA.Source = "SELECT NOM_CLIENTE, COD_COBRANCA FROM COBRANCAS O, CLIENTES C  WHERE C.COD_CLIENTE=O.COD_CLIENTE AND C.COD_CLIENTE="& COD_CLIENTE 
		COBRANCA.Open()
		COD_COBRANCA = COBRANCA("COD_COBRANCA")
		NOM_CLIENTE = COBRANCA("NOM_CLIENTE")
		COBRANCA.Close()

		
		' INSERIR TITULO
		set INSERE_TITULO = Server.CreateObject("ADODB.Command")
		INSERE_TITULO.ActiveConnection = MM_Conexao_STRING
		INSERE_TITULO.CommandText = "INSERT INTO MENSALIDADES (COD_TIPO_COBRANCA,COD_COBRANCA, REG_VND_COBRANCA, COD_CLIENTE, TIPO_MENSALIDADE, DE_MENSALIDADE, NUM_TITULO_MENSALIDADE, NOSSO_NUM_MENSALIDADE, DAT_VENC_MENSALIDADE, VAL_COBR_MENSALIDADE, COD_PROVEDOR,COD_ORDEM) "_
								  & "VALUES ("& TIPO &", "& COD_COBRANCA &", "& COD_COBRANCA &", "& COD_CLIENTE &", '"& TIPO &"', '"& NOM_CLIENTE &"', "& NOSSONUMERO &", "& NOSSONUMERO &", '"& VENCIMENTO &"', "& VALOR_SQL &", "& Session("cod_provedor") &","& COD_ORDEM &") "
		'Response.Write(INSERE_TITULO.CommandText)
		INSERE_TITULO.Execute()
		
		
		' PEGAR COD_MENSALIDADE
		Set VerTit = Server.CreateObject("ADODB.Recordset")
		VerTit.ActiveConnection = MM_Conexao_STRING
		VerTit.Source = "SELECT MAX(COD_MENSALIDADE) AS COD_MENSALIDADE FROM MENSALIDADES WHERE COD_COBRANCA="& COD_COBRANCA &" AND COD_PROVEDOR="& Session("cod_provedor") 
		VerTit.Open()
		If Not VerTit.EOF Then COD_MENSALIDADE = VerTit("COD_MENSALIDADE")
		VerTit.Close()
		
		' RETORNAR CODIGO
		Inserir_Mensalidade_Ordem = COD_MENSALIDADE
		

	End Function



	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA INSERIR SERVICOS A UM TITULO EMITIDO (MENSALIDADES_OUTROS)    *****'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	Function Inserir_MensalidadeOutro(COD_MENSALIDADE,DESCRICAO_MO,NF_MO,VALOR_MO)

		VALOR_SQL  = REPLACE(REPLACE(FORMATNUMBER(VALOR_MO),".","") ,",",".")

		set INSERIR = Server.CreateObject("ADODB.Command")
		INSERIR.ActiveConnection = MM_Conexao_STRING
		INSERIR.CommandText = "INSERT INTO MENSALIDADES_OUTROS (COD_MENSALIDADE,DESCRICAO_MO,NF_MO,VALOR_MO) VALUES ("& COD_MENSALIDADE &", '"& DESCRICAO_MO &"', '"& NF_MO &"', "& VALOR_SQL &")"
		INSERIR.Execute()

	End Function


	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA INSERIR SERVICOS A UM TITULO EMITIDO (MENSALIDADES_SERVICOS)  *****'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	Function Inserir_MensalidadeServico_Alt(COD_MENSALIDADE,COD_SERVICO,VALOR)

		VALOR_SQL  = REPLACE(REPLACE(FORMATNUMBER(VALOR),".","") ,",",".")

		set INSERIR = Server.CreateObject("ADODB.Command")
		INSERIR.ActiveConnection = MM_Conexao_STRING
		INSERIR.CommandText = "INSERT INTO MENSALIDADES_SERVICOS (COD_MENSALIDADE,COD_SERVICO,COD_SC,PARCELA_MS,VALOR_MS,VALOR_REFERENCIA) VALUES ("& COD_MENSALIDADE &", "& COD_SERVICO &", 0, 1, "& VALOR_SQL &", "& VALOR_SQL &")"
		INSERIR.Execute()

	End Function



	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA INSERIR SERVICOS A UM TITULO EMITIDO (MENSALIDADES_SERVICOS)  *****'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	Function Inserir_MensalidadeServico(COD_MENSALIDADE, COD_COBRANCA, DATA, Cod_Adesao,Val_Adesao)
	
	   'Prepara InsersÃo
		set INSERIR = Server.CreateObject("ADODB.Command")
		INSERIR.ActiveConnection = MM_Conexao_STRING
		 
		IF Cod_Adesao>0 Then
		 
		 	Valor_Inserir = FormatNumber(Val_Adesao)
			Valor_Inserir = Replace( Replace(Valor_Inserir,".","") ,",",".")		

			' Codigo SERVICOS_COMISSIONADOS
			COD_SC = Pegar_Servico_Comissionado(COD_COBRANCA,Cod_Adesao)

			INSERIR.CommandText = "INSERT INTO MENSALIDADES_SERVICOS (COD_MENSALIDADE,COD_SERVICO,COD_SC,PARCELA_MS,VALOR_MS) VALUES ("& COD_MENSALIDADE &", "& Cod_Adesao &", "& COD_SC &", 0, "& Valor_Inserir &")"
			INSERIR.Execute()

		Else		 


		   'Pegar Desconto da Cobranca
			DESCONTO   			 = 0
			Ja_descontado 		 = false
			Pode_descontar_de_um = false
			RETENCAO			 = 0
			
			
			Set Cob = Server.CreateObject("ADODB.Recordset")
			Cob.ActiveConnection = MM_Conexao_STRING
			Cob.Source = "SELECT DESCONTO, RETENCAO_VALOR, RETENCAO_PERCENTUAL, STA_COBRANCA, CONTROLE_HORARIO FROM COBRANCAS WHERE COD_COBRANCA="& COD_COBRANCA '&" AND COD_CLIENTE="& COD_CLIENTE
			Cob.Open()
			IF NOT Cob.EOF THEN 
				CONTROLE_HORARIO = Cob("CONTROLE_HORARIO")
				DESCONTO 		 = Cob("DESCONTO")
				RETENCAO 		 = Cob("RETENCAO_VALOR")
				RETENCAO_PERCENTUAL = Cob("RETENCAO_PERCENTUAL")
				STATUS	 	   = Cob("STA_COBRANCA")
			END IF
			Cob.Close()
	
	
			' RETENCAO DE IMPOSTO
			IF RETENCAO>0 THEN
				INSERIR.CommandText = "INSERT INTO MENSALIDADES_OUTROS (COD_MENSALIDADE,DESCRICAO_MO,NF_MO,VALOR_MO) VALUES ("& COD_MENSALIDADE &", 'RETENÇÃO COFINS, PIS, IRPJ, CSLL = "&Replace(Replace(RETENCAO_PERCENTUAL,".",""),",",".")&"%', 'icms', "& Replace(Replace(RETENCAO,".",""),",",".") &")"
				INSERIR.Execute()
			END IF
			
	
		   'PEGAR VALOR_TOTAL 
			VALOR_TOTAL = Calcular_ValorFaturamento(COD_COBRANCA,0,NOW())
			
			
			' FILTRO EXCECAO GAMBIARRA CLIENTE EMPRESARIAL
			FILTRO_EXCECAO = ""
			' COPAZA
			'IF COD_COBRANCA=790 THEN  FILTRO_EXCECAO = "TOP 1"
	
		
			Set CONTRATADO = Server.CreateObject("ADODB.Recordset")
			CONTRATADO.ActiveConnection = MM_Conexao_STRING
			CONTRATADO.Source = "SELECT "&FILTRO_EXCECAO&" C.*, DES_SERVICO FROM COBRANCAS_SERVICOS C, SERVICOS S  WHERE S.COD_SERVICO=C.COD_SERVICO AND COD_COBRANCA=" & COD_COBRANCA &" ORDER BY VAL_COBR_SERV DESC"
			CONTRATADO.CursorType = 2
			CONTRATADO.CursorLocation = 2
			CONTRATADO.LockType = 1
			CONTRATADO.Open()



			IF DESCONTO>0 THEN
				While NOT CONTRATADO.EOF
					if CONTRATADO("VAL_COBR_SERV")>=DESCONTO THEN Pode_descontar_de_um = true ': response.Write CONTRATADO("VAL_COBR_SERV")& " "
					CONTRATADO.MoveNext
				Wend
				CONTRATADO.Requery()
			END IF
			
			
			'Controle de Horario
			'If CONTROLE_HORARIO Then 
			'
			'	Percentual = 0
			'	If DESCONTO>0 and Not Pode_descontar_de_um Then 
			'		Percentual 	  = Session("val_controle_hor") / (VALOR_TOTAL+DESCONTO)
			'		Valor_Servico = VALOR_TOTAL * Percentual
			'	Else 
			'		Valor_Servico = Session("val_controle_hor") 
			'	End If
			'	Valor_Servico = FormatNumber(Valor_Servico)
			'	Valor_Servico = Replace(Replace(Valor_Servico,".",""),",",".")
			'	
			'	INSERIR.CommandText = "INSERT INTO MENSALIDADES_SERVICOS (COD_MENSALIDADE,COD_SERVICO,COD_SC,PARCELA_MS,VALOR_MS) VALUES ("& COD_MENSALIDADE &", 0, 0, 0, "& Valor_Servico &")"
			'	INSERIR.Execute()
			'	
			'
			'End if
			

			While NOT CONTRATADO.EOF
			
				if CONTRATADO("DAT_NAOFATURA_COBR_SERV")>cdate(DATA) Then 

					If CONTRATADO("DAT_FATURA_COBR_SERV")>cdate(DATA)  Then
				
						Valor_Servico = 0
						Valor_Referencia = 0
							
					Else
	
						' assume valor
						VAL_COBR_SERV = CONTRATADO("VAL_COBR_SERV")
						proporcional  = 0
						
						' assume valor de promocao se estiver
						'If CONTRATADO("DAT_FIM_PROMO_COBR_SERV")>=cdate("1/"&month(cdate(DATA))&"/"&year(cdate(DATA))) and CONTRATADO("VAL_PROMO_COBR_SERV")>0 Then 
						'	VAL_COBR_SERV =  CONTRATADO("VAL_PROMO_COBR_SERV")
						'End if
						
						
						' Promocional
						If CONTRATADO("DAT_FIM_PROMO_COBR_SERV")>=cdate("1/"&month(DATA)&"/"&year(DATA)) _
						   and cdate("1/"&month(CONTRATADO("DAT_FIM_PROMO_COBR_SERV"))&"/"&year(CONTRATADO("DAT_FIM_PROMO_COBR_SERV")))>=cdate("1/"&month(DATA)&"/"&year(DATA)) Then 
			
							VAL_COBR_SERV =  CONTRATADO("VAL_PROMO_COBR_SERV")
			
						' Dentro de data de faturamento ou proporcional
						ElseIf (CONTRATADO("DAT_FATURA_COBR_SERV")<=DATA AND CONTRATADO("DAT_NAOFATURA_COBR_SERV")>DATA) or datediff("m",CONTRATADO("DAT_FATURA_COBR_SERV"),DATA)=1 Then 
							'Or day(CONTRATADO("DAT_FATURA_COBR_SERV"))>1 and month(CONTRATADO("DAT_FATURA_COBR_SERV"))=month(DATA) and year(CONTRATADO("DAT_FATURA_COBR_SERV"))=year(DATA)  Then
							
							' PROPORCIONAL
							'If day(CONTRATADO("DAT_FATURA_COBR_SERV"))>1 and month(CONTRATADO("DAT_FATURA_COBR_SERV"))=month(DATA) and year(CONTRATADO("DAT_FATURA_COBR_SERV"))=year(DATA) and lcase(session("usuario"))="rodrigo" THen
							If datediff("m",CONTRATADO("DAT_FATURA_COBR_SERV"),DATA)=1 and int(DAY(CONTRATADO("DAT_FATURA_COBR_SERV")))>1 Then 
								
								proporcional = 1	
								DIAS	= 30 - int(DAY(CONTRATADO("DAT_FATURA_COBR_SERV")))
								if DIAS<=0 Then DIAS = 1 
								VALOR_PROPORCIONAL	= ((CONTRATADO("VAL_COBR_SERV")/30) * DIAS)
								VAL_COBR_SERV = VALOR_PROPORCIONAL
			
							' INTEGRAL
							Elseif datediff("m",CONTRATADO("DAT_FATURA_COBR_SERV"),DATA)<>0 Then	
							
								VAL_COBR_SERV = CONTRATADO("VAL_COBR_SERV")
							
							End If
							
						End if
						
						
				
						' Se pode descontar o desconto de apenas um servico e se nenhum ainda foi descontado
						If DESCONTO>0 and VAL_COBR_SERV>=DESCONTO and Pode_descontar_de_um and Not Ja_descontado Then
							Valor_Servico = VAL_COBR_SERV - DESCONTO
							Ja_descontado = true
				
						' Se o desconto for fragmentado nos servicos
						ElseIf DESCONTO>0 and Not Pode_descontar_de_um Then
							Percentual 	  = VAL_COBR_SERV / (VALOR_TOTAL+DESCONTO)
							Valor_Servico = VALOR_TOTAL * Percentual
				
						' Se nao houver desconto ou o desconto ja foi dado em um servico
						Else
							Valor_Servico = VAL_COBR_SERV
						End If
	
						' FILTRO EXCECAO GAMBIARRA CLIENTE EMPRESARIAL
						if FILTRO_EXCECAO<>"" Then Valor_Servico = VAL_COBR_SERV
						
						
						Valor_Servico = FormatNumber(Valor_Servico)
						Valor_Servico = Replace(Replace(Valor_Servico,".",""),",",".")
						
						Valor_Referencia = Replace(Replace(FormatNumber(CONTRATADO("VAL_COBR_SERV")),".",""),",",".")
			
					
					End If
					
					Parcela = Pegar_ProxParcela(CONTRATADO("COD_SERVICO"),COD_COBRANCA)	
					
					if Valor_Referencia=0 then Valor_Referencia = Valor_Servico
			
					' Codigo SERVICOS_COMISSIONADOS
					COD_SC = Pegar_Servico_Comissionado(COD_COBRANCA,CONTRATADO("COD_SERVICO"))
			
					
					INSERIR.CommandText = "INSERT INTO MENSALIDADES_SERVICOS (COD_MENSALIDADE,COD_SERVICO,COD_SC,PARCELA_MS,VALOR_MS,VALOR_REFERENCIA,QTD_MS,PROPORCIONAL) VALUES ("& COD_MENSALIDADE &", "& CONTRATADO("COD_SERVICO") &", "& COD_SC &", "& Parcela &", "& Valor_Servico &", "& Valor_Referencia &", "& CONTRATADO("QTD_COBR_SERV") &", "& proporcional &")"
					INSERIR.Execute()
					
				end if				
				
				CONTRATADO.MoveNext()
				
			
			Wend
			
			CONTRATADO.Close()
			
			


		
		END IF
		
		

	End Function
	

	
	' Funcao anterior a 30/04/2008
	'Function Inserir_MensalidadeServico(COD_MENSALIDADE,COD_COBRANCA,Cod_Adesao,Val_Adesao)
	'
	'	'Prepara InsersÃo
	'	 set INSERIR = Server.CreateObject("ADODB.Command")
	'	 INSERIR.ActiveConnection = MM_Conexao_STRING
	'	 
	'	 
	'	 IF Cod_Adesao>0 Then
	'	 
	'		 Valor_Inserir = FormatNumber(Val_Adesao)
	'		 Valor_Inserir = Replace( Replace(Valor_Inserir,".","") ,",",".")		
	'
	'		 ' Codigo SERVICOS_COMISSIONADOS
	'		 COD_SC = Pegar_Servico_Comissionado(COD_COBRANCA,Cod_Adesao)
	'
	'		 INSERIR.CommandText = "INSERT INTO MENSALIDADES_SERVICOS (COD_MENSALIDADE,COD_SERVICO,COD_SC,PARCELA_MS,VALOR_MS) VALUES ("& COD_MENSALIDADE &", "& Cod_Adesao &", "& COD_SC &", 0, "& Valor_Inserir &")"
	'		 INSERIR.Execute()
	'	  
	'	 
	'	 
	'	 ELSE
	'	
	'		' Pegar Desconto da Cobranca
	'		DESCONTO = 0
	'		Set Cob = Server.CreateObject("ADODB.Recordset")
	'		Cob.ActiveConnection = MM_Conexao_STRING
	'		Cob.Source = "SELECT DESCONTO FROM COBRANCAS O, CLIENTES C WHERE C.COD_CLIENTE=O.COD_CLIENTE AND COD_COBRANCA="& COD_COBRANCA &" AND COD_PROVEDOR="& Session("cod_provedor")
	'		Cob.Open()
	'		IF NOT Cob.EOF THEN DESCONTO = Cob("DESCONTO")
	'		Cob.Close()
	'
	'
	'		' PEGAR VALOR_TOTAL 
	'		' Se tiver Desconto recalcula Valores de cada servico referente ao desconto
	'		If DESCONTO>0 THEN
	
	'			 Possui_Dedicado = false
	'			 VALOR_TOTAL = FormatNumber(0)
	'			 Set CONTRATADO = Server.CreateObject("ADODB.Recordset")
	'			 CONTRATADO.ActiveConnection = MM_Conexao_STRING
	'			 CONTRATADO.Source = "SELECT VAL_COBR_SERV, DAT_FATURA_COBR_SERV, DES_SERVICO, C.COD_SERVICO FROM COBRANCAS_SERVICOS C, SERVICOS S  WHERE S.COD_SERVICO=C.COD_SERVICO AND COD_COBRANCA=" & COD_COBRANCA
	'			 CONTRATADO.Open()
	'			 While NOT CONTRATADO.EOF
	'				If inStr(CONTRATADO("DES_SERVICO"),"DEDICADO CONDOMINIO") Then Possui_Dedicado = true
	'				If CONTRATADO("DAT_FATURA_COBR_SERV")<=NOW() Then VALOR_TOTAL = (VALOR_TOTAL + CONTRATADO("VAL_COBR_SERV"))
	'				CONTRATADO.MoveNext()
	'			 Wend
	'	
	'			' Controle de Horario
	'			 If CONTROLE_HORARIO AND Possui_Dedicado Then VALOR_TOTAL = FormatNumber(VALOR_TOTAL + Session("val_controle_hor"))
	'	
	'			 CONTRATADO.Requery()
	'			 While NOT CONTRATADO.EOF
	'				
	'				If CONTRATADO("DAT_FATURA_COBR_SERV")<=NOW() Then
	'
	'					Percentual 	  = CONTRATADO("VAL_COBR_SERV") / VALOR_TOTAL
	'					Valor_Inserir = CONTRATADO("VAL_COBR_SERV") - (Percentual * DESCONTO)
	'					Valor_Inserir = FormatNumber(Valor_Inserir)
	'					Valor_Inserir = Replace( Replace(Valor_Inserir,".","") ,",",".")		
	'					
	'					Parcela = Pegar_ProxParcela(CONTRATADO("COD_SERVICO"),COD_COBRANCA)	
	'			
	'					' Codigo SERVICOS_COMISSIONADOS
	'					COD_SC = Pegar_Servico_Comissionado(COD_COBRANCA,CONTRATADO("COD_SERVICO"))
	'			
	'					
	'					INSERIR.CommandText = "INSERT INTO MENSALIDADES_SERVICOS (COD_MENSALIDADE,COD_SERVICO,COD_SC,PARCELA_MS,VALOR_MS) VALUES ("& COD_MENSALIDADE &", "& CONTRATADO("COD_SERVICO") &", "& COD_SC &", "& Parcela &", "& Valor_Inserir &")"
	'					INSERIR.Execute()
	'				
	'				End If
	'					
	'				CONTRATADO.MoveNext()
	'				
	'			 Wend
	'			 CONTRATADO.Close()
	'
	'
	'		' Inseri os proprios valores dos servicos
	'		ELSE
	'
	'
	'			 Set CONTRATADO = Server.CreateObject("ADODB.Recordset")
	'			 CONTRATADO.ActiveConnection = MM_Conexao_STRING
	'			 CONTRATADO.Source = "SELECT VAL_COBR_SERV, DAT_FATURA_COBR_SERV, DES_SERVICO, C.COD_SERVICO FROM COBRANCAS_SERVICOS C, SERVICOS S  WHERE S.COD_SERVICO=C.COD_SERVICO AND COD_COBRANCA=" & COD_COBRANCA
	'			 CONTRATADO.Open()
	'			 While NOT CONTRATADO.EOF
	'
	'				If CONTRATADO("DAT_FATURA_COBR_SERV")<=NOW() Then
	'					
	'					Valor_Inserir = FormatNumber(CONTRATADO("VAL_COBR_SERV"))
	'					Valor_Inserir = Replace( Replace(Valor_Inserir,".","") ,",",".")		
	'					
	'					Parcela = Pegar_ProxParcela(CONTRATADO("COD_SERVICO"),COD_COBRANCA)	
	'
	'					' Codigo SERVICOS_COMISSIONADOS
	'					COD_SC = Pegar_Servico_Comissionado(COD_COBRANCA,CONTRATADO("COD_SERVICO"))
	'					
	'					INSERIR.CommandText = "INSERT INTO MENSALIDADES_SERVICOS (COD_MENSALIDADE,COD_SERVICO,COD_SC,PARCELA_MS,VALOR_MS) VALUES ("& COD_MENSALIDADE &", "& CONTRATADO("COD_SERVICO") &", "& COD_SC &", "& Parcela &", "& Valor_Inserir &")"
	'					INSERIR.Execute()
	'					
	'				End If
	'
	'				CONTRATADO.MoveNext()
	'				
	'			 Wend
	'			 CONTRATADO.Close()
	'
	'
	'		
	'		END IF
	'	
	'	END IF
	'	
	'	
	'
	'End Function





	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA PEGAR PROXIMA PARCELA DO SERVICO DO CLIENTE  **********************'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	Function Pegar_ProxParcela(COD_SERVICO,COD_COBRANCA)
	
		' Pegar Proxima Parcela
		Prox_Parcela = 1
		Set Mens = Server.CreateObject("ADODB.Recordset")
		Mens.ActiveConnection = MM_Conexao_STRING
		Mens.Source = "SELECT TOP 1 PARCELA_MS FROM MENSALIDADES_SERVICOS MS, MENSALIDADES M WHERE M.COD_MENSALIDADE=MS.COD_MENSALIDADE AND M.COD_COBRANCA="& COD_COBRANCA &" AND MS.COD_SERVICO="& COD_SERVICO &" AND M.COD_PROVEDOR="& Session("cod_provedor") &" ORDER BY PARCELA_MS DESC"
		Mens.Open()
		IF NOT Mens.EOF THEN Prox_Parcela = Mens("PARCELA_MS") + 1
		Mens.Close()

		Pegar_ProxParcela = Prox_Parcela
		
	End Function



	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA Atualizar/Inserir SERVICO COMISSIONADO  ***************************'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	Function Associar_Servico_Comissionado(COD_COBRANCA,COD_SERVICO,COD_SC)
	
		' Deletar
	 	set Del = Server.CreateObject("ADODB.Command")
		Del.ActiveConnection = MM_Conexao_STRING
		Del.CommandText = "DELETE FROM COBRANCAS_SERVICOS_COMISSIONADOS WHERE COD_COBRANCA="& COD_COBRANCA &" AND COD_SERVICO="& COD_SERVICO 
		Del.Execute()

		' Inserir
	 	set Ins = Server.CreateObject("ADODB.Command")
		Ins.ActiveConnection = MM_Conexao_STRING
		Ins.CommandText = "INSERT INTO COBRANCAS_SERVICOS_COMISSIONADOS (COD_COBRANCA,COD_SERVICO,COD_SC) VALUES ("& COD_COBRANCA &","& COD_SERVICO &","& COD_SC &")"

		'response.Write COD_COBRANCA &" . "& COD_SERVICO &" . "&  COD_SC &" <br> " 
		'response.Write Ins.CommandText 
		'response.Flush()

		Ins.Execute()

		
	End Function



	'/////////////////////////////////////////////////////////////////////////////////////////'
	'*******   FUNCAO PARA Pegar Codigo SERVICO COMISSIONADO  ********************************'
	'/////////////////////////////////////////////////////////////////////////////////////////'

	Function Pegar_Servico_Comissionado(COD_COBRANCA,COD_SERVICO)
	
		 x = 0 
		 Set Ver = Server.CreateObject("ADODB.Recordset")
		 Ver.ActiveConnection = MM_Conexao_STRING
		 Ver.Source = "SELECT COD_SC FROM COBRANCAS_SERVICOS_COMISSIONADOS WHERE COD_SERVICO="& COD_SERVICO &" AND COD_COBRANCA="& COD_COBRANCA
		 Ver.Open()
		 If NOT Ver.EOF Then x = Ver("COD_SC")
		 Ver.Close()

		 Pegar_Servico_Comissionado = x
		
	End Function


function aspLog(value)
    response.Write("<script language=javascript>console.log('" & value & "'); </script>")
end function

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'╔═════════════════════════════════════════════════════════════════╗
'║ **** FUNÇÃO PARA UNIFICAR CLIENTE COM MULPLICOS CADASTROS ****  ║
'╚═════════════════════════════════════════════════════════════════╝

Function UnificarClientesMultiploCadastros(COD_CLIENTE,TIPO)

	Set UNIFIQUE = Server.CreateObject("ADODB.Recordset")
	UNIFIQUE.ActiveConnection = MM_Conexao_STRING
	UNIFIQUE.Source = "SELECT COUNT(CPF_CLIENTE) AS numreg, "_
				&" CPF_CLIENTE, (SELECT SUM(VAL_COBR_SERV) "_ 
				&" FROM SIAF_PLUS.PRODUCAO.CLIENTES CLI "_ 
				&" JOIN SIAF_PLUS.producao.COBRANCAS C ON cli.COD_CLIENTE = c.COD_CLIENTE "_ 
				&" JOIN SIAF_PLUS.producao.COBRANCAS_SERVICOS CS ON C.COD_COBRANCA = CS.COD_COBRANCA "_ 	
				&" WHERE CPF_CLIENTE = (SELECT CPF_CLIENTE FROM SIAF_PLUS.producao.CLIENTES WHERE COD_CLIENTE = "& COD_CLIENTE &") AND CLI.COD_CLIENTE <> "& COD_CLIENTE _ 
				&" AND CS.DAT_FATURA_COBR_SERV < GETDATE() "_ 
				&" GROUP BY COD_SERVICO,QTD_COBR_SERV,INS_COBR_SERV,VAL_PROMO_COBR_SERV) as TotalCobranca "_			 
				&" FROM [SIAF_PLUS].producao.CLIENTES c INNER JOIN [SIAF_PLUS].producao.COBRANCAS o ON o.COD_CLIENTE=c.COD_CLIENTE "_
				&" WHERE COD_TIPO_COBRANCA= "& TIPO &" AND c.COD_PROVEDOR= 1 "_			
				&" AND (((STA_COBRANCA IN ('ATIVO','CANCELADO','CJ - REGULARIZANDO')) AND (GER_COBRANCA=1)) or (STA_COBRANCA='BLOQUEADO' AND DATEDIFF(d,DTB_COBRANCA,GETDATE())<30)) "_
				&" AND (VAL_COBRANCA> 0 AND (VAL_PROMO_COBRANCA>0 OR (VAL_PROMO_COBRANCA=0 AND (EXP_PROMO_COBRANCA IS NULL OR EXP_PROMO_COBRANCA<=GETDATE())))) "_
				&" AND NOT EXISTS (SELECT COD_COBRANCA FROM [SIAF_PLUS].producao.MENSALIDADES "_
				&" WHERE o.COD_COBRANCA = COD_COBRANCA AND (DATEDIFF(d, GETDATE(), DAT_VENC_MENSALIDADE) >= 7) AND COD_TIPO_COBRANCA="& TIPO &" ) "_
				&" AND CPF_CLIENTE = (SELECT CPF_CLIENTE FROM [SIAF_PLUS].producao.CLIENTES  WHERE COD_CLIENTE = "& COD_CLIENTE &") "_
				&" GROUP BY CPF_CLIENTE "_
				&" HAVING COUNT(CPF_CLIENTE) > 1 "

	UNIFIQUE.CursorType = 0
	UNIFIQUE.CursorLocation = 2
	UNIFIQUE.LockType = 1
	UNIFIQUE.Open()

	IF not UNIFIQUE.EOF THEN
		VALORCOBRANCA = UNIFIQUE("TotalCobranca")   
	END IF	

	IF VALORCOBRANCA <> "" THEN
		Dim UPDATEUNIFIQUE
		set UPDATEUNIFIQUE = Server.CreateObject("ADODB.Command")
		UPDATEUNIFIQUE.ActiveConnection = MM_Conexao_STRING
		UPDATEUNIFIQUE.CommandText = "IF exists (SELECT * FROM [SIAF_PLUS].producao.COBRANCAS WHERE COD_CLIENTE = "& COD_CLIENTE &")  UPDATE  [SIAF_PLUS].producao.COBRANCAS SET VAL_COBRANCA = " & VALORCOBRANCA & " WHERE COD_CLIENTE= "& COD_CLIENTE		
	UPDATEUNIFIQUE.execute()
	END IF


	'╔══════════════════════════════════════════════════════════════════════════════════════════════╗
	'║ **** 'ADICIONAR AS COBRANÇAS E SERVIÇOS DOS CLIENTES UNIFICADOS NO CLIENTE UNIFICADOR. ****  ║
	'╚══════════════════════════════════════════════════════════════════════════════════════════════╝

	DIM DELETESERVICOSUNIFICADO
	set DELETESERVICOSUNIFICADO = Server.CreateObject("ADODB.Command")
		DELETESERVICOSUNIFICADO.ActiveConnection = MM_Conexao_STRING
		DELETESERVICOSUNIFICADO.CommandText = "DELETE FROM SIAF_PLUS.producao.COBRANCAS_SERVICOS WHERE COD_COBRANCA = (SELECT COD_COBRANCA FROM SIAF_PLUS.producao.COBRANCAS WHERE COD_CLIENTE = "& COD_CLIENTE &")"
		DELETESERVICOSUNIFICADO.Execute()


	DIM UNIFIQUESERVICOS
	SET UNIFIQUESERVICOS = Server.CreateObject("ADODB.Recordset")
	UNIFIQUESERVICOS.ActiveConnection = MM_Conexao_STRING
	UNIFIQUESERVICOS.Source = "INSERT INTO SIAF_PLUS.producao.COBRANCAS_SERVICOS "_
    &" (COD_COBRANCA,COD_SERVICO,QTD_COBR_SERV,INS_COBR_SERV,VAL_COBR_SERV,VAL_PROMO_COBR_SERV,DAT_FATURA_COBR_SERV) "_
    &" SELECT "_
    &" (SELECT COD_COBRANCA FROM SIAF_PLUS.producao.COBRANCAS WHERE COD_CLIENTE = "& COD_CLIENTE &") AS COD_COBRANCA, "_
    &" COD_SERVICO, "_
    &" COUNT(COD_SERVICO) as QTD_COBR_SERV, "_
    &" SUM(INS_COBR_SERV) as INS_COBR_SERV, "_
    &" SUM(VAL_COBR_SERV) as TOTAL, "_
    &" SUM(VAL_PROMO_COBR_SERV) as VAL_PROMO_COBR_SERV, "_
    &" (GETDATE()-DAY(GETDATE())+1) AS DAT_FATURA_COBR_SERV "_
    &" FROM "_
    &" SIAF_PLUS.PRODUCAO.CLIENTES CLI "_
    &" JOIN SIAF_PLUS.producao.COBRANCAS C ON cli.COD_CLIENTE = c.COD_CLIENTE "_ 
    &" JOIN SIAF_PLUS.producao.COBRANCAS_SERVICOS CS ON C.COD_COBRANCA = CS.COD_COBRANCA "_     
    &" WHERE "_
    &" CPF_CLIENTE = (SELECT CPF_CLIENTE FROM SIAF_PLUS.producao.CLIENTES WHERE COD_CLIENTE = "& COD_CLIENTE &") "_
    &" AND CLI.COD_CLIENTE <> "& COD_CLIENTE &" "_
    &" AND CS.DAT_FATURA_COBR_SERV < GETDATE() "_
    &" GROUP BY COD_SERVICO "

	UNIFIQUESERVICOS.Open()

End Function

'╔══════════════════════════════════════════════════════════════════════════════════════════════╗
'║ **** MONTA PDF DEMONSTRATIVO PARA CLIENTE UNIFICADO, ESPECIFICANDO OS SERVIÇOS INDIVIDUAIS*  ║
'╚══════════════════════════════════════════════════════════════════════════════════════════════╝

Function GeraPdfServicosUnificado(COD_CLIENTE)

	DIM UNIFIQUEPDF
	SET UNIFIQUEPDF = Server.CreateObject("ADODB.Recordset")
	UNIFIQUEPDF.ActiveConnection = MM_Conexao_STRING
	UNIFIQUEPDF.Source = "SELECT CLI.COD_CLIENTE, CLI.NOM_CLIENTE, CLI.APELIDO,  S.DES_SERVICO AS SERVICO, QTD_COBR_SERV, VAL_COBR_SERV AS VALOR "_
	&" FROM SIAF_PLUS.PRODUCAO.CLIENTES CLI "_
	&" JOIN SIAF_PLUS.producao.COBRANCAS C ON cli.COD_CLIENTE = c.COD_CLIENTE "_
	&" JOIN SIAF_PLUS.producao.COBRANCAS_SERVICOS CS ON C.COD_COBRANCA = CS.COD_COBRANCA "_	
	&" JOIN SIAF_PLUS.producao.SERVICOS S ON S.COD_SERVICO = CS.COD_SERVICO "_
	&" WHERE CPF_CLIENTE = (SELECT CPF_CLIENTE FROM SIAF_PLUS.producao.CLIENTES WHERE COD_CLIENTE = " & COD_CLIENTE &") AND CLI.COD_CLIENTE <> "& COD_CLIENTE &" "_
	&" AND CS.DAT_FATURA_COBR_SERV < GETDATE() "
	UNIFIQUEPDF.Open()

	DIM TOTALPDF
	SET TOTALPDF = Server.CreateObject("ADODB.Recordset")
	TOTALPDF.ActiveConnection = MM_Conexao_STRING
	TOTALPDF.Source = "SELECT SUM(VAL_COBR_SERV) AS VALOR "_
	&" FROM SIAF_PLUS.PRODUCAO.CLIENTES CLI "_
	&" JOIN SIAF_PLUS.producao.COBRANCAS C ON cli.COD_CLIENTE = c.COD_CLIENTE "_
	&" JOIN SIAF_PLUS.producao.COBRANCAS_SERVICOS CS ON C.COD_COBRANCA = CS.COD_COBRANCA "_	
	&" JOIN SIAF_PLUS.producao.SERVICOS S ON S.COD_SERVICO = CS.COD_SERVICO "_
	&" WHERE CPF_CLIENTE = (SELECT CPF_CLIENTE FROM SIAF_PLUS.producao.CLIENTES WHERE COD_CLIENTE = " & COD_CLIENTE &") AND CLI.COD_CLIENTE <> "& COD_CLIENTE &" "_
	&" AND CS.DAT_FATURA_COBR_SERV < GETDATE() "
	TOTALPDF.Open()

	VALORTOTAL   = ""&FORMATNUMBER(TOTALPDF("VALOR"),2)&"" 

	'OBTEM O MÊS DE REFERÊNCIA
	DIM MESPRESTACAO
	SET MESPRESTACAO = Server.CreateObject("ADODB.Recordset")
	MESPRESTACAO.ActiveConnection = MM_Conexao_STRING
	MESPRESTACAO.Source = "select MES_PRESTACAO from SIAF_PLUS.producao.MENSALIDADES where COD_CLIENTE = " & COD_CLIENTE &""
	MESPRESTACAO.Open
	MESAPURACAO = MESPRESTACAO("MES_PRESTACAO")
	 
	HTML = HTML & "<html>"
	HTML = HTML & "<head>"
	HTML = HTML & "<font face='calibri'>"                     
	HTML = HTML & "<style>"
	HTML = HTML & "table, th, td {"
	HTML = HTML & "border: 1px solid black;"
	HTML = HTML & "border-collapse: collapse;"
	HTML = HTML & "} th, td {"
	HTML = HTML & "padding: 5px;"
	HTML = HTML & "} th {"
	HTML = HTML & "text-align: left;"
	HTML = HTML & "}"
	HTML = HTML & "</style>"
	HTML = HTML & "</head>"
	HTML = HTML & "<body>"
	HTML = HTML & "<h3>APURACAO DOS SERVICOS REFERENTE AO MES "&MESAPURACAO&"</h3>"
	HTML = HTML & "<p>CLIENTE : ASSOCIACAO DOS REGISTRADORES CIVIS DAS PESSOAS DE SC - ARPEN</p>"
	HTML = HTML & "<table style='width:1000'>"
	HTML = HTML & "<tr>"
	HTML = HTML & "<th>CADASTRO</th>"
	HTML = HTML & "<th>SERVICO</th>"
	HTML = HTML & "<th>QUANTIDADE</th>"
	HTML = HTML & "<th>VALOR</th>"
	HTML = HTML & "</tr>"

	While NOT UNIFIQUEPDF.EOF
		COD_CLIENTE = UNIFIQUEPDF("COD_CLIENTE")   
		    APELIDO = UNIFIQUEPDF("APELIDO")   
		    SERVICO = UNIFIQUEPDF("SERVICO")   
		       QTDE = UNIFIQUEPDF("QTD_COBR_SERV")   
		      VALOR = ""&FORMATNUMBER(UNIFIQUEPDF("VALOR"),2)&"" 

		HTML = HTML & "<tr>"
		HTML = HTML & "<td>"& APELIDO &"</td>"
		HTML = HTML & "<td>"& SERVICO &"</td>"
		HTML = HTML & "<td style='text-align: center'>"& QTDE &"</td>"
		HTML = HTML & "<td style='text-align: right'> R$ "& VALOR &"</td>"
		HTML = HTML & "</tr>"
		UNIFIQUEPDF.MoveNext()	
	Wend

	HTML = HTML & "</table>"
	HTML = HTML & "<table style='width:1000'> <tR>"
	HTML = HTML & "<td WIDTH=555>SUBTOTAL</td>"
	HTML = HTML & "<td style='text-align: right'> <B>R$ "& VALORTOTAL &"</B></td>"
	HTML = HTML & "</table>"
	HTML = HTML & "</body>"
	HTML = HTML & "</html>"

	While NOT UNIFIQUEPDF.EOF
		APELIDO = UNIFIQUEPDF("APELIDO")   
		SERVICO = UNIFIQUEPDF("SERVICO")   
		VALOR = ""&FORMATNUMBER(UNIFIQUEPDF("VALOR"),2)&"" 

		UNIFIQUEPDF.MoveNext()	
	Wend

	'enviar_email Session("email_usuario"), "Cadastro", "contato@arpen-sc.org.br", Assunto, HTML
	'###########################################################
	'DESCOMENTAR A LINHA DE CIMA PARA ENVIAR AO USUÁRIO CORRETO.
	 'enviar_email "noreply@engeplus.com.br", "Cadastro", "financeiro@engeplus.com.br", "Assunto", HTML
	 'enviar_email "noreply@engeplus.com.br", "Cadastro", "rodrigo@engeplus.com.br", "Assunto", HTML
	 enviar_email "noreply@engeplus.com.br", "Cadastro", "psn1462@gmail.com", "Assunto", HTML

End Function
%>

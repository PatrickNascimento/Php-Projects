Alteração na Tabela Cobrança Fidelidade

```
ALTER TABLE SIAF_PLUS.dbo.COBRANCAS_FIDELIDADES
ADD ADESAO_ANTECIPADA bit NOT NULL Default (0);
```


ALTER TABLE SIAF_PLUS.producao.COBRANCAS_SERVICOS
ADD ADESAO_ANTECIPADA bit NOT NULL Default (0);


	' RESIDENCIAL 

	IF instr(DES_SERVICO,"RESIDENCIAL") and instr(DES_SERVICO,"SMART PLUS") Then	  

	%>
		SMART PLUS RESIDENCIAL		
	<%

	ELSE IF instr(DES_SERVICO,"RESIDENCIAL") and instr(DES_SERVICO,"BANDA LARGA") Then

	%>				 
	    BANDA LARGA RESIDENCIAL		
	<%
	ELSE	

	 IF instr(DES_SERVICO,"CONDOMINIO") and instr(DES_SERVICO,"SMART PLUS") Then	 

	%>
	   	SMART PLUS CONDOMINIO		
	<%

	ELSE IF instr(DES_SERVICO,"CONDOMINIO") and instr(DES_SERVICO,"BANDA LARGA") Then

	%>				 
		BANDA LARGA CONDOMINIO			  
	<%
	 ELSE  
	%>
		OUTRO TIPOS (MANUAL)
	<%
	
	END IF
	End IF 
	END iF
	END iF



CADASTRO_CONSULTAR_CLIENTE_CONTRATO.ASP

<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Func.asp" -->
<!--#include file="Connections/Conexao.asp" -->

<% 
'/////////////////////////////////////////////////////////////////////////////////////////'
'*********  VERIFICA PERMISSAO DE ACESSO A PAGINA ****************************************'
'/////////////////////////////////////////////////////////////////////////////////////////'

IF Session("nivel_usuario")<20 and Session("nivel_usuario")<>11 THEN  '****  ADM  ****'
	RESPONSE.WRITE "<script language='javascript'>location.assign('Logoff.asp')</script>"
	RESPONSE.End() 'TERMINA O CODIGO DESTA PAGINA
END IF

'*****************************************************************************************'
%>

<%
COD_COBRANCA = Trim(Request("cobranca"))
COD_SERVICO = Trim(Request("servico"))
COD_CLIENTE = Trim(Request("cod_cliente"))
COD_CONDOMINIO = Trim(Request("condominio"))
CONTRATO	= ""
INS_SERVICO = 0
VAL_SERVICO = 0
DESCONTO    = 0
PROMOCAO	= 0

CONDICAO_SERVICO = ""
CONDICAO_SERVICO2 = "(CIDADE=(select CID_CLIENTE FROM producao.CLIENTES where COD_CLIENTE="&COD_CLIENTE&") AND UF=(select EST_CLIENTE FROM producao.CLIENTES where COD_CLIENTE="&COD_CLIENTE&"))"
CONDICAO_SERVICO3 = ""
IF COD_CONDOMINIO<>"" Then 
	CONDICAO_SERVICO  = "(TIPO = (SELECT TIP_SERVICO FROM CONDOMINIOS WHERE COD_CONDOMINIO="& COD_CONDOMINIO &") OR TIPO = (SELECT TIP2_SERVICO FROM CONDOMINIOS WHERE COD_CONDOMINIO="& COD_CONDOMINIO &")) AND "
	CONDICAO_SERVICO2 = "(P.COD_CONDOMINIO="&COD_CONDOMINIO&" OR "& CONDICAO_SERVICO2 &")"
	CONDICAO_SERVICO3 = "AND TIP_SERVICO = (SELECT TIP_SERVICO FROM producao.CONDOMINIOS WHERE COD_CONDOMINIO="&COD_CONDOMINIO&")"
END IF




'SERVICOS
Set Servicos = Server.CreateObject("ADODB.Recordset")
Servicos.ActiveConnection = MM_Conexao_STRING
'Servicos.Source = "SELECT * FROM view_SERVICOS WHERE "& CONDICAO_SERVICO &" CONTRATO<>'' AND CONTRATO IS NOT NULL AND (VALOR>0 OR OBS=1) AND COD_PROVEDOR="& Session("cod_provedor") &" AND NIVEL_CONSULTA<="& Session("nivel_usuario") &" ORDER BY TIPO DESC,NOME,VALOR"
				'= "SELECT COD_SERVICO, DES_SERVICO, COD_SERVICO, TIP_SERVICO, CON_SERVICO, VAL_SERVICO, INS_SERVICO  FROM SERVICOS "_
				'& "WHERE CON_SERVICO IS NOT NULL AND CON_SERVICO<>'' AND (VAL_SERVICO>0 OR OBS_SERVICO=1) AND COD_PROVEDOR="& Session("cod_provedor") &" "_
				'& "ORDER BY DES_SERVICO,VAL_SERVICO,TIP_SERVICO"

' RODRIGO TEMP ********-/-*-*/-*/-*/-/-*/-*/-/-/-*/-*/-/-*/-*/-*/-/-*/-*/-*/-*/-/-/-/-/-*/-/-*/-*/-*
'if Session("usuario")="rodrigo" then

	sql = "SELECT S.COD_SERVICO, S.DES_SERVICO AS NOME, S.INS_SERVICO AS ADESAO, CON_SERVICO AS CONTRATO, TIP_SERVICO AS TIPO, S.CORPORATIVO, "_
		& "S.VAL_SERVICO AS VALOR, P.DESCONTO_FIXO AS DESCONTO, P.MESES_VALOR AS MESES_PROMOCAO, P.VALOR AS VALOR_PROMOCIONAL, "_
		& "CASE WHEN (P.COD_PROMOCAO IS NOT NULL) THEN 1 ELSE 0 END AS PROMOCAO, NIV_USR_CON_SERVICO as NIVEL_CONSULTA "_
		& "FROM producao.SERVICOS S LEFT JOIN dbo.view_PROMOCOES_VIGENTES P "_
		& "ON S.COD_SERVICO=P.COD_SERVICO AND "& CONDICAO_SERVICO2 &" "_
		& "WHERE MENSAL=1 AND NIV_USR_CON_SERVICO<200 and S.COD_PROVEDOR=1 "& CONDICAO_SERVICO3 &" " _
		& "AND NIV_USR_CON_SERVICO<="& Session("nivel_usuario") &" " _
		& "ORDER BY CIDADE,NOME"
		
	
	Servicos.Source = sql
	'response.Write sql

'end if

'RESPONSE.Write Servicos.Source 
Servicos.Open()
encontrouServico = 0
CORPORATIVO		 = false
While (NOT Servicos.EOF)
	IF cstr(COD_SERVICO)=cstr(Servicos("COD_SERVICO")) and encontrouServico=0 Then 
		CONTRATO 	= Servicos("CONTRATO")
		TIP_SERVICO = Servicos("TIPO")
		DES_SERVICO = Servicos("NOME")
		CORPORATIVO = Servicos("CORPORATIVO")
		INS_SERVICO = Servicos("ADESAO")
		VAL_SERVICO = Servicos("VALOR")
		DESCONTO	= Servicos("DESCONTO") 'formatnumber(Servicos("DESCONTO"))
		if isnull(DESCONTO) Then DESCONTO=0
		PROMOCAO	= Servicos("PROMOCAO")
		MESES_PROMO = Servicos("MESES_PROMOCAO")
		VALOR_PROMO = Servicos("VALOR_PROMOCIONAL")
		encontrouServico = 1
	END IF
	Servicos.MoveNext
Wend
Servicos.Requery()


CONDICAO_CONDOMINIO = ""
'IF COD_SERVICO<>"" Then CONDICAO_CONDOMINIO = "(TIP_SERVICO = (SELECT TIP_SERVICO FROM SERVICOS WHERE COD_SERVICO="& COD_SERVICO &") OR TIP2_SERVICO = (SELECT TIP_SERVICO FROM SERVICOS WHERE COD_SERVICO="& COD_SERVICO &")) AND "



'CONDOMINIOS
Set Condominios = Server.CreateObject("ADODB.Recordset")
Condominios.ActiveConnection = MM_Conexao_STRING
Condominios.Source = "SELECT 0 AS X, COD_CONDOMINIO, NOM_CONDOMINIO, TIP_SERVICO FROM SIAF_PLUS.producao.CONDOMINIOS Z WHERE "& CONDICAO_CONDOMINIO &" TIP_SERVICO = 'ponto_fbr_res' AND COD_PROVEDOR = 1 AND NOM_CONDOMINIO NOT LIKE '%EMP FIBRA%'"_
& "UNION ALL "_
& "SELECT DISTINCT 1 AS X, CON.COD_CONDOMINIO, CON.NOM_CONDOMINIO, CON.TIP_SERVICO FROM SIAF_PLUS.PRODUCAO.CLIENTES CLI INNER JOIN SIAF_PLUS.PRODUCAO.PONTOS P ON CLI.COD_CLIENTE = P.COD_CLIENTE INNER JOIN SIAF_PLUS.PRODUCAO.CONDOMINIOS CON ON CON.COD_CONDOMINIO = P.COD_CONDOMINIO INNER JOIN (SELECT COD_CLIENTE AS CODIGO, CID_CLIENTE AS CIDADE, END_CLIENTE AS ENDERECO, CEP_CLIENTE AS CEP FROM SIAF_PLUS.PRODUCAO.CLIENTES WHERE COD_CLIENTE =  "&COD_CLIENTE&" ) AS DADOS_CLIENTE ON CLI.CEP_CLIENTE = DADOS_CLIENTE.CEP AND CLI.CID_CLIENTE = DADOS_CLIENTE.CIDADE COLLATE Latin1_General_CI_AI AND CLI.END_CLIENTE = DADOS_CLIENTE.ENDERECO COLLATE Latin1_General_CI_AI AND CLI.COD_CLIENTE <> DADOS_CLIENTE.CODIGO WHERE "& CONDICAO_CONDOMINIO &" TIP_SERVICO <> 'ponto_fbr_res' AND CON.COD_PROVEDOR = 1 AND  NOM_CONDOMINIO NOT LIKE '%EMP FIBRA%' " _
& "UNION ALL "_
& "SELECT 1 AS X, CON.COD_CONDOMINIO, CON.NOM_CONDOMINIO, CON.TIP_SERVICO FROM SIAF_PLUS.PRODUCAO.CONDOMINIOS CON WHERE CON.COD_CONDOMINIO NOT IN (SELECT DISTINCT COD_CONDOMINIO FROM  SIAF_PLUS.PRODUCAO.PONTOS WHERE COD_CLIENTE <> "&COD_CLIENTE&") AND CON.TIP_SERVICO <> 'ponto_fbr_res' AND CON.COD_PROVEDOR = 1 AND CON.NOM_CONDOMINIO not like '%EMP FIBRA%' ORDER BY X DESC, NOM_CONDOMINIO "
Condominios.Open()

%>


<%
'/////////////////////////////////////////////////////////////////////////////////////////'
'*******   REQUISITA LISTA DE PRODUTOS   *************************************************'
'/////////////////////////////////////////////////////////////////////////////////////////'
Set PRODUTOS = Server.CreateObject("ADODB.Recordset")
PRODUTOS.ActiveConnection = MM_Conexao_STRING
PRODUTOS.Source = "SELECT COD_PRODUTO, NOM_PRODUTO  FROM PRODUTOS  WHERE STA_PRODUTO=1 AND (COD_PROVEDOR="& Session("cod_provedor") &")  ORDER BY NOM_PRODUTO"
PRODUTOS.Open()
'*****************************************************************************************'
%>




<html>
<head>
<title>Contrato</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script src="js/jquery-1.7.2.js" language="javascript"></script>
<link href="Style.css" rel="stylesheet" type="text/css">
<link href="js/select2.min.css" rel="stylesheet" />
<script src="js/select2.min.js"></script>
<style>
	.select2-selection, .select2-selection__rendered, .select2-selection__arrow {
		height: 20px !important;
		line-height: inherit !important;
		color: #000099 !important;
		font-family:Verdana, Arial, Helvetica, sans-serif !important;
		font-size: 12px !important;
		text-transform: uppercase !important;
	}
	
	.select2-dropdown {
		left: 8px !important;
		color: #000099 !important;
		font-family:Verdana, Arial, Helvetica, sans-serif !important;
		font-size: 12px !important;
		text-transform: uppercase !important;
	}
	.produto td {
		border: 1px solid #CCC;
		font-family:Verdana, Arial, Helvetica, sans-serif;
		font-size: 10px;
	
	}
	.produto th {
		font-family:Verdana, Arial, Helvetica, sans-serif;
		font-size: 10px;
	
	}
</style>
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>

<body background="imagens/fundo.gif">

 <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-bottom:5px;">
   <tr>
      <td height="16" bgcolor="#EFEFEF"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><font color="#000000" size="1">&nbsp;CADASTRO 
      <font color="#000066">&gt;&gt;</font> CONSULTAR <font color="#000066">&gt;&gt;</font> CLIENTE <font color="#000066">&gt;&gt;</font> CONTRATO</font></font></td>
   </tr>
</table>
	
    <script type="text/javascript">
    function adesao(){
		console.log('adesao');
		/*
		VM = document.getElementById('ADESAO_MINIMA').value;
		VE = document.getElementById('ENTRADA').value;
		VP = document.getElementById('VALOR_PARCELADO').value;
		VM = parseFloat(VM.replace(",",".","ig"));
		VE = parseFloat(VE.replace(",",".","ig"));
		VP = parseFloat(VP.replace(",",".","ig"));
				
		if(document.getElementById('sem_fidelidade').checked!=1){
			if((VE+VP)<VM && (VE+VP)!=0 && (VE+VP)!=180 && (VE+VP)!=100 && (VE+VP)!=150 && (VE+VP)!=80 && (VE+VP)!=230) alert("Valor mínimo de adesão não alcançado.");
			else document.getElementById('formulario').submit();
		} else document.getElementById('formulario').submit();
		*/
		document.getElementById('formulario').submit();
		
		//alert(VE+''+VP+''+VM);
		//return false;
		
	};
	
	window.onload = function() {
		var btnAdesao = document.getElementById('btn-adesao');
		var selectComoConheceu = document.getElementById('como-conheceu');
		selectComoConheceu.onchange = function() {
			var comoConheceu = selectComoConheceu.options[selectComoConheceu.selectedIndex].value;
			btnAdesao.disabled = comoConheceu === '';
		};	
	};
	
	$(document).ready(function() {
		$('.campo-condomino').select2({
			width:'400px',
			height:'10px'
		});
		$('#GRATUIDADE_SEM_FIDELIDADE').click(function(){
			const gratuidadeSemFidelidade = $(this).is(':checked');
			$('#ENTRADA').attr('readonly', gratuidadeSemFidelidade);
			$('#VALOR_PARCELADO').attr('readonly', gratuidadeSemFidelidade);
			$('#INST_DESCONTO').attr('readonly', gratuidadeSemFidelidade);
			$('#parcelas').attr('readonly', gratuidadeSemFidelidade);
			$('#sem_fidelidade').attr('readonly', gratuidadeSemFidelidade);
			if (gratuidadeSemFidelidade) {
				$('#ENTRADA').val('0,00');
				$('#VALOR_PARCELADO').val(0);
				$('#INST_DESCONTO').val(0);
				$('#parcelas').attr('readonly', true);
				$('#sem_fidelidade').attr('checked', true);
				$('#ADESAO_TXT').val('Gratuidade sem fidelidade, referente a alteração de plano');
			} else {
				$('#ADESAO_TXT').val('');
			}
		});
		$('#ADESAO_AUTO').change(function(){
			const gratuidadeSemFidelidade = $(this).val() === '0;0;0;1;0;Gratuidade sem fidelidade, referente a alteração de plano'
			$('#sem_fidelidade').attr('readonly', gratuidadeSemFidelidade);
			if (gratuidadeSemFidelidade) {
				$('#sem_fidelidade').attr('checked', true);
			}
		});
	});

    </script>
    
	<form id="formulario" action="Contratos/<%=CONTRATO%>" name="form" method="get" style="margin-top:7px;">
	<input type="hidden" value="<%=COD_COBRANCA%>" name="cobranca">
	<input type="hidden" value="<%=COD_CLIENTE%>" name="cod_cliente">
	<input type="hidden" value="1" name="imp">
   	<table width="650" align="center" style="border:1px solid #666666" border="0" cellspacing="1" cellpadding="1">
	<tr>
	  <td height="21" colspan="4" align="center" valign="middle" bgcolor="#F2F2F2"><strong><font size="2" face="Verdana">Imprimir Contrato</font></strong></td>
	</tr>


	<%
	IF (NOT Condominios.EOF) Then

		%>
		<tr> 
		<td height="21" valign="middle" bgcolor="#F2F2F2"> <p align="justify"><font size="2" face="Verdana">&nbsp;Condomínio</font></p></td>
		<td colspan="3" bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font> 
		  <select name="condominio"  class="campo-condomino" id="condominio"   onChange="location='?cobranca=<%=COD_COBRANCA%>&cod_cliente=<%=COD_CLIENTE%>&servico='+ document.form.servico.value +'&condominio='+ this.value" style="width:400px;">
	  		  <option selected="selected" value=""><- Selecione Condomínio -></option>
			  <%
			  TIP2_SERVICO = ""
			  While (NOT Condominios.EOF)
				IF cstr(request("condominio"))=cstr(Condominios("COD_CONDOMINIO")) Then TIP2_SERVICO = Condominios("TIP_SERVICO")
				%>
				<option <% IF cstr(request("condominio"))=cstr(Condominios("COD_CONDOMINIO")) Then Response.Write "selected" %> 
				value="<%=Condominios("COD_CONDOMINIO")%>"><%=Condominios("NOM_CONDOMINIO")%></option>
				<%
				Condominios.MoveNext()
			  Wend
			  Condominios.Close()
			  %>
		  </select>
		  &nbsp;</td>
		</tr>
		<%
	
	End IF	
	
	%>
    

	<tr>
	<td width="161" height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Servi&ccedil;o</font></td>
	<td width="456" colspan="3" bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font>
	<select name="servico"  class="Campo" id="servico"  onChange="location='?cobranca=<%=COD_COBRANCA%>&cod_cliente=<%=COD_CLIENTE%>&condominio='+ document.form.condominio.value +'&servico='+ this.value" style="width:400px;">
	  <option selected="selected" value=""><- Selecione Serviço -></option>
		<%
		While (NOT Servicos.EOF)
			%>
			<option <% IF cstr(COD_SERVICO)=cstr(Servicos("COD_SERVICO")) Then RESPONSE.Write "selected" %>
			value="<%=Servicos("COD_SERVICO")%>"><%=Servicos("NOME")%></option>
			<%
			Servicos.MoveNext()
		Wend
		Servicos.Close()
		
		%>
	</select>
	&nbsp;</td>
	</tr>
	
  




		<%



		If COD_SERVICO<>"" Then
			
		
			' Lista Comissionados do Servi&ccedil;o Selecionado
			Set COMISSIONADOS = Server.CreateObject("ADODB.Recordset")
			COMISSIONADOS.ActiveConnection = MM_Conexao_STRING
			COMISSIONADOS.Source = "SELECT NOM_COMISSIONADO, COD_SC  FROM SERVICOS_COMISSIONADOS S, COMISSIONADOS C "_
							     & "WHERE C.COD_COMISSIONADO=S.COD_COMISSIONADO AND S.COD_SERVICO="& COD_SERVICO &" AND (C.COD_PROVEDOR="& Session("cod_provedor") &") "_
								 & "AND APENAS_VENDA_SC=1  ORDER BY NOM_COMISSIONADO DESC"
			COMISSIONADOS.Open()
			
			'response.Write COMISSIONADOS.Source
			
			%>
			<tr>
			<td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Vendedor</font></td>
			<td colspan="3" bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font>
			<select name="vendedor"  class="Campo" id="vendedor" >
			<%
			While (NOT COMISSIONADOS.EOF)
				COD_SC = ORD_DIREITA(COMISSIONADOS("COD_SC"),4,"0")
				%>
				<option value="<%=COD_SC & COMISSIONADOS("NOM_COMISSIONADO")%>"><%=COMISSIONADOS("NOM_COMISSIONADO")%></option>
				<%
				COMISSIONADOS.MoveNext()
			Wend
			COMISSIONADOS.Close()
			
			%>
			<option value="0000ENGEPLUS">ENGEPLUS</option>
			</select>
			&nbsp;</td>
			</tr>
			<%
			
		End if
		
		
		
		%>

		<tr>
		  <td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Mensalidade</font></td>
		  <td bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font>
		  
		  <% IF VAL_SERVICO>0 Then %>
		  	<font color="#000000" face="Verdana" size="2"><%=FORMATCURRENCY(VAL_SERVICO)%>
            <% 
				IF PROMOCAO=1 THEN 
					Response.Write " / Promoção <b>"& FORMATCURRENCY(VALOR_PROMO) &"</b> por "& MESES_PROMO &" mensalidades."
					if session("nivel_usuario")>=30 Then Response.Write "<br><input type='checkbox' name='NaoPromocao' value=1 style=''> Anular Promoção"
				End If
			%>
            </font>
		  <% ELSE %>
		  	<input type="text" maxlength="7" size="7" name="VALOR" class="Campo" value="0" >
		  <% END IF %>		  
		  </td>
		  <td bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;M - Desconto </font></td>
		  <td bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font> <font color="#000000">
          <input type="text" maxlength="7" size="7" name="DESCONTO" class="Campo" value="<%=DESCONTO%>" <% IF Session("nivel_usuario")<50 and NOT CORPORATIVO and lcase(session("usuario"))<>"aline.dias" and lcase(session("usuario"))<>"vanuza" Then Response.Write "readonly" %>>
          </font></td>
		</tr>

		
		<%
		
	' RESIDENCIAL 

	IF instr(DES_SERVICO,"RESIDENCIAL") and instr(DES_SERVICO,"SMART PLUS") Then	  
	
	%>
			<tr>
			<td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Ades&atilde;o</font></td>
			<td colspan="3" bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font> <font color="#000000">
			<select id="ADESAO_AUTO" name="ADESAO_AUTO" class="Campo">
			<% 
			'  instalação ; desconto ; parcelado ; parcelas ; gratuidade ; texto 
			%>
 			<option value="<%=499+0+89%> ;499;0;1;0; R$ 89,00 (à vista)"> R$ 0,00 instalação + R$ 89,00 Taxa Configuração </option>
 			<option value="<%=499+45+89%>;499;0;1;0; R$ 134,00 "> R$ 45,00 transferência + R$ 89,00 Taxa Configuração</option> 
 			<option value="<%=499+90+89%>;499;0;1;0; R$ 179,00 "> R$ 90,00 transferência + R$ 89,00 Taxa Configuração</option>
 			<option value="<%=499+45+89%>;499;45;1;0;R$ 179,00 (1 + 1)"> R$ 90,00 Transf + R$ 89,00 Taxa Configuração (1 + 1)</option>

 			<% IF Session("usuario")="vanuza" or Session("usuario")="aline.dias" THEN %>
 			<option value="0;0;0;1;0;Gratuidade sem fidelidade, referente a alteração de plano">Gratuidade sem fidelidade (alteração de plano)</option>
 			<% END IF %>	

 			</select>
 			<br>
 			<input id="anteciparcontrato" type="checkbox"  name="contrato_comodato" onclick="" value="1"> 
 			<font  id="anteciparlabel" size="2" face="Verdana">&nbsp; Cobrança Antecipada</font>				  
			</font>
			</td>
			</tr>		
	<%

	ELSE IF instr(DES_SERVICO,"RESIDENCIAL") and instr(DES_SERVICO,"BANDA LARGA") Then
    
	%>				 
	    	<tr>
			<td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Ades&atilde;o</font></td>
			<td colspan="3" bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font> <font color="#000000">
			<select id="ADESAO_AUTO" name="ADESAO_AUTO" class="Campo">
			<% 
			'  instalação ; desconto ; parcelado ; parcelas ; gratuidade ; texto 
			%>
			<option value="<%=499+0%> ;499;0;1;0; R$ 89,00 (à vista)"> R$ 0,00 instalação </option>
 			<option value="<%=499+45%>;499;0;1;0; R$ 134,00 "> R$ 45,00 transferência </option> 
 			<option value="<%=499+90%>;499;0;1;0; R$ 179,00 "> R$ 90,00 transferência </option>
 			<option value="<%=499+45%>;499;45;1;0;R$ 179,00 (1 + 1)"> R$ 90,00 Transferência (1 + 1)</option>
				  
			<% IF Session("usuario")="vanuza" or Session("usuario")="aline.dias" THEN %>				  
			<option value="0;0;0;1;0;Gratuidade sem fidelidade, referente a alteração de plano">Gratuidade sem fidelidade (alteração de plano)</option>
			<% END IF %>				  
				  
			</select>				  
			</font>
			</td>
			</tr>	
	<%
	
	ELSE	

	'CONDOMINIO
	
	 IF instr(DES_SERVICO,"CONDOMINIO") and instr(DES_SERVICO,"SMART PLUS") Then	 
	
	%>
	   	    <tr>
			<td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Ades&atilde;o</font></td>
			<td colspan="3" bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font> <font color="#000000">
			<select id="ADESAO_AUTO" name="ADESAO_AUTO" class="Campo">
			<% 
			'  instalação ; desconto ; parcelado ; parcelas ; gratuidade ; texto 
			%>
 			<option value="<%=150+0+89%> ;150;0;1;0; R$ 89,00 (à vista)"> R$ 0,00 instalação + R$ 89,00 Taxa Configuração </option>
 			<option value="<%=150+45+89%>;150;0;1;0; R$ 134,00 "> R$ 45,00 transferência + R$ 89,00 Taxa Configuração</option> 
 			<option value="<%=150+90+89%>;150;0;1;0; R$ 179,00 "> R$ 90,00 transferência + R$ 89,00 Taxa Configuração</option>
 			<option value="<%=150+45+89%>;150;45;1;0;R$ 179,00 (1 + 1)"> R$ 90,00 Transf + R$ 89,00 Taxa Configuração (1 + 1)</option>

 			<% IF Session("usuario")="vanuza" or Session("usuario")="aline.dias" THEN %>
 			<option value="0;0;0;1;0;Gratuidade sem fidelidade, referente a alteração de plano">Gratuidade sem fidelidade (alteração de plano)</option>
 			<% END IF %>	

 			</select>
 			<br>
 			<input id="anteciparcontrato" type="checkbox"  name="contrato_comodato" onclick="" value="1"> 
 			<font  id="anteciparlabel" size="2" face="Verdana">&nbsp; Cobrança Antecipada</font>				  
			</font>
			</td>
			</tr>			
	<%

	ELSE IF instr(DES_SERVICO,"CONDOMINIO") and instr(DES_SERVICO,"BANDA LARGA") Then
	
	%>				 
		    <tr>
			<td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Ades&atilde;o</font></td>
			<td colspan="3" bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font> <font color="#000000">
			<select id="ADESAO_AUTO" name="ADESAO_AUTO" class="Campo">
			<% 
			'  instalação ; desconto ; parcelado ; parcelas ; gratuidade ; texto 
			%>
			<option value="<%=150+0%> ;150;0;1;0; R$ 89,00 (à vista)"> R$ 0,00 instalação </option>
 			<option value="<%=150+45%>;150;0;1;0; R$ 134,00 "> R$ 45,00 transferência </option> 
 			<option value="<%=150+90%>;150;0;1;0; R$ 179,00 "> R$ 90,00 transferência </option>
 			<option value="<%=150+45%>;150;45;1;0;R$ 179,00 (1 + 1)"> R$ 90,00 Transferência (1 + 1)</option>
				  
			<% IF Session("usuario")="vanuza" or Session("usuario")="aline.dias" THEN %>				  
			<option value="0;0;0;1;0;Gratuidade sem fidelidade, referente a alteração de plano">Gratuidade sem fidelidade (alteração de plano)</option>
			<% END IF %>				  
				  
			</select>				  
			</font>
			</td>
			</tr>			  
	<%
	 ELSE 
	 'OUTROS TIPOS 
	%>	

				<% IF Session("usuario")="vanuza" THEN %>
				<tr>
				<% ELSE %>
				<tr style="display:none">
				<% END IF %>
					<td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Ades&atilde;o</font></td>
					<td width="130" bgcolor="#F2F2F2" colspan="3">
						<input type="hidden" id="ADESAO_TXT" name="ADESAO_TXT" value="">
						<input type="checkbox" id="GRATUIDADE_SEM_FIDELIDADE" name="GRATUIDADE_SEM_FIDELIDADE" onclick="onClickGratuidadeSemFidelidade()" value="1">
						<font size="2" face="Verdana">&nbsp;Gratuidade sem fidelidade (alteração de plano)</font>
					</td>
				</tr>
				<tr>
				  <td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Ades&atilde;o - Entrada</font></td>
				  <td width="130" bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font> <font color="#000000">
				  <input type="text" maxlength="8" size="7" id="ENTRADA" name="ENTRADA" class="Campo" value="<%=formatnumber(INS_SERVICO)%>" <% 'IF Session("nivel_usuario")<20 Then Response.Write "readonly" %>>
				  <input type="hidden" id="ADESAO_MINIMA" name="ADESAO_MINIMA" value="<%=formatnumber(INS_SERVICO)%>">
				  </font></td>
				  <td width="100" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;A - Desconto </font></td>
				  <td bgcolor="#F2F2F2">
				  
				  <font color="#000000" size="2" face="Verdana">:</font> <font color="#000000">
				  <input name="INST_DESCONTO" type="text" class="Campo" id="INST_DESCONTO" value="0" size="7" maxlength="6" <% 'IF Session("nivel_usuario")<50 and NOT CORPORATIVO Then Response.Write "readonly" %>>
				  </font>
					<% 
					'END IF

				  %>           
				  
				  </td>
				</tr>

				<tr>
				  <td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Ades&atilde;o - Parcelado</font></td>
				  <td bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font> <font color="#000000">
						  <input type="text" maxlength="8" size="7" id="VALOR_PARCELADO" name="VALOR_PARCELADO" class="Campo" value="0" <% 'IF Session("nivel_usuario")<20 Then Response.Write "readonly" %>>
				  </font></td>
				  <td bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Condi&ccedil;&atilde;o</font></td>
				  <td bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font><font color="#000000">
							<select id="parcelas" name="parcelas" class="Campo">
							  <option value="1">1X</option>
							<% If Session("nivel_usuario")>=11 or INS_SERVICO>200 or lcase(session("usuario"))="rubia" or lcase(session("usuario"))="julia.candiotto" or lcase(session("usuario"))="marcus" then %>
							  <option value="2">2X</option>
							  <option value="3">3X</option>
							  <option value="4">4X</option>
							<% End If %>
							<% If Session("nivel_usuario")>=50 and INS_SERVICO>120 then %>
							  <option value="12">12X</option>
							<% End If %>
							<% If Session("nivel_usuario")>=50 or lcase(session("usuario"))="rubia" or lcase(session("usuario"))="julia.candiotto" or lcase(session("usuario"))="marcus" then %>
							<option value="24">24X</option>
							<option value="36">36X</option>
							<% End If %>
							</select>
						</font></td>
				</tr>
			<%			
	
	' END IF DAS CONDICÕES RESIDENCIAL E CONDOMINIO
	END IF
	End IF 
	END iF
	END iF		

	%>  
	  
	  
    	<tr>
    	  <td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Dura&ccedil;&atilde;o</font></td>
    	  <td bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font> <font color="#000000" face="Verdana" size="2">
          
		  <% 

			IF instr(DES_SERVICO,"EMPRESARIAL") AND instr(TIP_SERVICO,"fbr")=0 Then
				selected12 = "selected"
				selected24 = ""
			ElseIF instr(DES_SERVICO,"RESIDENCIAL") and Session("cod_provedor")=1 Then
				selected12 = "selected"
				selected24 = ""
			Else
				selected12 = "selected"
				selected24 = ""
			End if


		  IF Session("nivel_usuario")>=45 Then 
		   	%>
		  	<select name="DURACAO" class="Campo" style="width:100px">
			<% 
						
			%>
			<% 
			
			if session("nivel_usuario")=100 Then 
				%><option value="2 (dois)">2 meses</option>
				<option value="3 (tres)">3 meses</option>
                  <option value="4 (quatro)">4 meses</option>
                  <option value="5 (cinco)">5 meses</option>
				  <option value="6 (seis)">6 meses</option><% 
			end if 
			
			%>
			
			

			<option value="12 (doze)" <%=selected12%>>12 meses</option>
			<option value="24 (vinte e quatro)" <%=selected24%>>24 meses</option>
			<option value="36 (trinta e seis)" >36 meses</option>            
            
             <% 
			
			if session("nivel_usuario")>=40 Then 
				%><option value="0">Fidelidade atual - se tiver</option><% 
			end if 
			
			%>   
            </select>
		  <%
		  
		  ElseIf lcase(session("usuario"))="rubia" or lcase(session("usuario"))="julia.candiotto" or lcase(session("usuario"))="marcus" Then 
		  	%>
			<select name="DURACAO" class="Campo" style="width:100px">
			<option value="3 (tres)">3 meses</option>
			<option value="6 (seis)">6 meses</option>
			<option value="12 (doze)" <%=selected12%>>12 meses</option>
			<option value="16 (dezesseis)">16 meses</option>
			<option value="18 (dezoito)">18 meses</option>
			<option value="24 (vinte e quatro)" <%=selected24%>>24 meses</option>
			<option value="36 (trinta e seis)" <%=selected36%>>36 meses</option>            
            </select>
            <%

		  
		  ElseIF selected24<>"" Then

			%><input type="hidden" name="DURACAO" value="24 (vinte e quatro)">24 meses<%

		  ElseIf session("nivel_usuario")>=30 Then 
		  	%>
			<select name="DURACAO" class="Campo" style="width:100px">
			<option value="6 (seis)">6 meses</option>
			<option value="12 (doze)" <%=selected12%>>12 meses</option>
            </select>
            <%
			
		  ElseIf session("grupo")="qualidade" Then 
		  	%>
			<select name="DURACAO" class="Campo" style="width:100px">
			<option value="1 (um)">1 mês</option>
			<option value="6 (seis)">6 meses</option>
			<option value="12 (doze)" <%=selected12%>>12 meses</option>
            </select>
            <%	
			
		  Else

		  	%><input type="hidden" name="DURACAO" value="12 (doze)">12 meses<%

		  End If
		  
		  %>
          </font>
		  </td>
          
		  <td colspan="2" bgcolor="#F2F2F2">
		  <table border=0 cellpadding=0 cellspacing=0><tr><td><input type="checkbox" id="sem_fidelidade" name="sem_fidelidade" value="1"></td>
		  <td><font color="#000000" face="Verdana" size="2">&nbsp;SEM FIDELIDADE</font></td></tr></table>
          </td>
		  
   	  </tr>
      
    	<tr>
    	  <td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Gratuidade</font></td>
    	  <td bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font>
   	      <select name="GRATUIDADE" class="Campo" style="width:100px">
            <option value="0">Nenhuma</option>
            
			<% IF (instr(DES_SERVICO,"EMPRESARIAL") or instr(DES_SERVICO,"INTERLAN") or session("nivel_usuario")>=11) AND (NOT Verificar_PlanoFaturadoPassado("0"&COD_SERVICO,COD_CLIENTE,DATE()) or session("nivel_usuario")>=20) Then %>
                <option value="7">7 dias</option>
                <option value="15">15 dias</option>
                <option value="20">20 dias</option>                        
				<option value="30">30 dias</option>
				<% 'if session("nivel_usuario")>=20 then %>
					<option value="45">45 dias</option>
					<option value="60">60 dias</option>
					<option value="90">90 dias</option>
				<% 'end if %>
            <% End If %>
            
          </select></td>
    	  <td bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Abordagem</font></td>
    	  <td bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font>
            <select name="ABORDAGEM" class="Campo" style="width:100px">
              <option value="SIMPLES">SIMPLES</option>
              <option value="DUPLA">DUPLA</option>
          </select></td>
		</tr>
		
		
		
		
		<% IF (instr(DES_SERVICO,"INTERLAN")) or (instr(DES_SERVICO,"LAN TO LAN")) Then %>

			<tr>
			  <td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Prazo Instala&ccedil;&atilde;o</font></td>
			  <td bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font>
			  <select name="INST_PRAZO" class="Campo" style="width:100px">
					<option value="10">10 dias uteis</option>
					<option value="15">15 dias uteis</option>
					<option value="20">20 dias uteis</option>                        
					<option value="30">30 dias uteis</option>                                  
					<option value="40">40 dias uteis</option>                        
					<option value="60">60 dias uteis</option>                        
					<option value="90">90 dias uteis</option>                        
					<option value="120">120 dias uteis</option>                        				
					</select></td>
			  <td bgcolor="#F2F2F2"></td>
			  <td bgcolor="#F2F2F2"></td>
			</tr>	
			<tr>
			  <td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Pontos </font></td>
			  <td height="40" colspan="3" bgcolor="#F2F2F2">
			  
			  <% 
			  For i=1 to 2
				%>
				<font color="#000000" size="1" face="Verdana">
				<% if i<10 then response.write "0" %>
				<%=i%>.
				<input type="text" maxlength="100" size="30" name="PONTO<%=i%>" class="Campo" value="">
				<input type="text" maxlength="4" size="4" name="VELOCIDADE<%=i%>" class="Campo" value=""> Mbps
				<select name="TECNOLOGIA<%=i%>" class="Campo" style="width:70px">
				  <option value="FIBRA ÓPTICA">FIBRA</option>
				  <option value="RÁDIO FREQ.">RÁDIO</option>
				</select>
				</font>
				<br>
				<%
			  Next
			  %>
			  </td>
			</tr>
			
			
			
			

		<% ELSEIF (instr(DES_SERVICO,"FIBRA APAGADA")) Then %>

			<tr>
			  <td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Prazo Instala&ccedil;&atilde;o</font></td>
			  <td bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font>
			  <select name="INST_PRAZO" class="Campo" style="width:100px">
					<option value="10">10 dias uteis</option>
					<option value="15">15 dias uteis</option>
					<option value="20">20 dias uteis</option>                        
					<option value="30">30 dias uteis</option>                                  
					<option value="40">40 dias uteis</option>                        
					<option value="60">60 dias uteis</option>                        
					<option value="90">90 dias uteis</option>                        
					<option value="120">120 dias uteis</option>                        				
					</select></td>
			  <td bgcolor="#F2F2F2"> </td>
			  <td bgcolor="#F2F2F2"> </td>
			</tr>	
			<tr>
			  <td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Trechos </font></td>
			  <td height="40" colspan="3" bgcolor="#F2F2F2">
			  
			  <% 
			  For i=1 to 3
				%>
				<font color="#000000" size="1" face="Verdana">
				<%=i%>.
				<textarea rows="3" cols="30" name="TRECHO<%=i%>" class="Campo" value=""></textarea>
				<select name="FIBRAS<%=i%>" class="Campo" style="width:80px">
				  <option value="1">1 Fibra</option><option value="2">2 Fibras</option><option value="3">3 Fibras</option><option value="4">4 Fibras</option>
				</select>
				<input type="text" maxlength="4" size="4" name="DISTANCIA<%=i%>" class="Campo" value=""> Km
				</font>
				<br>
				<%
			  Next
			  %>
			  </td>
			</tr>

        <% End If %>
		

			  
		  <tr>
			  <td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Equipamentos</font></td>
			  <td colspan="3" bgcolor="#F2F2F2" style="padding-left:8px;">
              
              <div style="height:100px; overflow:auto;"> 
			  <table width="100%" border="0" cellspacing="1" cellpadding="2" class="produto">
			  <% 
				While NOT PRODUTOS.EOF
					%>
					<tr>
					  <td width="25" align="center"><input type="checkbox" name="CODIGOS_EQUIPAMENTOS" value="<%=PRODUTOS("COD_PRODUTO")%>"></td>
					  <td ><%=PRODUTOS("NOM_PRODUTO")%></td>
					  <td width="100" align="center"><select class="Campo" name="EquipamentoDisponibilizado_<%=PRODUTOS("COD_PRODUTO")%>">
									  <option value="COMODATO" selected>Comodato</option><option value="LOCAÇÃO">Locação</option>
									  </select>
					</tr>                
					<%
					PRODUTOS.MoveNext
				Wend
			  %>
			  </table>          
              </div>
			  
			  
			  </td>
		  </tr>

		  

			<tr>
			  <td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Observa&ccedil;&otilde;es </font></td>
			  <td height="40" colspan="3" bgcolor="#F2F2F2">			  
				<font color="#000000" size="1" face="Verdana">
				<textarea class="Campo" rows="3" cols="60" name="DADOS_ADICIONAIS"></textarea>
				</font>
				<br>
			  </td>
			</tr>
		
		  
		  
		  
		  
		  
			<tr>
				<td height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Como conheceu a Engeplus?</font></td>
				<td colspan="3" bgcolor="#F2F2F2">			  
					<select name="COMO_CONHECEU" class="Campo" id="como-conheceu" style="width:400px" required>
						<option value=""><- Selecione -></option>
						<option value="Venda Ativa">VENDA ATIVA</option>
						<option value="Indicação">INDICAÇÃO</option>
						<option value="Site Engeplus">SITE ENGEPLUS</option>
						<option value="Panfleto">PANFLETO</option>
						<option value="Outdoor">OUTDOOR</option>
						<option value="Google">GOOGLE</option>
						<option value="Facebook">FACEBOOK</option>
						<option value="Instagram">INSTAGRAM</option>
						<option value="Email Marketing">EMAIL MARKETING</option>
						<option value="Rádio">RÁDIO</option>
						<option value="Não Sabe">NÃO SABE / JÁ É CLIENTE</option>
					</select>
				</td>
			</tr>

      
    	<tr>
    	  <td height="21" valign="middle" bgcolor="#F2F2F2">&nbsp;</td>
    	  <td height="40" colspan="3" bgcolor="#F2F2F2"> &nbsp;
    	    <% 
			If COD_SERVICO<>"" Then 
				if session("nivel_usuario")>40 Then 
					%><!--<input type="button" value="ImprimirX" onClick="adesao()">--><input type="submit" name="Submit" value="Imprimir" id="btn-adesao" disabled="disabled"><%
				else 
					%><input type="button" value="Imprimir" onClick="adesao()" id="btn-adesao" disabled="disabled"><%
				end if
			eND iF
			%>				</td>
  	  </tr>
	</table>
	

</form>

<form action="Contratos/<%=CONTRATO%>" name="form2" method="get" ><BR>
	<input type="hidden" value="<%=COD_COBRANCA%>" name="cobranca">
	<input type="hidden" value="<%=COD_CLIENTE%>" name="cod_cliente">
	<input type="hidden" value="1" name="imp">

	<table width="650" align="center" style="border:1px solid #666666" border="0" cellspacing="1" cellpadding="1">
      <tr>
        <td height="21" colspan="2" align="center" valign="middle" bgcolor="#F2F2F2"><strong><font size="2" face="Verdana">Imprimir Adendo </font></strong></td>
      </tr>
      <tr>
        <td width="161" height="21" valign="middle" bgcolor="#F2F2F2"><font size="2" face="Verdana">&nbsp;Tipo</font></td>
        <td width="456" bgcolor="#F2F2F2"><font color="#000000" size="2" face="Verdana">:</font>
            <select name="adendo"  class="Campo" id="adendo" >
				<!--
				<option value="Contratos/AdendoContratual_TrocaPlanoDedicado.asp">Adendo - Troca Plano Dedicado</option>			
				<option value="Contratos/AdendoContratual_TrocaEndereco.asp">Adendo - Troca de Endereço</option>			
				<option value="Contratos/AdendoContratual_ComodatoEquipamentos.asp">Adendo - Contrato Comodato</option>
				-->
			<% IF Session("nivel_usuario")>=20 or Session("nivel_usuario")=11 THEN %>
				<!--
				<option value="Contratos/TermoEncerramento.asp">Termo de Encerramento</option>
				-->
				<option value="Contratos/__TermoEncerramento.asp">Termo de Encerramento</option>
				<option value="Contratos/__TermoEncerramento-Radio.asp">Termo de Encerramento Rádio</option>
				<option value="Contratos/__TermoEncerramento-Radio-Transferencia.asp">Termo de Encerramento Rádio - Transferência</option>
				<option value="Contratos/Ferias.asp">Plano F&eacute;rias</option>
				<!--
				<option value="Contratos/TermoEncerramento_Telefonia.asp">Termo de Encerramento TELEFONIA</option>
				<option value="Contratos/TermoEncerramento_Telefonia2.asp">Termo de Encerramento TELEFONIA 2</option>
				-->
			<% END IF %>
			<% IF Session("nivel_usuario")>=100 THEN %>

			<% END IF %>

            </select>
          &nbsp;</td>
      </tr>
      <tr>
        <td height="21" valign="middle" bgcolor="#F2F2F2">&nbsp;</td>
        <td height="40" bgcolor="#F2F2F2"> &nbsp;
        <input type="button" name="Submit2" value="Prosseguir" onClick="document.form2.action=document.form2.adendo.value;document.form2.submit()"></td>
      </tr>
      
      <%



		If COD_SERVICO<>"" Then
			
		
			' Lista Comissionados do Servi&ccedil;o Selecionado
			Set COMISSIONADOS = Server.CreateObject("ADODB.Recordset")
			COMISSIONADOS.ActiveConnection = MM_Conexao_STRING
			COMISSIONADOS.Source = "SELECT NOM_COMISSIONADO  FROM SERVICOS_COMISSIONADOS S, COMISSIONADOS C "_
							     & "WHERE C.COD_COMISSIONADO=S.COD_COMISSIONADO AND S.COD_SERVICO="& COD_SERVICO &" AND (C.COD_PROVEDOR="& Session("cod_provedor") &") "_
								 & "AND APENAS_VENDA_SC=1  ORDER BY NOM_COMISSIONADO"
			COMISSIONADOS.Open()
			
			'response.Write COMISSIONADOS.Source
			
			%>
      
      <%
			
		End if
		
		
		
		%>
    </table>
	</form>
	
	


	

</body>
</html>

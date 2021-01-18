<link href="Style.css" rel="stylesheet" type="text/css">
<%'<!--#include file="Menu.asp"-->%>
<!--#include file="Connections/Conexao.asp" -->
<!--#include file="Func.asp"-->
<% 
'/////////////////////////////////////////////////////////////////////////////////////////'
'*********  VERIFICA PERMISSAO DE ACESSO A PAGINA ****************************************'
'/////////////////////////////////////////////////////////////////////////////////////////'

IF Session("nivel_usuario")<100 THEN  '****  MASTER  ****'
	RESPONSE.WRITE "<script language='javascript'>location.assign('Logoff.asp')</script>"
	RESPONSE.End() 'TERMINA O CODIGO DESTA PAGINA
END IF

'*****************************************************************************************'
%>

<%
TIPO = TRIM(REQUEST("TIPO"))
VENCIMENTO = TRIM(REQUEST("VENCIMENTO"))

IF  not ISDATE(VENCIMENTO) THEN  
	RESPONSE.WRITE "<script language='javascript'>location.assign('Financeiro_Remeter_Adesao.asp?erro=Vencimento Inv�lido!')</script>"
	RESPONSE.End() 'TERMINA O CODIGO DESTA PAGINA
END IF

%>


<% 
'/////////////////////////////////////////////////////////////////////////////////////////'
'*********  REQUISITA OS TIPOS DE COBRANCA ***********************************************'
'/////////////////////////////////////////////////////////////////////////////////////////'
%>
<%
Dim TIPOS_COBRANCA__PROVEDOR
TIPOS_COBRANCA__PROVEDOR = "0"
If (Session("cod_provedor") <> "") Then 
  TIPOS_COBRANCA__PROVEDOR = Session("cod_provedor")
End If
%>
<%
Dim TIPOS_COBRANCA__CODIGO
TIPOS_COBRANCA__CODIGO = "0"
If (TIPO <> "") Then 
  TIPOS_COBRANCA__CODIGO = TIPO
End If
%>
<%
Dim TIPOS_COBRANCA
Dim TIPOS_COBRANCA_numRows

Set TIPOS_COBRANCA = Server.CreateObject("ADODB.Recordset")
TIPOS_COBRANCA.ActiveConnection = MM_Conexao_STRING
TIPOS_COBRANCA.Source = "SELECT COD_TIPO_COBRANCA, ARQ_TIPO_COBRANCA, DES_TIPO_COBRANCA +' - '+ BNC_TIPO_COBRANCA AS TIPO  FROM dbo.TIPOS_COBRANCA  WHERE DES_TIPO_COBRANCA<>'TELESC' AND COD_PROVEDOR=" + Replace(TIPOS_COBRANCA__PROVEDOR, "'", "''") + " AND COD_TIPO_COBRANCA=" + Replace(TIPOS_COBRANCA__CODIGO, "'", "''") + " ORDER BY TIPO"
TIPOS_COBRANCA.CursorType = 0
TIPOS_COBRANCA.CursorLocation = 2
TIPOS_COBRANCA.LockType = 1
TIPOS_COBRANCA.Open()
TIPOS_COBRANCA_numRows = 0
%>
<%
IF NOT TIPOS_COBRANCA.EOF THEN NOME_TIPO = TIPOS_COBRANCA("TIPO")
'*****************************************************************************************'
%>

<% 
'/////////////////////////////////////////////////////////////////////////////////////////'
'*********  REQUISITA NOVOS CLIENTES *****************************************************'
'/////////////////////////////////////////////////////////////////////////////////////////'
%>
<%
Set NOVOS = Server.CreateObject("ADODB.Recordset")
NOVOS.ActiveConnection = MM_Conexao_STRING
'Mudanca Taliha
NOVOS.Source = "SELECT * FROM view_FINANCEIRO_ADESOES WHERE (ADESAO_ENTRADA>0 or ADESAO_VALOR_PARCELADO>0) AND COD_PROVEDOR="& Session("cod_provedor") &" ORDER BY NOME"
'NOVOS.Source = "SELECT * FROM view_FINANCEIRO_ADESOES WHERE ADESAO_ENTRADA>0 AND COD_PROVEDOR="& Session("cod_provedor")
NOVOS.Open()
%>


<% 
'/////////////////////////////////////////////////////////////////////////////////////////'
'*********  REQUISITA NOVOS CLIENTES SEM ADES�O - VENDAS *********************************'
'/////////////////////////////////////////////////////////////////////////////////////////'
%>
<%
'Set NOVOS2 = Server.CreateObject("ADODB.Recordset")
'NOVOS2.ActiveConnection = MM_Conexao_STRING
'NOVOS2.Source = "SELECT * FROM view_FINANCEIRO_ADESOES WHERE ADESAO_ENTRADA=0 AND COD_SC=0 AND COD_PROVEDOR="& Session("cod_provedor") 
'NOVOS2.Open()
%>




 
<title>SIAF | Fin - Enviar Cobran�a Ades�o | <%=Session("provedor")%></title>
<style type="text/css">
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
      <td valign="top">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
               <td height="16" bgcolor="#EFEFEF"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><font color="#000000" size="1">&nbsp;FINANCEIRO 
      <font color="#000066">&gt;&gt;</font> REMETER ARQUIVO <font color="#000066">&gt;&gt;</font> ADES&Atilde;O <font color="#000066">&gt;&gt;</font> SELECIONAR</font></font></td>
            </tr>
        </table>
         
      </td>
      <td width="72" align="right"><a href="javascript: history.go(-1)"><img src="imagens/voltar.gif" width="69" height="17" border="0"></a></td>
  </tr>
</table>
  
<form name="Form" action="Financeiro_Remeter_Adesao_Visualizar.asp?<%=Request.QueryString%>" method="POST">
   <table width="550" border="0" align="center" cellpadding="0" cellspacing="0">
     <tr>
       <td><img src="imagens/abas/remeter_arquivo.gif" width="187" height="15" border="0"></td>
     </tr>
   </table>
   <table width="550" border="0" align="center" cellpadding="1" cellspacing="1" style="border:1px solid #666666 ">
    <tr> 
      <td width="177" valign="middle" bgcolor="#99FF99"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Tipo 
        de Cobran&ccedil;a</font></td>
      <td width="364" height="10" bgcolor="#99FF99" ><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
&nbsp;<% =("&nbsp;"& NOME_TIPO)%>
        </font></td>
    </tr>
    <tr>
      <td valign="middle" bgcolor="#99FF99"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Data 
        de Vencimento </font></td>
      <td height="4" bgcolor="#99FF99" >&nbsp;<font size="2" face="Verdana, Arial, Helvetica, sans-serif"> <%= VENCIMENTO %></font></td>
    </tr>
  </table>
   <br>
   <table width="830" border="0" align="center" cellpadding="0" cellspacing="0">
     <tr>
       <td><img src="imagens/abas/novos_clientes.gif" width="130" height="15" border="0"></td>
     </tr>
  </table>
   <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" style="border:1px solid #666666 ">
      <tr> 
         <td width="1%" valign="middle" bgcolor="#666666"><font color="#FFFFFF">&nbsp;</font></td>
         <td valign="middle" bgcolor="#666666"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Nome</font></strong></td>
         <td bgcolor="#666666" ><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Servi&ccedil;o</font></strong></td>
         <td width="7%" align="center" bgcolor="#666666" ><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Entrada</font><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; </font></strong></td>
         <td width="7%" align="center" bgcolor="#666666" ><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Parcelado</font></strong></td>
         <td width="7%" height="10" align="center" bgcolor="#666666" ><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Parcelas</font></strong></td>
         <td width="12%" bgcolor="#666666" ><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Vendedor</font></strong></td>
		 <td width="7%" align="center" bgcolor="#666666"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Antecipar</font></strong></td>
		 <td width="7%" align="center" bgcolor="#666666"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Faturar</font></strong></td>
     </tr>

	<% 

	' Abre Objeto de Comissionados
	Set COMISSIONADOS = Server.CreateObject("ADODB.Recordset")
	COMISSIONADOS.ActiveConnection = MM_Conexao_STRING

   	DO WHILE NOT NOVOS.EOF 

		ADESAO_PARCELAS = NOVOS("ADESAO_PARCELAS")
 		COD_SC			= NOVOS("COD_SC")

		sql_comissionado=""
		IF COD_SC>0 THEN sql_comissionado = "AND COD_SC="& COD_SC
		
		%>
      	<tr bgcolor="#EDEDED" onMouseOver="javascript:className='LinhaAtiva'" onMouseOut="javascript:className=''">         
        <td valign="middle">
		    <input name="GERA_<%=NOVOS("COD_COBRANCA")%>_<%=NOVOS("COD_SERVICO")%>" type="checkbox" value="1" checked>        </td>
        <td height="25" valign="middle"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<%=(NOVOS.Fields.Item("NOME").Value)%></font></td>
        <td ><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> &nbsp;<%=lcase(NOVOS.Fields.Item("SERVICO").Value)%></font></td>
        <td align="center" ><input name="ENTRADA_<%=NOVOS("COD_COBRANCA")%>_<%=NOVOS("COD_SERVICO")%>" type="text" class="Campo" id="ENTRADA_<%=NOVOS("COD_COBRANCA")%>_<%=NOVOS("COD_SERVICO")%>" onFocus="this.select()" value="<%=FormatNumber(NOVOS("ADESAO_ENTRADA"))%>" size="8"></td>
        <td align="center" ><input name="VALOR_PARCELADO_<%=NOVOS("COD_COBRANCA")%>_<%=NOVOS("COD_SERVICO")%>" type="text" class="Campo" id="VALOR_PARCELADO_<%=NOVOS("COD_COBRANCA")%>_<%=NOVOS("COD_SERVICO")%>" onFocus="this.select()" value="<%=FormatNumber(NOVOS("ADESAO_VALOR_PARCELADO"))%>" size="6"></td>
        <td height="10" align="center" ><select name="PARCELAS_<%=NOVOS("COD_COBRANCA")%>_<%=NOVOS("COD_SERVICO")%>" class="Campo9" id="PARCELAS_<%=NOVOS("COD_COBRANCA")%>_<%=NOVOS("COD_SERVICO")%>">
          <option value="1" <% IF ADESAO_PARCELAS=1 Then Response.Write "selected" %>>1x</option>
          <option value="2" <% IF ADESAO_PARCELAS=2 Then Response.Write "selected" %>>2x</option>
          <option value="3" <% IF ADESAO_PARCELAS=3 Then Response.Write "selected" %>>3x</option>
          <% if session("nivel_usuario")>=50 then %>
          <option value="12" <% IF ADESAO_PARCELAS=12 Then Response.Write "selected" %>>12x</option>
          <% end if %>
        </select></td>
    	<td style="padding-left:4px">
		<%
		
		
		' Lista Comissionados do Servi�o Selecionado
		COMISSIONADOS.Source = "SELECT NOM_COMISSIONADO, COD_SC  FROM SERVICOS_COMISSIONADOS S, COMISSIONADOS C WHERE C.COD_COMISSIONADO=S.COD_COMISSIONADO "_
							 & "AND S.COD_SERVICO="& NOVOS("COD_SERVICO") &" AND (C.COD_PROVEDOR="& Session("cod_provedor") &") "_
							 & "AND APENAS_VENDA_SC=1 "&sql_comissionado&" ORDER BY NOM_COMISSIONADO"
		COMISSIONADOS.Open()

		%>
		<select name="COMISSIONADO_<%=NOVOS("COD_COBRANCA")%>_<%=NOVOS("COD_SERVICO")%>" class="Campo9" style="width:110px;" id="COMISSIONADO_<%=NOVOS("COD_COBRANCA")%>_<%=NOVOS("COD_SERVICO")%>"
		 onChange="document.Form.GERA_<%=NOVOS("COD_COBRANCA")%>_<%=NOVOS("COD_SERVICO")%>.checked=true">
		 <% if COMISSIONADOS.EOF Then %><option value="0">ENGEPLUS</option><% End If %>
		 <% 
		 DO WHILE NOT COMISSIONADOS.EOF 
			%>
			<option value="<%=COMISSIONADOS("COD_SC") %>"><%=(COMISSIONADOS("NOM_COMISSIONADO"))%></option>
			<% 	
			COMISSIONADOS.MOVENEXT
		 LOOP 
		 COMISSIONADOS.Close()
		 %>
		</select>
  
   <script>

    function chkAntecipa(val) {  
    let antecipar  = document.getElementById(val).checked;    
       if (antecipar) {             
          document.getElementById("fatura_programada_"+val.substr(18)).checked = false;        
       }
    }
    
    function chkProgramada(val) {
    let programada = document.getElementById(val).checked;        
       if (programada) {             
          document.getElementById("fatura_antecipada_"+val.substr(18)).checked = false;        
       }  
    }      
   </script>		

		</td>
		<td align="center" ><input name="fatura_antecipada_<%=NOVOS("COD_COBRANCA")%>_<%=NOVOS("COD_SERVICO")%>" id="fatura_antecipada_<%=NOVOS("COD_COBRANCA")%>_<%=NOVOS("COD_SERVICO")%>" type="checkbox" value="1" onchange="chkAntecipa(this.id)"></td>
		<td align="center" ><input name="fatura_programada_<%=NOVOS("COD_COBRANCA")%>_<%=NOVOS("COD_SERVICO")%>" id="fatura_programada_<%=NOVOS("COD_COBRANCA")%>_<%=NOVOS("COD_SERVICO")%>" type="checkbox" value="1" onchange="chkProgramada(this.id)" checked></td>
    </tr>
		<%    
		NOVOS.MOVENEXT
    LOOP 
	%>
  </table>
  
<br> 
   
   
   
   <%
   
   	' Vendas - Associar comissionado sem ades�o
   	'
   	'% >
    '<table width="830" border="0" align="center" cellpadding="0" cellspacing="0">
    ' <tr>
    '   <td><img src="imagens/abas/vendas.gif"  border="0"></td>
    ' </tr>
    '</table>
   	'<table width="830" border="0" align="center" cellpadding="1" cellspacing="1" style="border:1px solid #666666 ">
    ' <tr>
    '   <td width="302" valign="middle" bgcolor="#666666"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Nome</font></strong></td>
    '   <td width="280" bgcolor="#666666" ><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Servi&ccedil;o</font></strong></td>
    '   <td width="236" bgcolor="#666666" ><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Vendedor</font></strong></td>
    ' </tr>
    ' <% 
	'
	' Abre Objeto de Comissionados
	'Set COMISSIONADOS = Server.CreateObject("ADODB.Recordset")
	'COMISSIONADOS.ActiveConnection = MM_Conexao_STRING
	'
   	'DO WHILE NOT NOVOS2.EOF 
 	'	% >
    ' <tr bgcolor="#EDEDED" onMouseOver="javascript:className='LinhaAtiva'" onMouseOut="javascript:className=''">
    '   <td width="302" height="25" valign="middle">
    '   <font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<%=(NOVOS2.Fields.Item("NOME").Value)% ></font></td>
    '   <td width="280" ><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> &nbsp;<%=lcase(NOVOS2.Fields.Item("SERVICO").Value)% ></font></td>
    '   <td width="236" align="left" > &nbsp;
    '     <%
	'	
	'	
	'	' Lista Comissionados do Servi&ccedil;o Selecionado
	'	COMISSIONADOS.Source = "SELECT NOM_COMISSIONADO, COD_SC  FROM SERVICOS_COMISSIONADOS S, COMISSIONADOS C WHERE C.COD_COMISSIONADO=S.COD_COMISSIONADO AND S.COD_SERVICO="& NOVOS2("COD_SERVICO") &" AND (C.COD_PROVEDOR="& Session("cod_provedor") &") AND APENAS_VENDA_SC=1  ORDER BY NOM_COMISSIONADO"
	'	COMISSIONADOS.Open()
	'
	'	% >
    '       <select name="COMISSIONADO_<%=NOVOS2("COD_COBRANCA")% >_<%=NOVOS2("COD_SERVICO")% >" class="Campo9" id="COMISSIONADO_<%=NOVOS2("COD_COBRANCA")% >_<%=NOVOS2("COD_SERVICO")% >">
    '         <% if COMISSIONADOS.EOF Then % ><option value="0">ENGEPLUS</option><% End If % >
	'		 <% 
	'	 DO WHILE NOT COMISSIONADOS.EOF 
	'		% >
    '         <option value="<%=COMISSIONADOS("COD_SC") % >"><%=(COMISSIONADOS("NOM_COMISSIONADO"))% ></option>
    '         <% 	
	'		COMISSIONADOS.MOVENEXT
	'	 LOOP 
	'	 COMISSIONADOS.Close()
	'	 % >
    '       </select>       </td>
    ' </tr>
    ' <%    
	'	NOVOS2.MOVENEXT
    'LOOP 
	'% >
   	'</table>
    '<br>
    %>
    
   <br>
<table width="200" border="0" align="center" cellpadding="0" cellspacing="0">
     <tr>
       <td ><img src="imagens/abas/opcoes_salvar.gif" width="114" height="15" border="0"></td>
     </tr>
  </table>
   <table width="200" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #666666">
     <tr>
       <td height="33" colspan="4" align="center" valign="middle" bgcolor="#F2F2F2">
         <input name="Cancelar2" type="button" value="Voltar" onClick="history.go(-1)">
&nbsp; <strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">
      <input name="Submit2" type="submit" value="Salvar">
    </font></strong></td>
     </tr>
   </table>
   
   
   
</form>

<br>
<%
TIPOS_COBRANCA.Close()
Set TIPOS_COBRANCA = Nothing
%>
<%
NOVOS.Close()
Set NOVOS = Nothing
%>
<%
'NOVOS2.Close()
'Set NOVOS2 = Nothing
%>
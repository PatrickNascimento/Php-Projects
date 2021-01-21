 '╔══════════════════════════════════════════════════════════════════════════════════════════════╗
 '║▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒ MONTA ARRAY PARA DOS DADOS PARA ENVIAR PARA ROTA VIA AJAX ▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒║
 '╚══════════════════════════════════════════════════════════════════════════════════════════════╝
    %>
      <script>
        cobranca_arr.push(<%=GERAR("COD_COBRANCA")%>);
        cod_servico_arr.push(<%=GERAR("COD_SERVICO")%>);
        valor_arr.push(<%=(GERAR.Fields.Item("VALOR").Value)%>);
        parcela_arr.push(<%=(GERAR.Fields.Item("PARCELAS").Value)%>);
        servicos_arr.push("<%=(GERAR.Fields.Item("SERVICO").Value)%>");  
        cod_cliente_arr.push(<%=(GERAR.Fields.Item("COD_CLIENTE").Value)%>);
        status_arr.push("<%=ChkStatus%>");
      </script>

			<% 
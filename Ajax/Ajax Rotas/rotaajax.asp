<link href="Style.css" rel="stylesheet" type="text/css">
<script src="http://code.jquery.com/jquery-1.9.1.js"></script>

<%
TIPO         = TRIM(REQUEST("TIPO"))
VENCIMENTO   = TRIM("'"&REQUEST("VENCIMENTO")&"'")
COUNT        = TRIM("'"&REQUEST("TOTALCOUNT")&"'")
COD_COBRANCA = TRIM("'"&REQUEST("COD_COBRANCA")&"'")
COD_SERVICO  = TRIM("'"&REQUEST("COD_SERVICO")&"'") 
COD_CLIENTE  = TRIM("'"&REQUEST("COD_CLIENTE")&"'") 
COD_CLIENTE  = TRIM("'"&REQUEST("COD_CLIENTE")&"'") 
PARCELAS     = TRIM("'"&REQUEST("PARCELAS")&"'") 
VALOR        = TRIM("'"&REQUEST("VALOR")&"'") 
STATUS       = TRIM("'"&REQUEST("STATUS")&"'")
SERVICOS     = TRIM("'"&REQUEST("SERVICO")&"'")
%>

<SCRIPT>
function SendRoute() {
var CB = <%=COD_COBRANCA%>;
var CS = <%=COD_SERVICO%>;
var CC = <%=COD_CLIENTE%>;
var PA = <%=PARCELAS%>;
var VA = <%=VALOR%>;
var SE = <%=SERVICOS%>
var ST = <%=STATUS%>
var TIPO = <%=TIPO%>
var VENCIMENTO = <%=VENCIMENTO%>

  for (i=0; i < <%=COUNT%>; i++) {
    var COD_COBRANCA = CB.toString().split(",")[i];    
    var COD_SERVICO = CS.toString().split(",")[i];
    var COD_CLIENTE = CC.toString().split(",")[i];
    var VALOR = PA.toString().split(",")[i];
    var PARCELA = VA.toString().split(",")[i];
    var sta = ST.toString().split(",")[i];
    var SERVICO = SE.toString().split(",")[i];        
        if (sta.trim() == "ANTECIPAR") {                       
          $.ajax({            
                    url: "http://localhost/rota/rotaantecipar.php",
                    type: "POST",
                    dataType: "json",
                    data: {"TIPO"      : TIPO,
                          "VENCIMENTO" : VENCIMENTO,
                        "COD_COBRANCA" : COD_COBRANCA,
                         "COD_SERVICO" : COD_SERVICO,
                         "COD_CLIENTE" : COD_CLIENTE,
                               "VALOR" : VALOR,
                             "PARCELA" : PARCELA,
                             "SERVICO" : SERVICO,
                         "ANTECIPACAO" : 1
                           },      
                    success: function(response) {
                          console.log(response);
                    }
                });
    } else {            
          $.ajax({
                    url: "http://localhost/rota/rotafaturar.php",
                   type: "POST",
                    dataType: "json",
                    data: {"TIPO"      : TIPO,
                          "VENCIMENTO" : VENCIMENTO,
                        "COD_COBRANCA" : COD_COBRANCA,
                         "COD_SERVICO" : COD_SERVICO,
                         "COD_CLIENTE" : COD_CLIENTE,
                               "VALOR" : VALOR,
                             "PARCELA" : PARCELA,
                             "SERVICO" : SERVICO                         
                           },      
                    success: function(response) {
                          console.log(response);
                    }
                });
    }
  }
}

</SCRIPT>  

<script>SendRoute()</script>

  

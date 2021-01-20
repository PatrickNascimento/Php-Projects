<?php
 header("Access-Control-Allow-Origin: *");

 $TIPO = $_POST['TIPO'];
 $VENCIMENTO = $_POST['VENCIMENTO'];
 $COD_COBRANCA = $_POST['COD_COBRANCA']; 
 $COD_SERVICO = $_POST['COD_SERVICO'];
 $COD_CLIENTE = $_POST['COD_CLIENTE'];
 $VALOR = $_POST['VALOR'];
 $PARCELA = $_POST['PARCELA'];
 $SERVICO = $_POST['SERVICO']; 

 echo json_encode($TIPO." | ".$VENCIMENTO." | ".$COD_COBRANCA." | ".$COD_SERVICO." | ".$COD_CLIENTE." | ".$VALOR." | ".$PARCELA." | ".$SERVICO); 
 

 
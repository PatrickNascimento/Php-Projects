<?php
include_once("config.php");


#define o encoding do cabeçalho para utf-8
@header('Content-Type: text/html; charset=utf-8');

#carrega o arquivo XML e retornando um Array
$xml = simplexml_load_file(ARQUIVO);

include "top.htm";

#para cada nó atribui à variavel (obj simplexml)
foreach($xml->{'user-manager'}->{'auth-config'}->{'user'} as $user)
{

include "item.htm";

}

include "roda.htm";

?>
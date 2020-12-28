<?php
#Noticia constructor Php Orientado Objetos

class noticia
{
    public $titulo;
    public $texto;

    function __construct($val_tit, $val_txt)
    {
    $this->titulo = $val_tit;    
    $this->texto = $val_txt;    
    }

    function exibenoticias()
    {
        echo $this->titulo;
        echo $this->texto;

    }
}

$noticia = new noticia('Titulo constructor','exemplo de aplicação do constructor');
$noticia->exibenoticias();
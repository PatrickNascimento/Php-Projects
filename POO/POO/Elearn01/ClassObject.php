<?php
#classe Notífica.php Php Orientado a Objetos



class noticias {

    public $titulo;
    public $texto;

    function setTitulo($valor){
        $this->titulo = $valor;
    }
    function setTexto($valor){
        $this->titulo = $valor;
    }

    function exibeNoticias(){        
        echo $this->titulo. '<br>' ;
        echo $this->texto;
    }
    
}

$noticia = new noticias();
$noticia->titulo = "Hello World";
$noticia->texto = "Orientação a Objetos com PHP";
$noticia->exibeNoticias();
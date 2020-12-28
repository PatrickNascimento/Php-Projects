<?php
include_once("ClassObject.php");

Class heranca extends noticias{

    public $image;

    function setImage($valor){
        $this->image = $valor;
    }

    function exibeNoticias()
    {
        echo $this->titulo ."<br>";
        echo $this->texto . "<br>";
        echo "<img src=\"". $this->image .">";
    }
}

$noticia = new heranca;

$noticia->titulo .= "Php";
$noticia->texto .= "Heranca";
$noticia->image = "php.jpg";

$noticia->exibeNoticias();


/** Php Orientado objeto */


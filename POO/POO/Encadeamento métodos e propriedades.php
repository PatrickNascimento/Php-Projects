<?php
class Car{
    public $tanque;
//adicionando litros de gasolina no tanque
public function fill($float) {
    $this->tanque +=$float;
    return $this;
}

//Subtrair galoes de gasolina do tanque
public function ride($float) {
    $km = $float;    

//Encadeando métodos e propriedades
$litros = $km/12;
$this->tanque -= $litros;
return $this;
}
}

$bmw = new Car();
// Adicionado 45 galoes da gasolina para rodar 500 km
//e pegando o numero de galoes do tanque
$tanque = $bmw -> fill(45) -> ride(500) -> tanque;
// Imprime o resultado na tela
echo "Numero de litros restantes no tanque é: " . $tanque . " litros.";




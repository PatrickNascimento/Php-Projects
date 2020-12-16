<?php
class Car // Classe
{
    public $comp; //Propriedade da Classe
    public $color = "beige";
    public $hasSunRoof = true;
    public function hello()
    {
        return "beep";
    }

}

$car1 = new Car(); //Instancia da Classe
$mercedes = new Car(); //Instancia da Classe
$mercedes->comp = 'MercedesBenz';
echo $mercedes->comp;
echo "<br>";
echo $car1->color; //getColor
echo "<br>";
echo $car1->color = 'green'; //setColor
echo "<br>";
echo $car1->hello();

?>


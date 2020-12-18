<?php 
Class User{
    public $primeiroNome;
    public $SobreNome;
    

public function hello()
{
                    return "hello " .$this->primeiroNome;
}
}

$user1 = new User();

$this->primeiroNome = "Patrick";


echo $user1->hello();
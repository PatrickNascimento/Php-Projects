<?php

include_once("class/Crud.php");

$crud = new Crud();

$id = $crud->escape_string($_GET['id']);

$result = $crud->delete($id, 'usuarios');

if ($result) {
	header("Location: index.php");
}

?>
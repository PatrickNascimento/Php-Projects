<?php

include_once("class/Crud.php");

$crud = new Crud();

$id = $crud->escape_string($_GET['id']);

$result = $crud->getData("SELECT * FROM usuarios WHERE id = $id");

foreach ($result as $res) {
	$nome = $res['nome'];
	$idade = $res['idade'];
	$email = $res['email'];
}
?>

<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8">
	<title>Atualizar usuário - <?=$nome;?></title>
	<link rel="stylesheet" href="assets/css/bootstrap.css">
</head>
<body class="container mt-4">

	<div class="row">
		<div class="col-lg-10">
			<h1 class="text-center">Atualizar o usuário</h1>
		</div>
		<div class="col-lg-2">
			<a href="index.php" class="btn btn-success">Lista de usuários</a>
		</div>
	</div>

	<form action="editaction.php" method="POST" class="form">
		<div class="form-group">
			<label for="nome">Nome:</label>
			<input type="text" name="nome" value="<?=$nome;?>" class="form-control">
		</div>
		<div class="form-group">
			<label for="idade">Idade:</label>
			<input type="text" name="idade" value="<?=$idade;?>" class="form-control">
		</div>
		<div class="form-group">
			<label for="email">E-mail:</label>
			<input type="text" name="email" value="<?=$email;?>" class="form-control">
		</div>
		<input type="hidden" name="id" value="<?=$id;?>">
		<input type="submit" name="Update" class="btn btn-info float-right" value="Salvar alteração">
	</form>
</body>
</html>
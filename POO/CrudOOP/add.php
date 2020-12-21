<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8">
	<title>Adicionar usuário</title>
	<link rel="stylesheet" href="assets/css/bootstrap.css">
</head>
<body class="container text-center mt-4">

<?php
include_once("class/Crud.php");
include_once("class/Validation.php");

$crud = new Crud();
$validation = new Validation();

if (isset($_POST['Submit'])) {
	$nome = $crud->escape_string($_POST['nome']);
	$idade = $crud->escape_string($_POST['idade']);
	$email = $crud->escape_string($_POST['email']);

	$msg = $validation->check_empty($_POST, array('nome', 'idade', 'email'));
	$check_age = $validation->is_age_valid($_POST['idade']);
	$check_email = $validation->is_email_valid($_POST['email']);

	if ($msg != NULL) {
		echo $msg;
		echo "<br/> <a href='javascript:self.history.back();' class='btn btn-primary'>Voltar</a>";
	} elseif (!$check_age) {
		echo "<span class='alert alert-danger'>Por favor, forneça a idade apropriada!</span>";
	} elseif (!$check_email) {
		echo "<span class='alert alert-danger'>Por favor, forneça o email apropriado!</span>";
	} else {
		$result = $crud->execute("INSERT INTO usuarios (nome, idade, email) VALUES ('$nome', '$idade', '$email')");
		echo "<span class='alert alert-success'>Usuário adicionado com sucesso!</span>";
		echo "<br/><a href='index.php' class='btn btn-primary mt-4'>Veja o resultado</a>";
	}
}

?>
	
</body>
</html>
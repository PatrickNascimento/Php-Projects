<?php 

include_once("class/Crud.php");
include_once("class/Validation.php");

$crud = new Crud();
$validation = new Validation();

if (isset($_POST['Update'])) {
	$id = $crud->escape_string($_POST['id']);

	$nome = $crud->escape_string($_POST['nome']);
	$idade = $crud->escape_string($_POST['idade']);
	$email = $crud->escape_string($_POST['email']);

	$msg = $validation->check_empty($_POST, array('nome', 'idade', 'email'));
	$check_age = $validation->is_age_valid($_POST['idade']);
	$check_email = $validation->is_email_valid($_POST['email']);

	if ($msg) {
		echo $msg;
		echo "<br/> <a href='javascript:self.history.back();' class='btn btn-primary'>Voltar</a>";
	} elseif (!$check_age) {
		echo "<span class='alert alert-danger'>Por favor, forneça a idade apropriada!</span>";
	} elseif (!$check_email) {
		echo "<span class='alert alert-danger'>Por favor, forneça o email apropriado!</span>";
	} else {
		$result = $crud->execute("UPDATE usuarios SET nome='$nome', idade='$idade', email='$email' WHERE id = $id");
		header("Location: index.php");
	}
}

?>
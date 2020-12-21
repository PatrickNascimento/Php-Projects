<?php

include_once("class/Crud.php");

$crud = new Crud();

$query = "SELECT * FROM usuarios ORDER BY id DESC";
$result = $crud->getData($query);

?>

<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8">
	<title></title>
	<link rel="stylesheet" href="assets/css/bootstrap.css">
</head>
<body class="container mt-4">
	<div class="row">
		<div class="col-lg-10">
			<h1 class="text-center">Lista de usuários</h1>
		</div>
		<div class="col-lg-2">
			<a href="add.html" class="btn btn-success">Adicionar usuário</a><br/>
		</div>

	</div>
	
	<table class="table mt-3">
		<tr>
			<th>Nome</th>
			<th>Idade</th>
			<th>E-mail</th>
			<th>Opções</th>
		</tr>
		<tr>
			<?php
				foreach ($result as $key => $res) {
					echo "<tr>";
						echo "<td>".$res['nome']."</td>";
						echo "<td>".$res['idade']."</td>";
						echo "<td>".$res['email']."</td>";
						echo "<td><a href=\"edit.php?id=$res[id]\" class='btn btn-warning mr-2'>Editar</a><a href=\"delete.php?id=$res[id]\" onClick=\"return confirm('Você tem certeza que quer deletar esse usuário?')\" class='btn btn-danger'>Deletar</a></td>";
					echo "</tr>";
				}
			?>
		</tr>
	</table>
</body>
</html>
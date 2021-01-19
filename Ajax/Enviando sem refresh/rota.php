<?php
$data['mensagem'] = "Nome {$_POST['nome']} | e-mail {$_POST['email']}";
echo json_encode($data);
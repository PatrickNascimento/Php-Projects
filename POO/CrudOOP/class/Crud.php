<?php

include 'DbConfig.php';

class Crud extends DbConfig {

	public function __construct() {
		parent::__construct();
	}


	public function getData($query) {
		$result = $this->connection->query($query);

		if (!$result) {
			return FALSE;
		}

		$rows = array();

		while ($row = $result->fetch_assoc()) {
			$rows[] = $row;
		}

		return $rows;
	}


	public function execute($query) {
		$result = $this->connection->query($query);

		if (!$result) {
			echo 'Erro ao executar o comando!';
			return FALSE;
		} else {
			return TRUE;
		}
	}


	public function delete($id, $table) {
		$query = "DELETE FROM $table WHERE id = $id";

		$result = $this->connection->query($query);

		if (!$result) {
			echo '<span class="alert alert-danger text-center mt-3">Error ao tentar deletar o usuário de id ' . $id . ' da tabela ' . $table . '</span>';
			return FALSE;
		} else {
			return TRUE;
		}
	}


	public function escape_string($value) {
		return $this->connection->real_escape_string($value);
	}
}

?>
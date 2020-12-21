<?php

class DbConfig {
	private $host = 'localhost';
	private $database = 'php_oop';
	private $user = 'root';
	private $password = '54321';

	protected $connection;

	public function __construct() {
		if (!isset($this->connection)) {
			$this->connection = new mysqli($this->host, $this->user, $this->password, $this->database);
			if (!$this->connection) {
				echo 'Erro ao conectar!';
				exit;
			}
		}
		return $this->connection;
	}
}

?>
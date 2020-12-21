<?php

class Validation {

	public function check_empty($data, $fields) {
		$msg = null;
		foreach ($fields as $value) {
			if (empty($data[$value])) {
				$msg = "<span class='alert alert-danger'>Campo $value vazio!</span><br/>";
			}
		}
		return $msg;
	}


	public function is_age_valid($age) {
		if (preg_match("/^[0-9]+$/", $age)) {
			return TRUE;
		}
		return FALSE;
	}


	public function is_email_valid($email) {
		if (filter_var($email, FILTER_VALIDATE_EMAIL)) {
			return TRUE;
		} 
		return FALSE;
	}

}


?>
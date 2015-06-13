<?php
require_once('../vendor/autoload.php');


$data = array(
	'Name'  => 'Redmann',
	'Sex'   => 'male'
);

$excel = new EasyPHPExcel('My first Excel');

$excel->setHeader(array_keys($data))
		->addRow(array_values($data))
		->save('example1.xlsx');

<?php
require_once('../vendor/autoload.php');


$data = array(
	'Name'  => 'Tobias Redmann',
	'Sex'   => 'male',
	'Job'   => 'Freelance Software Developer and Consultant'
);

$excel = new EasyPHPExcel('My first Excel');

$excel->setHeader(array_keys($data))
		->addRow(array_values($data))
		->save('example1.xlsx');

<?php
require_once('../vendor/autoload.php');

$header = array('Name', 'Sex', 'Job');

$data = array(
	array('Tobias Redmann', 'male', 'Freelance Software Developer'),
	array('Michael Schumacher', 'male', 'Formula One World Champion'),
	array('Michael Jackson', 'male', 'King Of Pop')
);

$excel = new EasyPHPExcel('');

$excel->setHeader($header)
      ->addRows($data)
      ->save('example-02.xlsx');

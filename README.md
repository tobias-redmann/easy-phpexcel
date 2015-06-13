# easy-phpexcel
It's a simple wrapper to support easy Excel document creation with PHPExcel

## Usage

The following minimal example will create an Excel file.

	$data = array(
		'Name'  => 'Redmann',
		'Sex'   => 'male'
	);
	
	$excel = new EasyPHPExcel('My first Excel');
	
	$excel->setHeader(array_keys($data))
			->addRow(array_values($data))
			->save('example1.xlsx');

That's it. Really! It's that simple.
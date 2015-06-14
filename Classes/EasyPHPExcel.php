<?php

class EasyPHPExcel{

	protected $title;
	protected $description;
	protected $creator;

	private $header;
	private $rows;

	private $columnCount;
	private $currentRow;

	private $objPHPExcel;

	private $columnNames;


	/**
	 * @param string $title
	 * @param string $description
	 * @param string $creator
	 */
	function __construct($title = '', $description = '', $creator = '')
	{

		$this->title = $title;
		$this->description = $description;
		$this->creator = $creator;

		$this->columnCount = 0;
		$this->currentRow = 1;

		$this->header   = array();
		$this->rows     = array();

		$this->objPHPExcel = new PHPExcel();
		$this->objPHPExcel->getProperties()->setCreator($creator);
		$this->objPHPExcel->getProperties()->setLastModifiedBy($creator);
		$this->objPHPExcel->getProperties()->setTitle($title);
		$this->objPHPExcel->getProperties()->setDescription($description);
		$this->objPHPExcel->setActiveSheetIndex(0);

		$this->columnNames = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

	}


	/**
	 * @param array $header
	 *
	 * @return $this
	 */
	public function setHeader($header)
	{

		$this->header = $header;

		if (count($header) > $this->columnCount) {

			$this->columnCount = count($header);

		}

		return $this;

	}

	/**
	 * @param array $row
	 *
	 * @return $this
	 */
	public function addRow($row)
	{

		$this->rows[] = $row;

		if (count($row) > $this->columnCount) {

			$this->columnCount = count($row);

		}

		return $this;

	}

	/**
	 * @param int $index
	 *
	 * @return string
	 */
	private function getColumnCharacter($index)
	{

		return substr($this->columnNames, $index, 1);

	}


	/**
	 * Generate based on rows and header the document
	 *
	 * @return $this
	 */
	private function buildDocument()
	{

		// header

		if (count($this->header) > 0) {

			for($i = 0; $i < $this->columnCount; $i++) {

				$currentCell = $this->getColumnCharacter($i) . $this->currentRow;

				$currentData = '';

				if (isset($this->header[$i])) {
					$currentData = $this->header[$i];
				}

				$this->objPHPExcel->getActiveSheet()->SetCellValue($currentCell , $currentData);

				$this->objPHPExcel->getActiveSheet()->getStyle($currentCell)->getFont()->setBold(true);

			}

			$this->currentRow++;

		}

		if (count($this->rows) > 0) {

			foreach($this->rows as $row) {

				for($i = 0; $i < $this->columnCount; $i++) {

					$currentCell = $this->getColumnCharacter($i) . $this->currentRow;

					$currentData = '';

					if (isset($row[$i])) {
						$currentData = $row[$i];
					}

					$this->objPHPExcel->getActiveSheet()->SetCellValue($currentCell , $currentData);

				}

				$this->currentRow++;

			}

		}

		$this->applyAutoSizing();

		return $this;

	}


	/**
	 * Save the document to file system
	 *
	 * @param $file
	 *
	 * @throws PHPExcel_Writer_Exception
	 */
	public function save($file)
	{
		$this->buildDocument();

		$objWriter = new PHPExcel_Writer_Excel2007($this->objPHPExcel);
		$objWriter->save($file);


	}

	/**
	 * Apply auto sizing 
	 */
	private function applyAutoSizing()
	{

		foreach (range('A', $this->objPHPExcel->getActiveSheet()->getHighestDataColumn()) as $col) {

			$this->objPHPExcel->getActiveSheet()
			               ->getColumnDimension($col)
			               ->setAutoSize(true);

		}

	}


}
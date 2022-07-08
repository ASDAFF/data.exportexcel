<?php
/**
 * DATAPHPExcel
 *
 * Copyright (c) 2006 - 2013 DATAPHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   DATAPHPExcel
 * @package    DATAPHPExcel_Reader
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.7.9, 2013-06-02
 */


/** DATAPHPExcel root directory */
if (!defined('DATAPHPEXCEL_ROOT')) {
	/**
	 * @ignore
	 */
	define('DATAPHPEXCEL_ROOT', dirname(__FILE__) . '/../../');
	require(DATAPHPEXCEL_ROOT . 'DATAPHPExcel/Autoloader.php');
}

/**
 * DATAPHPExcel_Reader_CSV
 *
 * @category   DATAPHPExcel
 * @package    DATAPHPExcel_Reader
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 */
class DATAPHPExcel_Reader_DBF extends DATAPHPExcel_Reader_Abstract implements DATAPHPExcel_Reader_IReader
{
	/**
	 * Input encoding
	 *
	 * @access	private
	 * @var	string
	 */
	private $_inputEncoding	= 'CP866';

	/**
	 * Sheet index to read
	 *
	 * @access	private
	 * @var	int
	 */
	private $_sheetIndex	= 0;

	/**
	 * Load rows contiguously
	 *
	 * @access	private
	 * @var	int
	 */
	private $_contiguous	= false;

	/**
	 * Row counter for loading rows contiguously
	 *
	 * @var	int
	 */
	private $_contiguousRow	= -1;
	
	/**
	 * File row for start reading
	 *
	 * @var	int
	 */
	private $_startFileRow = 1;


	/**
	 * Create a new DATAPHPExcel_Reader_CSV
	 */
	public function __construct() {
		$this->_readFilter		= new DATAPHPExcel_Reader_DefaultReadFilter();
		
		$dir = dirname(__FILE__).'/XBase/';
		require_once($dir.'Table.php');
		require_once($dir.'Column.php');
		require_once($dir.'Record.php');
		require_once($dir.'Memo.php');
	}

	/**
	 * Validate that the current file is a CSV file
	 *
	 * @return boolean
	 */
	protected function _isValidFormat()
	{
		return TRUE;
	}
	
	/**
	 * Get start file row
	 *
	 * @return int
	 */
	public function getStartFileRow()
	{
		if(isset($this->_startFileRow)) return (int)$this->_startFileRow;
		else return 1;
	}

	/**
	 * Get input encoding
	 *
	 * @return string
	 */
	public function getInputEncoding()
	{
		return $this->_inputEncoding;
	}

	/**
	 * Return worksheet info (Name, Last Column Letter, Last Column Index, Total Rows, Total Columns)
	 *
	 * @param 	string 		$pFilename
	 * @throws	DATAPHPExcel_Reader_Exception
	 */
	public function listWorksheetInfo($pFilename)
	{
		// Open file
		$this->_openFile($pFilename);
		if (!$this->_isValidFormat()) {
			fclose($this->_fileHandle);
			throw new DATAPHPExcel_Reader_Exception($pFilename . " is an Invalid Spreadsheet file.");
		}
		fclose($this->_fileHandle);

		$table = new \Xbase\Table($pFilename);

		$worksheetInfo = array();
		$worksheetInfo[0]['worksheetName'] = 'Worksheet';
		$worksheetInfo[0]['lastColumnLetter'] = 'A';
		$worksheetInfo[0]['lastColumnIndex'] = 0;
		$worksheetInfo[0]['totalRows'] = 0;
		$worksheetInfo[0]['totalColumns'] = 0;
		
		$worksheetInfo[0]['totalRows'] = $table->getRecordCount();
		$worksheetInfo[0]['lastColumnIndex'] = $table->getColumnCount() - 1;

		$worksheetInfo[0]['lastColumnLetter'] = DATAPHPExcel_Cell::stringFromColumnIndex($worksheetInfo[0]['lastColumnIndex']);
		$worksheetInfo[0]['totalColumns'] = $worksheetInfo[0]['lastColumnIndex'] + 1;

		// Close file
		$table->close();

		return $worksheetInfo;
	}

	/**
	 * Loads DATAPHPExcel from file
	 *
	 * @param 	string 		$pFilename
	 * @return DATAPHPExcel
	 * @throws DATAPHPExcel_Reader_Exception
	 */
	public function load($pFilename)
	{		
		// Create new DATAPHPExcel
		$objDATAPHPExcel = new DATAPHPExcel();

		// Load into this instance
		return $this->loadIntoExisting($pFilename, $objDATAPHPExcel);
	}

	/**
	 * Loads DATAPHPExcel from file into DATAPHPExcel instance
	 *
	 * @param 	string 		$pFilename
	 * @param	DATAPHPExcel	$objDATAPHPExcel
	 * @return 	DATAPHPExcel
	 * @throws 	DATAPHPExcel_Reader_Exception
	 */
	public function loadIntoExisting($pFilename, DATAPHPExcel $objDATAPHPExcel)
	{
		// Open file
		$this->_openFile($pFilename);
		if (!$this->_isValidFormat()) {
			fclose ($this->_fileHandle);
			throw new DATAPHPExcel_Reader_Exception($pFilename . " is an Invalid Spreadsheet file.");
		}
		fclose($this->_fileHandle);
		
		$table = new \Xbase\Table($pFilename);
		$columns = $table->getColumns();

		// Create new DATAPHPExcel object
		while ($objDATAPHPExcel->getSheetCount() <= $this->_sheetIndex) {
			$objDATAPHPExcel->createSheet();
		}
		$sheet = $objDATAPHPExcel->setActiveSheetIndex($this->_sheetIndex);
		if(is_callable(array($sheet, 'setRealHighestRow')))
		{
			$sheet->setRealHighestRow($table->getRecordCount());
		}

		// Set our starting row based on whether we're in contiguous mode or not
		$currentRow = 1;
		$endRow = $sheet->getHighestRow();
		/*$currentRow = $this->getStartFileRow();
		if ($this->_contiguous) {
			$currentRow = ($this->_contiguousRow == -1) ? $sheet->getHighestRow(): $this->_contiguousRow;
		}*/
		
		if(method_exists($this->_readFilter, 'getStartRow')) $currentRow = $this->_readFilter->getStartRow();
		if(method_exists($this->_readFilter, 'getEndRow')) $endRow = $this->_readFilter->getEndRow();
		
		// Loop through each line of the file in turn
		while($currentRow<=$endRow)
		{
			if($currentRow - 1 == 0)
			{
				
				foreach($columns as $column)
				{
					$rowData[] = $column->getRawname();
				}
			}
			else
			{
				$record = $table->moveTo($currentRow - 2);
				$rowData = array();
				foreach ($columns as $column) {
					$rowData[] = $record->forceGetString($column->name);
				}
			}
			$columnLetter = 'A';
			foreach($rowData as $rowDatum) {
				if ($rowDatum != '' && $this->_readFilter->readCell($columnLetter, $currentRow)) {
					// Convert encoding if necessary
					/*if ($this->_inputEncoding !== 'UTF-8') {
						$rowDatum = DATAPHPExcel_Shared_String::ConvertEncoding($rowDatum, 'UTF-8', $this->_inputEncoding);
					}*/

					// Set cell value
					$sheet->getCell($columnLetter . $currentRow)->setValue($rowDatum);
				}
				++$columnLetter;
			}
			++$currentRow;
		}

		// Close file
		$table->close();

		if ($this->_contiguous) {
			$this->_contiguousRow = $currentRow;
		}

		// Return
		return $objDATAPHPExcel;
	}

	/**
	 * Get sheet index
	 *
	 * @return integer
	 */
	public function getSheetIndex() {
		return $this->_sheetIndex;
	}

	/**
	 * Set sheet index
	 *
	 * @param	integer		$pValue		Sheet index
	 * @return DATAPHPExcel_Reader_CSV
	 */
	public function setSheetIndex($pValue = 0) {
		$this->_sheetIndex = $pValue;
		return $this;
	}

	/**
	 * Set Contiguous
	 *
	 * @param boolean $contiguous
	 */
	public function setContiguous($contiguous = FALSE)
	{
		$this->_contiguous = (bool) $contiguous;
		if (!$contiguous) {
			$this->_contiguousRow = -1;
		}

		return $this;
	}

	/**
	 * Get Contiguous
	 *
	 * @return boolean
	 */
	public function getContiguous() {
		return $this->_contiguous;
	}

}

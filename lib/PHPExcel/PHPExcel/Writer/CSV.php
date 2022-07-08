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
 * @package	DATAPHPExcel_Writer_CSV
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 * @license	http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version	1.7.9, 2013-06-02
 */


/**
 * DATAPHPExcel_Writer_CSV
 *
 * @category   DATAPHPExcel
 * @package	DATAPHPExcel_Writer_CSV
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 */
class DATAPHPExcel_Writer_CSV extends DATAPHPExcel_Writer_Abstract implements DATAPHPExcel_Writer_IWriter {
	/**
	 * DATAPHPExcel object
	 *
	 * @var DATAPHPExcel
	 */
	private $_phpExcel;

	/**
	 * Delimiter
	 *
	 * @var string
	 */
	private $_delimiter	= ',';

	/**
	 * Enclosure
	 *
	 * @var string
	 */
	private $_enclosure	= '"';

	/**
	 * Line ending
	 *
	 * @var string
	 */
	private $_lineEnding	= PHP_EOL;

	/**
	 * Sheet index to write
	 *
	 * @var int
	 */
	private $_sheetIndex	= 0;

	/**
	 * Whether to write a BOM (for UTF8).
	 *
	 * @var boolean
	 */
	private $_useBOM = false;

	/**
	 * Whether to write a fully Excel compatible CSV file.
	 *
	 * @var boolean
	 */
	private $_excelCompatibility = false;

	/**
	 * Create a new DATAPHPExcel_Writer_CSV
	 *
	 * @param	DATAPHPExcel	$phpExcel	DATAPHPExcel object
	 */
	public function __construct(DATAPHPExcel $phpExcel) {
		$this->_phpExcel	= $phpExcel;
	}

	/**
	 * Save DATAPHPExcel to file
	 *
	 * @param	string		$pFilename
	 * @throws	DATAPHPExcel_Writer_Exception
	 */
	public function save($pFilename = null) {
		// Fetch sheet
		$sheet = $this->_phpExcel->getSheet($this->_sheetIndex);

		$saveDebugLog = DATAPHPExcel_Calculation::getInstance($this->_phpExcel)->getDebugLog()->getWriteDebugLog();
		DATAPHPExcel_Calculation::getInstance($this->_phpExcel)->getDebugLog()->setWriteDebugLog(FALSE);
		$saveArrayReturnType = DATAPHPExcel_Calculation::getArrayReturnType();
		DATAPHPExcel_Calculation::setArrayReturnType(DATAPHPExcel_Calculation::RETURN_ARRAY_AS_VALUE);

		// Open file
		$fileHandle = fopen($pFilename, 'wb+');
		if ($fileHandle === false) {
			throw new DATAPHPExcel_Writer_Exception("Could not open file $pFilename for writing.");
		}

		if ($this->_excelCompatibility) {
			// Write the UTF-16LE BOM code
			fwrite($fileHandle, "\xFF\xFE");	//	Excel uses UTF-16LE encoding
			$this->setEnclosure();				//	Default enclosure is "
			$this->setDelimiter("\t");			//	Excel delimiter is a TAB
		} elseif ($this->_useBOM) {
			// Write the UTF-8 BOM code
			fwrite($fileHandle, "\xEF\xBB\xBF");
		}

		//	Identify the range that we need to extract from the worksheet
		$maxCol = $sheet->getHighestColumn();
		$maxRow = $sheet->getHighestRow();

		// Write rows to file
		for($row = 1; $row <= $maxRow; ++$row) {
			// Convert the row to an array...
			$cellsArray = $sheet->rangeToArray('A'.$row.':'.$maxCol.$row,'', $this->_preCalculateFormulas);
			// ... and write to the file
			$this->_writeLine($fileHandle, $cellsArray[0]);
		}

		// Close file
		fclose($fileHandle);

		DATAPHPExcel_Calculation::setArrayReturnType($saveArrayReturnType);
		DATAPHPExcel_Calculation::getInstance($this->_phpExcel)->getDebugLog()->setWriteDebugLog($saveDebugLog);
	}

	/**
	 * Get delimiter
	 *
	 * @return string
	 */
	public function getDelimiter() {
		return $this->_delimiter;
	}

	/**
	 * Set delimiter
	 *
	 * @param	string	$pValue		Delimiter, defaults to ,
	 * @return DATAPHPExcel_Writer_CSV
	 */
	public function setDelimiter($pValue = ',') {
		$this->_delimiter = $pValue;
		return $this;
	}

	/**
	 * Get enclosure
	 *
	 * @return string
	 */
	public function getEnclosure() {
		return $this->_enclosure;
	}

	/**
	 * Set enclosure
	 *
	 * @param	string	$pValue		Enclosure, defaults to "
	 * @return DATAPHPExcel_Writer_CSV
	 */
	public function setEnclosure($pValue = '"') {
		if ($pValue == '') {
			$pValue = null;
		}
		$this->_enclosure = $pValue;
		return $this;
	}

	/**
	 * Get line ending
	 *
	 * @return string
	 */
	public function getLineEnding() {
		return $this->_lineEnding;
	}

	/**
	 * Set line ending
	 *
	 * @param	string	$pValue		Line ending, defaults to OS line ending (PHP_EOL)
	 * @return DATAPHPExcel_Writer_CSV
	 */
	public function setLineEnding($pValue = PHP_EOL) {
		$this->_lineEnding = $pValue;
		return $this;
	}

	/**
	 * Get whether BOM should be used
	 *
	 * @return boolean
	 */
	public function getUseBOM() {
		return $this->_useBOM;
	}

	/**
	 * Set whether BOM should be used
	 *
	 * @param	boolean	$pValue		Use UTF-8 byte-order mark? Defaults to false
	 * @return DATAPHPExcel_Writer_CSV
	 */
	public function setUseBOM($pValue = false) {
		$this->_useBOM = $pValue;
		return $this;
	}

	/**
	 * Get whether the file should be saved with full Excel Compatibility
	 *
	 * @return boolean
	 */
	public function getExcelCompatibility() {
		return $this->_excelCompatibility;
	}

	/**
	 * Set whether the file should be saved with full Excel Compatibility
	 *
	 * @param	boolean	$pValue		Set the file to be written as a fully Excel compatible csv file
	 *								Note that this overrides other settings such as useBOM, enclosure and delimiter
	 * @return DATAPHPExcel_Writer_CSV
	 */
	public function setExcelCompatibility($pValue = false) {
		$this->_excelCompatibility = $pValue;
		return $this;
	}

	/**
	 * Get sheet index
	 *
	 * @return int
	 */
	public function getSheetIndex() {
		return $this->_sheetIndex;
	}

	/**
	 * Set sheet index
	 *
	 * @param	int		$pValue		Sheet index
	 * @return DATAPHPExcel_Writer_CSV
	 */
	public function setSheetIndex($pValue = 0) {
		$this->_sheetIndex = $pValue;
		return $this;
	}

	/**
	 * Write line to CSV file
	 *
	 * @param	mixed	$pFileHandle	PHP filehandle
	 * @param	array	$pValues		Array containing values in a row
	 * @throws	DATAPHPExcel_Writer_Exception
	 */
	private function _writeLine($pFileHandle = null, $pValues = null) {
		if (is_array($pValues)) {
			// No leading delimiter
			$writeDelimiter = false;

			// Build the line
			$line = '';

			foreach ($pValues as $element) {
				// Escape enclosures
				$element = str_replace($this->_enclosure, $this->_enclosure . $this->_enclosure, $element);

				// Add delimiter
				if ($writeDelimiter) {
					$line .= $this->_delimiter;
				} else {
					$writeDelimiter = true;
				}

				// Add enclosed string
				$line .= $this->_enclosure . $element . $this->_enclosure;
			}

			// Add line ending
			$line .= $this->_lineEnding;

			// Write to file
			if ($this->_excelCompatibility) {
				fwrite($pFileHandle, mb_convert_encoding($line,"UTF-16LE","UTF-8"));
			} else {
				fwrite($pFileHandle, $line);
			}
		} else {
			throw new DATAPHPExcel_Writer_Exception("Invalid data row passed to CSV writer.");
		}
	}

}

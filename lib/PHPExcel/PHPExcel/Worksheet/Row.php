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
 * @package    DATAPHPExcel_Worksheet
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.7.9, 2013-06-02
 */


/**
 * DATAPHPExcel_Worksheet_Row
 *
 * Represents a row in DATAPHPExcel_Worksheet, used by DATAPHPExcel_Worksheet_RowIterator
 *
 * @category   DATAPHPExcel
 * @package    DATAPHPExcel_Worksheet
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 */
class DATAPHPExcel_Worksheet_Row
{
	/**
	 * DATAPHPExcel_Worksheet
	 *
	 * @var DATAPHPExcel_Worksheet
	 */
	private $_parent;

	/**
	 * Row index
	 *
	 * @var int
	 */
	private $_rowIndex = 0;

	/**
	 * Create a new row
	 *
	 * @param DATAPHPExcel_Worksheet 		$parent
	 * @param int						$rowIndex
	 */
	public function __construct(DATAPHPExcel_Worksheet $parent = null, $rowIndex = 1) {
		// Set parent and row index
		$this->_parent 		= $parent;
		$this->_rowIndex 	= $rowIndex;
	}

	/**
	 * Destructor
	 */
	public function __destruct() {
		unset($this->_parent);
	}

	/**
	 * Get row index
	 *
	 * @return int
	 */
	public function getRowIndex() {
		return $this->_rowIndex;
	}

	/**
	 * Get cell iterator
	 *
	 * @return DATAPHPExcel_Worksheet_CellIterator
	 */
	public function getCellIterator() {
		return new DATAPHPExcel_Worksheet_CellIterator($this->_parent, $this->_rowIndex);
	}
}

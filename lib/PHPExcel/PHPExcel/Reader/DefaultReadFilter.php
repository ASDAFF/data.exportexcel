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
 * DATAPHPExcel_Reader_DefaultReadFilter
 *
 * @category   DATAPHPExcel
 * @package    DATAPHPExcel_Reader
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 */
class DATAPHPExcel_Reader_DefaultReadFilter implements DATAPHPExcel_Reader_IReadFilter
{
	/**
	 * Should this cell be read?
	 *
	 * @param 	$column		String column index
	 * @param 	$row			Row index
	 * @param	$worksheetName	Optional worksheet name
	 * @return	boolean
	 */
	public function readCell($column, $row, $worksheetName = '') {
		return true;
	}
}

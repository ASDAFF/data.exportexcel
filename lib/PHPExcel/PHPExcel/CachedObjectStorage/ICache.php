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
 * @package    DATAPHPExcel_CachedObjectStorage
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.7.9, 2013-06-02
 */


/**
 * DATAPHPExcel_CachedObjectStorage_ICache
 *
 * @category   DATAPHPExcel
 * @package    DATAPHPExcel_CachedObjectStorage
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 */
interface DATAPHPExcel_CachedObjectStorage_ICache
{
    /**
     * Add or Update a cell in cache identified by coordinate address
     *
     * @param	string			$pCoord		Coordinate address of the cell to update
     * @param	DATAPHPExcel_Cell	$cell		Cell to update
	 * @return	void
     * @throws	DATAPHPExcel_Exception
     */
	public function addCacheData($pCoord, DATAPHPExcel_Cell $cell);

    /**
     * Add or Update a cell in cache
     *
     * @param	DATAPHPExcel_Cell	$cell		Cell to update
	 * @return	void
     * @throws	DATAPHPExcel_Exception
     */
	public function updateCacheData(DATAPHPExcel_Cell $cell);

    /**
     * Fetch a cell from cache identified by coordinate address
     *
     * @param	string			$pCoord		Coordinate address of the cell to retrieve
     * @return DATAPHPExcel_Cell 	Cell that was found, or null if not found
     * @throws	DATAPHPExcel_Exception
     */
	public function getCacheData($pCoord);

    /**
     * Delete a cell in cache identified by coordinate address
     *
     * @param	string			$pCoord		Coordinate address of the cell to delete
     * @throws	DATAPHPExcel_Exception
     */
	public function deleteCacheData($pCoord);

	/**
	 * Is a value set in the current DATAPHPExcel_CachedObjectStorage_ICache for an indexed cell?
	 *
	 * @param	string		$pCoord		Coordinate address of the cell to check
	 * @return	boolean
	 */
	public function isDataSet($pCoord);

	/**
	 * Get a list of all cell addresses currently held in cache
	 *
	 * @return	array of string
	 */
	public function getCellList();

	/**
	 * Get the list of all cell addresses currently held in cache sorted by column and row
	 *
	 * @return	void
	 */
	public function getSortedCellList();

	/**
	 * Clone the cell collection
	 *
	 * @param	DATAPHPExcel_Worksheet	$parent		The new worksheet
	 * @return	void
	 */
	public function copyCellCollection(DATAPHPExcel_Worksheet $parent);

	/**
	 * Identify whether the caching method is currently available
	 * Some methods are dependent on the availability of certain extensions being enabled in the PHP build
	 *
	 * @return	boolean
	 */
	public static function cacheMethodIsAvailable();

}

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
 * @package    DATAPHPExcel_Shared_Escher
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.7.9, 2013-06-02
 */

/**
 * DATAPHPExcel_Shared_Escher_DgContainer
 *
 * @category   DATAPHPExcel
 * @package    DATAPHPExcel_Shared_Escher
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 */
class DATAPHPExcel_Shared_Escher_DgContainer
{
	/**
	 * Drawing index, 1-based.
	 *
	 * @var int
	 */
	private $_dgId;

	/**
	 * Last shape index in this drawing
	 *
	 * @var int
	 */
	private $_lastSpId;

	private $_spgrContainer = null;

	public function getDgId()
	{
		return $this->_dgId;
	}

	public function setDgId($value)
	{
		$this->_dgId = $value;
	}

	public function getLastSpId()
	{
		return $this->_lastSpId;
	}

	public function setLastSpId($value)
	{
		$this->_lastSpId = $value;
	}

	public function getSpgrContainer()
	{
		return $this->_spgrContainer;
	}

	public function setSpgrContainer($spgrContainer)
	{
		return $this->_spgrContainer = $spgrContainer;
	}

}

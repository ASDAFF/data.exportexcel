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
 * @category	DATAPHPExcel
 * @package		DATAPHPExcel_Chart
 * @copyright	Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 * @license		http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version		1.7.9, 2013-06-02
 */


/**
 * DATAPHPExcel_Chart_Title
 *
 * @category	DATAPHPExcel
 * @package		DATAPHPExcel_Chart
 * @copyright	Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 */
class DATAPHPExcel_Chart_Title
{

	/**
	 * Title Caption
	 *
	 * @var string
	 */
	private $_caption = null;

	/**
	 * Title Layout
	 *
	 * @var DATAPHPExcel_Chart_Layout
	 */
	private $_layout = null;

	/**
	 * Create a new DATAPHPExcel_Chart_Title
	 */
	public function __construct($caption = null, DATAPHPExcel_Chart_Layout $layout = null)
	{
		$this->_caption = $caption;
		$this->_layout = $layout;
	}

	/**
	 * Get caption
	 *
	 * @return string
	 */
	public function getCaption() {
		return $this->_caption;
	}

	/**
	 * Set caption
	 *
	 * @param string $caption
	 */
	public function setCaption($caption = null) {
		$this->_caption = $caption;
	}

	/**
	 * Get Layout
	 *
	 * @return DATAPHPExcel_Chart_Layout
	 */
	public function getLayout() {
		return $this->_layout;
	}

}

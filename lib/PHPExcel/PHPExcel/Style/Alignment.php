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
 * @package	DATAPHPExcel_Style
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 * @license	http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version	1.7.9, 2013-06-02
 */


/**
 * DATAPHPExcel_Style_Alignment
 *
 * @category   DATAPHPExcel
 * @package	DATAPHPExcel_Style
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 */
class DATAPHPExcel_Style_Alignment extends DATAPHPExcel_Style_Supervisor implements DATAPHPExcel_IComparable
{
	/* Horizontal alignment styles */
	const HORIZONTAL_GENERAL				= 'general';
	const HORIZONTAL_LEFT					= 'left';
	const HORIZONTAL_RIGHT					= 'right';
	const HORIZONTAL_CENTER					= 'center';
	const HORIZONTAL_CENTER_CONTINUOUS		= 'centerContinuous';
	const HORIZONTAL_JUSTIFY				= 'justify';

	/* Vertical alignment styles */
	const VERTICAL_BOTTOM					= 'bottom';
	const VERTICAL_TOP						= 'top';
	const VERTICAL_CENTER					= 'center';
	const VERTICAL_JUSTIFY					= 'justify';

	/**
	 * Horizontal
	 *
	 * @var string
	 */
	protected $_horizontal	= DATAPHPExcel_Style_Alignment::HORIZONTAL_GENERAL;

	/**
	 * Vertical
	 *
	 * @var string
	 */
	protected $_vertical		= DATAPHPExcel_Style_Alignment::VERTICAL_BOTTOM;

	/**
	 * Text rotation
	 *
	 * @var int
	 */
	protected $_textRotation	= 0;

	/**
	 * Wrap text
	 *
	 * @var boolean
	 */
	protected $_wrapText		= FALSE;

	/**
	 * Shrink to fit
	 *
	 * @var boolean
	 */
	protected $_shrinkToFit	= FALSE;

	/**
	 * Indent - only possible with horizontal alignment left and right
	 *
	 * @var int
	 */
	protected $_indent		= 0;

	/**
	 * Create a new DATAPHPExcel_Style_Alignment
	 *
	 * @param	boolean	$isSupervisor	Flag indicating if this is a supervisor or not
	 *									Leave this value at default unless you understand exactly what
	 *										its ramifications are
	 * @param	boolean	$isConditional	Flag indicating if this is a conditional style or not
	 *									Leave this value at default unless you understand exactly what
	 *										its ramifications are
	 */
	public function __construct($isSupervisor = FALSE, $isConditional = FALSE)
	{
		// Supervisor?
		parent::__construct($isSupervisor);

		if ($isConditional) {
			$this->_horizontal		= NULL;
			$this->_vertical		= NULL;
			$this->_textRotation	= NULL;
		}
	}

	/**
	 * Get the shared style component for the currently active cell in currently active sheet.
	 * Only used for style supervisor
	 *
	 * @return DATAPHPExcel_Style_Alignment
	 */
	public function getSharedComponent()
	{
		return $this->_parent->getSharedComponent()->getAlignment();
	}

	/**
	 * Build style array from subcomponents
	 *
	 * @param array $array
	 * @return array
	 */
	public function getStyleArray($array)
	{
		return array('alignment' => $array);
	}

	/**
	 * Apply styles from array
	 *
	 * <code>
	 * $objDATAPHPExcel->getActiveSheet()->getStyle('B2')->getAlignment()->applyFromArray(
	 *		array(
	 *			'horizontal' => DATAPHPExcel_Style_Alignment::HORIZONTAL_CENTER,
	 *			'vertical'   => DATAPHPExcel_Style_Alignment::VERTICAL_CENTER,
	 *			'rotation'   => 0,
	 *			'wrap'			=> TRUE
	 *		)
	 * );
	 * </code>
	 *
	 * @param	array	$pStyles	Array containing style information
	 * @throws	DATAPHPExcel_Exception
	 * @return DATAPHPExcel_Style_Alignment
	 */
	public function applyFromArray($pStyles = NULL) {
		if (is_array($pStyles)) {
			if ($this->_isSupervisor) {
				$this->getActiveSheet()->getStyle($this->getSelectedCells())
				    ->applyFromArray($this->getStyleArray($pStyles));
			} else {
				if (isset($pStyles['horizontal'])) {
					$this->setHorizontal($pStyles['horizontal']);
				}
				if (isset($pStyles['vertical'])) {
					$this->setVertical($pStyles['vertical']);
				}
				if (isset($pStyles['rotation'])) {
					$this->setTextRotation($pStyles['rotation']);
				}
				if (isset($pStyles['wrap'])) {
					$this->setWrapText($pStyles['wrap']);
				}
				if (isset($pStyles['shrinkToFit'])) {
					$this->setShrinkToFit($pStyles['shrinkToFit']);
				}
				if (isset($pStyles['indent'])) {
					$this->setIndent($pStyles['indent']);
				}
			}
		} else {
			throw new DATAPHPExcel_Exception("Invalid style array passed.");
		}
		return $this;
	}

	/**
	 * Get Horizontal
	 *
	 * @return string
	 */
	public function getHorizontal() {
		if ($this->_isSupervisor) {
			return $this->getSharedComponent()->getHorizontal();
		}
		return $this->_horizontal;
	}

	/**
	 * Set Horizontal
	 *
	 * @param string $pValue
	 * @return DATAPHPExcel_Style_Alignment
	 */
	public function setHorizontal($pValue = DATAPHPExcel_Style_Alignment::HORIZONTAL_GENERAL) {
		if ($pValue == '') {
			$pValue = DATAPHPExcel_Style_Alignment::HORIZONTAL_GENERAL;
		}

		if ($this->_isSupervisor) {
			$styleArray = $this->getStyleArray(array('horizontal' => $pValue));
			$this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
		}
		else {
			$this->_horizontal = $pValue;
		}
		return $this;
	}

	/**
	 * Get Vertical
	 *
	 * @return string
	 */
	public function getVertical() {
		if ($this->_isSupervisor) {
			return $this->getSharedComponent()->getVertical();
		}
		return $this->_vertical;
	}

	/**
	 * Set Vertical
	 *
	 * @param string $pValue
	 * @return DATAPHPExcel_Style_Alignment
	 */
	public function setVertical($pValue = DATAPHPExcel_Style_Alignment::VERTICAL_BOTTOM) {
		if ($pValue == '') {
			$pValue = DATAPHPExcel_Style_Alignment::VERTICAL_BOTTOM;
		}

		if ($this->_isSupervisor) {
			$styleArray = $this->getStyleArray(array('vertical' => $pValue));
			$this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
		} else {
			$this->_vertical = $pValue;
		}
		return $this;
	}

	/**
	 * Get TextRotation
	 *
	 * @return int
	 */
	public function getTextRotation() {
		if ($this->_isSupervisor) {
			return $this->getSharedComponent()->getTextRotation();
		}
		return $this->_textRotation;
	}

	/**
	 * Set TextRotation
	 *
	 * @param int $pValue
	 * @throws DATAPHPExcel_Exception
	 * @return DATAPHPExcel_Style_Alignment
	 */
	public function setTextRotation($pValue = 0) {
		// Excel2007 value 255 => DATAPHPExcel value -165
		if ($pValue == 255) {
			$pValue = -165;
		}

		// Set rotation
		if ( ($pValue >= -90 && $pValue <= 90) || $pValue == -165 ) {
			if ($this->_isSupervisor) {
				$styleArray = $this->getStyleArray(array('rotation' => $pValue));
				$this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
			} else {
				$this->_textRotation = $pValue;
			}
		} else {
			throw new DATAPHPExcel_Exception("Text rotation should be a value between -90 and 90.");
		}

		return $this;
	}

	/**
	 * Get Wrap Text
	 *
	 * @return boolean
	 */
	public function getWrapText() {
		if ($this->_isSupervisor) {
			return $this->getSharedComponent()->getWrapText();
		}
		return $this->_wrapText;
	}

	/**
	 * Set Wrap Text
	 *
	 * @param boolean $pValue
	 * @return DATAPHPExcel_Style_Alignment
	 */
	public function setWrapText($pValue = FALSE) {
		if ($pValue == '') {
			$pValue = FALSE;
		}
		if ($this->_isSupervisor) {
			$styleArray = $this->getStyleArray(array('wrap' => $pValue));
			$this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
		} else {
			$this->_wrapText = $pValue;
		}
		return $this;
	}

	/**
	 * Get Shrink to fit
	 *
	 * @return boolean
	 */
	public function getShrinkToFit() {
		if ($this->_isSupervisor) {
			return $this->getSharedComponent()->getShrinkToFit();
		}
		return $this->_shrinkToFit;
	}

	/**
	 * Set Shrink to fit
	 *
	 * @param boolean $pValue
	 * @return DATAPHPExcel_Style_Alignment
	 */
	public function setShrinkToFit($pValue = FALSE) {
		if ($pValue == '') {
			$pValue = FALSE;
		}
		if ($this->_isSupervisor) {
			$styleArray = $this->getStyleArray(array('shrinkToFit' => $pValue));
			$this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
		} else {
			$this->_shrinkToFit = $pValue;
		}
		return $this;
	}

	/**
	 * Get indent
	 *
	 * @return int
	 */
	public function getIndent() {
		if ($this->_isSupervisor) {
			return $this->getSharedComponent()->getIndent();
		}
		return $this->_indent;
	}

	/**
	 * Set indent
	 *
	 * @param int $pValue
	 * @return DATAPHPExcel_Style_Alignment
	 */
	public function setIndent($pValue = 0) {
		if ($pValue > 0) {
			if ($this->getHorizontal() != self::HORIZONTAL_GENERAL &&
				$this->getHorizontal() != self::HORIZONTAL_LEFT &&
				$this->getHorizontal() != self::HORIZONTAL_RIGHT) {
				$pValue = 0; // indent not supported
			}
		}
		if ($this->_isSupervisor) {
			$styleArray = $this->getStyleArray(array('indent' => $pValue));
			$this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
		} else {
			$this->_indent = $pValue;
		}
		return $this;
	}

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
		if ($this->_isSupervisor) {
			return $this->getSharedComponent()->getHashCode();
		}
		return md5(
			  $this->_horizontal
			. $this->_vertical
			. $this->_textRotation
			. ($this->_wrapText ? 't' : 'f')
			. ($this->_shrinkToFit ? 't' : 'f')
			. $this->_indent
			. __CLASS__
		);
	}

}

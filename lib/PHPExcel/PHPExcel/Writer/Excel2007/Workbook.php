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
 * @package    DATAPHPExcel_Writer_Excel2007
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.7.9, 2013-06-02
 */


/**
 * DATAPHPExcel_Writer_Excel2007_Workbook
 *
 * @category   DATAPHPExcel
 * @package    DATAPHPExcel_Writer_Excel2007
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 */
class DATAPHPExcel_Writer_Excel2007_Workbook extends DATAPHPExcel_Writer_Excel2007_WriterPart
{
	/**
	 * Write workbook to XML format
	 *
	 * @param 	DATAPHPExcel	$pDATAPHPExcel
	 * @param	boolean		$recalcRequired	Indicate whether formulas should be recalculated before writing
	 * @return 	string 		XML Output
	 * @throws 	DATAPHPExcel_Writer_Exception
	 */
	public function writeWorkbook(DATAPHPExcel $pDATAPHPExcel = null, $recalcRequired = FALSE)
	{
		// Create XML writer
		$objWriter = null;
		if ($this->getParentWriter()->getUseDiskCaching()) {
			$objWriter = new DATAPHPExcel_Shared_XMLWriter(DATAPHPExcel_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
		} else {
			$objWriter = new DATAPHPExcel_Shared_XMLWriter(DATAPHPExcel_Shared_XMLWriter::STORAGE_MEMORY);
		}

		// XML header
		$objWriter->startDocument('1.0','UTF-8','yes');

		// workbook
		$objWriter->startElement('workbook');
		$objWriter->writeAttribute('xml:space', 'preserve');
		$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
		$objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

			// fileVersion
			$this->_writeFileVersion($objWriter);

			// workbookPr
			$this->_writeWorkbookPr($objWriter);

			// workbookProtection
			$this->_writeWorkbookProtection($objWriter, $pDATAPHPExcel);

			// bookViews
			if ($this->getParentWriter()->getOffice2003Compatibility() === false) {
				$this->_writeBookViews($objWriter, $pDATAPHPExcel);
			}

			// sheets
			$this->_writeSheets($objWriter, $pDATAPHPExcel);

			// definedNames
			$this->_writeDefinedNames($objWriter, $pDATAPHPExcel);

			// calcPr
			$this->_writeCalcPr($objWriter,$recalcRequired);

		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}

	/**
	 * Write file version
	 *
	 * @param 	DATAPHPExcel_Shared_XMLWriter $objWriter 		XML Writer
	 * @throws 	DATAPHPExcel_Writer_Exception
	 */
	private function _writeFileVersion(DATAPHPExcel_Shared_XMLWriter $objWriter = null)
	{
		$objWriter->startElement('fileVersion');
		$objWriter->writeAttribute('appName', 'xl');
		$objWriter->writeAttribute('lastEdited', '4');
		$objWriter->writeAttribute('lowestEdited', '4');
		$objWriter->writeAttribute('rupBuild', '4505');
		$objWriter->endElement();
	}

	/**
	 * Write WorkbookPr
	 *
	 * @param 	DATAPHPExcel_Shared_XMLWriter $objWriter 		XML Writer
	 * @throws 	DATAPHPExcel_Writer_Exception
	 */
	private function _writeWorkbookPr(DATAPHPExcel_Shared_XMLWriter $objWriter = null)
	{
		$objWriter->startElement('workbookPr');

		if (DATAPHPExcel_Shared_Date::getExcelCalendar() == DATAPHPExcel_Shared_Date::CALENDAR_MAC_1904) {
			$objWriter->writeAttribute('date1904', '1');
		}

		$objWriter->writeAttribute('codeName', 'ThisWorkbook');

		$objWriter->endElement();
	}

	/**
	 * Write BookViews
	 *
	 * @param 	DATAPHPExcel_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	DATAPHPExcel					$pDATAPHPExcel
	 * @throws 	DATAPHPExcel_Writer_Exception
	 */
	private function _writeBookViews(DATAPHPExcel_Shared_XMLWriter $objWriter = null, DATAPHPExcel $pDATAPHPExcel = null)
	{
		// bookViews
		$objWriter->startElement('bookViews');

			// workbookView
			$objWriter->startElement('workbookView');

			$objWriter->writeAttribute('activeTab', $pDATAPHPExcel->getActiveSheetIndex());
			$objWriter->writeAttribute('autoFilterDateGrouping', '1');
			$objWriter->writeAttribute('firstSheet', '0');
			$objWriter->writeAttribute('minimized', '0');
			$objWriter->writeAttribute('showHorizontalScroll', '1');
			$objWriter->writeAttribute('showSheetTabs', '1');
			$objWriter->writeAttribute('showVerticalScroll', '1');
			$objWriter->writeAttribute('tabRatio', '600');
			$objWriter->writeAttribute('visibility', 'visible');

			$objWriter->endElement();

		$objWriter->endElement();
	}

	/**
	 * Write WorkbookProtection
	 *
	 * @param 	DATAPHPExcel_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	DATAPHPExcel					$pDATAPHPExcel
	 * @throws 	DATAPHPExcel_Writer_Exception
	 */
	private function _writeWorkbookProtection(DATAPHPExcel_Shared_XMLWriter $objWriter = null, DATAPHPExcel $pDATAPHPExcel = null)
	{
		if ($pDATAPHPExcel->getSecurity()->isSecurityEnabled()) {
			$objWriter->startElement('workbookProtection');
			$objWriter->writeAttribute('lockRevision',		($pDATAPHPExcel->getSecurity()->getLockRevision() ? 'true' : 'false'));
			$objWriter->writeAttribute('lockStructure', 	($pDATAPHPExcel->getSecurity()->getLockStructure() ? 'true' : 'false'));
			$objWriter->writeAttribute('lockWindows', 		($pDATAPHPExcel->getSecurity()->getLockWindows() ? 'true' : 'false'));

			if ($pDATAPHPExcel->getSecurity()->getRevisionsPassword() != '') {
				$objWriter->writeAttribute('revisionsPassword',	$pDATAPHPExcel->getSecurity()->getRevisionsPassword());
			}

			if ($pDATAPHPExcel->getSecurity()->getWorkbookPassword() != '') {
				$objWriter->writeAttribute('workbookPassword',	$pDATAPHPExcel->getSecurity()->getWorkbookPassword());
			}

			$objWriter->endElement();
		}
	}

	/**
	 * Write calcPr
	 *
	 * @param 	DATAPHPExcel_Shared_XMLWriter	$objWriter		XML Writer
	 * @param	boolean						$recalcRequired	Indicate whether formulas should be recalculated before writing
	 * @throws 	DATAPHPExcel_Writer_Exception
	 */
	private function _writeCalcPr(DATAPHPExcel_Shared_XMLWriter $objWriter = null, $recalcRequired = TRUE)
	{
		$objWriter->startElement('calcPr');

		$objWriter->writeAttribute('calcId', 			'124519');
		$objWriter->writeAttribute('calcMode', 			'auto');
		//	fullCalcOnLoad isn't needed if we've recalculating for the save
		$objWriter->writeAttribute('fullCalcOnLoad', 	($recalcRequired) ? '0' : '1');

		$objWriter->endElement();
	}

	/**
	 * Write sheets
	 *
	 * @param 	DATAPHPExcel_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	DATAPHPExcel					$pDATAPHPExcel
	 * @throws 	DATAPHPExcel_Writer_Exception
	 */
	private function _writeSheets(DATAPHPExcel_Shared_XMLWriter $objWriter = null, DATAPHPExcel $pDATAPHPExcel = null)
	{
		// Write sheets
		$objWriter->startElement('sheets');
		$sheetCount = $pDATAPHPExcel->getSheetCount();
		for ($i = 0; $i < $sheetCount; ++$i) {
			// sheet
			$this->_writeSheet(
				$objWriter,
				$pDATAPHPExcel->getSheet($i)->getTitle(),
				($i + 1),
				($i + 1 + 3),
				$pDATAPHPExcel->getSheet($i)->getSheetState()
			);
		}

		$objWriter->endElement();
	}

	/**
	 * Write sheet
	 *
	 * @param 	DATAPHPExcel_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	string 						$pSheetname 		Sheet name
	 * @param 	int							$pSheetId	 		Sheet id
	 * @param 	int							$pRelId				Relationship ID
	 * @param   string                      $sheetState         Sheet state (visible, hidden, veryHidden)
	 * @throws 	DATAPHPExcel_Writer_Exception
	 */
	private function _writeSheet(DATAPHPExcel_Shared_XMLWriter $objWriter = null, $pSheetname = '', $pSheetId = 1, $pRelId = 1, $sheetState = 'visible')
	{
		if ($pSheetname != '') {
			// Write sheet
			$objWriter->startElement('sheet');
			$objWriter->writeAttribute('name', 		$pSheetname);
			$objWriter->writeAttribute('sheetId', 	$pSheetId);
			if ($sheetState != 'visible' && $sheetState != '') {
				$objWriter->writeAttribute('state', $sheetState);
			}
			$objWriter->writeAttribute('r:id', 		'rId' . $pRelId);
			$objWriter->endElement();
		} else {
			throw new DATAPHPExcel_Writer_Exception("Invalid parameters passed.");
		}
	}

	/**
	 * Write Defined Names
	 *
	 * @param 	DATAPHPExcel_Shared_XMLWriter	$objWriter 		XML Writer
	 * @param 	DATAPHPExcel					$pDATAPHPExcel
	 * @throws 	DATAPHPExcel_Writer_Exception
	 */
	private function _writeDefinedNames(DATAPHPExcel_Shared_XMLWriter $objWriter = null, DATAPHPExcel $pDATAPHPExcel = null)
	{
		// Write defined names
		$objWriter->startElement('definedNames');

		// Named ranges
		if (count($pDATAPHPExcel->getNamedRanges()) > 0) {
			// Named ranges
			$this->_writeNamedRanges($objWriter, $pDATAPHPExcel);
		}

		// Other defined names
		$sheetCount = $pDATAPHPExcel->getSheetCount();
		for ($i = 0; $i < $sheetCount; ++$i) {
			// definedName for autoFilter
			$this->_writeDefinedNameForAutofilter($objWriter, $pDATAPHPExcel->getSheet($i), $i);

			// definedName for Print_Titles
			$this->_writeDefinedNameForPrintTitles($objWriter, $pDATAPHPExcel->getSheet($i), $i);

			// definedName for Print_Area
			$this->_writeDefinedNameForPrintArea($objWriter, $pDATAPHPExcel->getSheet($i), $i);
		}

		$objWriter->endElement();
	}

	/**
	 * Write named ranges
	 *
	 * @param 	DATAPHPExcel_Shared_XMLWriter	$objWriter 		XML Writer
	 * @param 	DATAPHPExcel					$pDATAPHPExcel
	 * @throws 	DATAPHPExcel_Writer_Exception
	 */
	private function _writeNamedRanges(DATAPHPExcel_Shared_XMLWriter $objWriter = null, DATAPHPExcel $pDATAPHPExcel)
	{
		// Loop named ranges
		$namedRanges = $pDATAPHPExcel->getNamedRanges();
		foreach ($namedRanges as $namedRange) {
			$this->_writeDefinedNameForNamedRange($objWriter, $namedRange);
		}
	}

	/**
	 * Write Defined Name for named range
	 *
	 * @param 	DATAPHPExcel_Shared_XMLWriter	$objWriter 		XML Writer
	 * @param 	DATAPHPExcel_NamedRange			$pNamedRange
	 * @throws 	DATAPHPExcel_Writer_Exception
	 */
	private function _writeDefinedNameForNamedRange(DATAPHPExcel_Shared_XMLWriter $objWriter = null, DATAPHPExcel_NamedRange $pNamedRange)
	{
		// definedName for named range
		$objWriter->startElement('definedName');
		$objWriter->writeAttribute('name',			$pNamedRange->getName());
		if ($pNamedRange->getLocalOnly()) {
			$objWriter->writeAttribute('localSheetId',	$pNamedRange->getScope()->getParent()->getIndex($pNamedRange->getScope()));
		}

		// Create absolute coordinate and write as raw text
		$range = DATAPHPExcel_Cell::splitRange($pNamedRange->getRange());
		for ($i = 0; $i < count($range); $i++) {
			$range[$i][0] = '\'' . str_replace("'", "''", $pNamedRange->getWorksheet()->getTitle()) . '\'!' . DATAPHPExcel_Cell::absoluteReference($range[$i][0]);
			if (isset($range[$i][1])) {
				$range[$i][1] = DATAPHPExcel_Cell::absoluteReference($range[$i][1]);
			}
		}
		$range = DATAPHPExcel_Cell::buildRange($range);

		$objWriter->writeRawData($range);

		$objWriter->endElement();
	}

	/**
	 * Write Defined Name for autoFilter
	 *
	 * @param 	DATAPHPExcel_Shared_XMLWriter	$objWriter 		XML Writer
	 * @param 	DATAPHPExcel_Worksheet			$pSheet
	 * @param 	int							$pSheetId
	 * @throws 	DATAPHPExcel_Writer_Exception
	 */
	private function _writeDefinedNameForAutofilter(DATAPHPExcel_Shared_XMLWriter $objWriter = null, DATAPHPExcel_Worksheet $pSheet = null, $pSheetId = 0)
	{
		// definedName for autoFilter
		$autoFilterRange = $pSheet->getAutoFilter()->getRange();
		if (!empty($autoFilterRange)) {
			$objWriter->startElement('definedName');
			$objWriter->writeAttribute('name',			'_xlnm._FilterDatabase');
			$objWriter->writeAttribute('localSheetId',	$pSheetId);
			$objWriter->writeAttribute('hidden',		'1');

			// Create absolute coordinate and write as raw text
			$range = DATAPHPExcel_Cell::splitRange($autoFilterRange);
			$range = $range[0];
			//	Strip any worksheet ref so we can make the cell ref absolute
			if (strpos($range[0],'!') !== false) {
				list($ws,$range[0]) = explode('!',$range[0]);
			}

			$range[0] = DATAPHPExcel_Cell::absoluteCoordinate($range[0]);
			$range[1] = DATAPHPExcel_Cell::absoluteCoordinate($range[1]);
			$range = implode(':', $range);

			$objWriter->writeRawData('\'' . str_replace("'", "''", $pSheet->getTitle()) . '\'!' . $range);

			$objWriter->endElement();
		}
	}

	/**
	 * Write Defined Name for PrintTitles
	 *
	 * @param 	DATAPHPExcel_Shared_XMLWriter	$objWriter 		XML Writer
	 * @param 	DATAPHPExcel_Worksheet			$pSheet
	 * @param 	int							$pSheetId
	 * @throws 	DATAPHPExcel_Writer_Exception
	 */
	private function _writeDefinedNameForPrintTitles(DATAPHPExcel_Shared_XMLWriter $objWriter = null, DATAPHPExcel_Worksheet $pSheet = null, $pSheetId = 0)
	{
		// definedName for PrintTitles
		if ($pSheet->getPageSetup()->isColumnsToRepeatAtLeftSet() || $pSheet->getPageSetup()->isRowsToRepeatAtTopSet()) {
			$objWriter->startElement('definedName');
			$objWriter->writeAttribute('name',			'_xlnm.Print_Titles');
			$objWriter->writeAttribute('localSheetId',	$pSheetId);

			// Setting string
			$settingString = '';

			// Columns to repeat
			if ($pSheet->getPageSetup()->isColumnsToRepeatAtLeftSet()) {
				$repeat = $pSheet->getPageSetup()->getColumnsToRepeatAtLeft();

				$settingString .= '\'' . str_replace("'", "''", $pSheet->getTitle()) . '\'!$' . $repeat[0] . ':$' . $repeat[1];
			}

			// Rows to repeat
			if ($pSheet->getPageSetup()->isRowsToRepeatAtTopSet()) {
				if ($pSheet->getPageSetup()->isColumnsToRepeatAtLeftSet()) {
					$settingString .= ',';
				}

				$repeat = $pSheet->getPageSetup()->getRowsToRepeatAtTop();

				$settingString .= '\'' . str_replace("'", "''", $pSheet->getTitle()) . '\'!$' . $repeat[0] . ':$' . $repeat[1];
			}

			$objWriter->writeRawData($settingString);

			$objWriter->endElement();
		}
	}

	/**
	 * Write Defined Name for PrintTitles
	 *
	 * @param 	DATAPHPExcel_Shared_XMLWriter	$objWriter 		XML Writer
	 * @param 	DATAPHPExcel_Worksheet			$pSheet
	 * @param 	int							$pSheetId
	 * @throws 	DATAPHPExcel_Writer_Exception
	 */
	private function _writeDefinedNameForPrintArea(DATAPHPExcel_Shared_XMLWriter $objWriter = null, DATAPHPExcel_Worksheet $pSheet = null, $pSheetId = 0)
	{
		// definedName for PrintArea
		if ($pSheet->getPageSetup()->isPrintAreaSet()) {
			$objWriter->startElement('definedName');
			$objWriter->writeAttribute('name',			'_xlnm.Print_Area');
			$objWriter->writeAttribute('localSheetId',	$pSheetId);

			// Setting string
			$settingString = '';

			// Print area
			$printArea = DATAPHPExcel_Cell::splitRange($pSheet->getPageSetup()->getPrintArea());

			$chunks = array();
			foreach ($printArea as $printAreaRect) {
				$printAreaRect[0] = DATAPHPExcel_Cell::absoluteReference($printAreaRect[0]);
				$printAreaRect[1] = DATAPHPExcel_Cell::absoluteReference($printAreaRect[1]);
				$chunks[] = '\'' . str_replace("'", "''", $pSheet->getTitle()) . '\'!' . implode(':', $printAreaRect);
			}

			$objWriter->writeRawData(implode(',', $chunks));

			$objWriter->endElement();
		}
	}
}

<?php
/**
 *  DATAPHPExcel
 *
 *  Copyright (c) 2006 - 2013 DATAPHPExcel
 *
 *  This library is free software; you can redistribute it and/or
 *  modify it under the terms of the GNU Lesser General Public
 *  License as published by the Free Software Foundation; either
 *  version 2.1 of the License, or (at your option) any later version.
 *
 *  This library is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 *  Lesser General Public License for more details.
 *
 *  You should have received a copy of the GNU Lesser General Public
 *  License along with this library; if not, write to the Free Software
 *  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 *  @category    DATAPHPExcel
 *  @package     DATAPHPExcel_Writer_PDF
 *  @copyright   Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 *  @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 *  @version     1.7.9, 2013-06-02
 */


/**  Require DomPDF library */
$pdfRendererClassFile = DATAPHPExcel_Settings::getPdfRendererPath() . '/dompdf_config.inc.php';
if (file_exists($pdfRendererClassFile)) {
    require_once $pdfRendererClassFile;
} else {
    throw new DATAPHPExcel_Writer_Exception('Unable to load PDF Rendering library');
}

/**
 *  DATAPHPExcel_Writer_PDF_DomPDF
 *
 *  @category    DATAPHPExcel
 *  @package     DATAPHPExcel_Writer_PDF
 *  @copyright   Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 */
class DATAPHPExcel_Writer_PDF_DomPDF extends DATAPHPExcel_Writer_PDF_Core implements DATAPHPExcel_Writer_IWriter
{
    /**
     *  Create a new DATAPHPExcel_Writer_PDF
     *
     *  @param   DATAPHPExcel    $phpExcel    DATAPHPExcel object
     */
    public function __construct(DATAPHPExcel $phpExcel)
    {
        parent::__construct($phpExcel);
    }

    /**
     *  Save DATAPHPExcel to file
     *
     *  @param   string     $pFilename   Name of the file to save as
     *  @throws  DATAPHPExcel_Writer_Exception
     */
    public function save($pFilename = NULL)
    {
        $fileHandle = parent::prepareForSave($pFilename);

        //  Default PDF paper size
        $paperSize = 'LETTER';    //    Letter    (8.5 in. by 11 in.)

        //  Check for paper size and page orientation
        if (is_null($this->getSheetIndex())) {
            $orientation = ($this->_phpExcel->getSheet(0)->getPageSetup()->getOrientation()
                == DATAPHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE)
                    ? 'L'
                    : 'P';
            $printPaperSize = $this->_phpExcel->getSheet(0)->getPageSetup()->getPaperSize();
            $printMargins = $this->_phpExcel->getSheet(0)->getPageMargins();
        } else {
            $orientation = ($this->_phpExcel->getSheet($this->getSheetIndex())->getPageSetup()->getOrientation()
                == DATAPHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE)
                    ? 'L'
                    : 'P';
            $printPaperSize = $this->_phpExcel->getSheet($this->getSheetIndex())->getPageSetup()->getPaperSize();
            $printMargins = $this->_phpExcel->getSheet($this->getSheetIndex())->getPageMargins();
        }

        //  Override Page Orientation
        if (!is_null($this->getOrientation())) {
            $orientation = ($this->getOrientation() == DATAPHPExcel_Worksheet_PageSetup::ORIENTATION_DEFAULT)
                ? DATAPHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT
                : $this->getOrientation();
        }
        //  Override Paper Size
        if (!is_null($this->getPaperSize())) {
            $printPaperSize = $this->getPaperSize();
        }

        if (isset(self::$_paperSizes[$printPaperSize])) {
            $paperSize = self::$_paperSizes[$printPaperSize];
        }

        $orientation = ($orientation == 'L') ? 'landscape' : 'portrait';

        //  Create PDF
        $pdf = new DOMPDF();
        $pdf->set_paper(strtolower($paperSize), $orientation);

        $pdf->load_html(
            $this->generateHTMLHeader(FALSE) .
            $this->generateSheetData() .
            $this->generateHTMLFooter()
        );
        $pdf->render();

        //  Write to file
        fwrite($fileHandle, $pdf->output());

		parent::restoreStateAfterSave($fileHandle);
    }

}

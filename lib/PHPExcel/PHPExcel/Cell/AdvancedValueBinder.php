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
 * @package    DATAPHPExcel_Cell
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
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
 * DATAPHPExcel_Cell_AdvancedValueBinder
 *
 * @category   DATAPHPExcel
 * @package    DATAPHPExcel_Cell
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 */
class DATAPHPExcel_Cell_AdvancedValueBinder extends DATAPHPExcel_Cell_DefaultValueBinder implements DATAPHPExcel_Cell_IValueBinder
{
    /**
     * Bind value to a cell
     *
     * @param  DATAPHPExcel_Cell  $cell  Cell to bind value to
     * @param  mixed $value          Value to bind in cell
     * @return boolean
     */
    public function bindValue(DATAPHPExcel_Cell $cell, $value = null)
    {
        // sanitize UTF-8 strings
        if (is_string($value)) {
            $value = DATAPHPExcel_Shared_String::SanitizeUTF8($value);
        }

        // Find out data type
        $dataType = parent::dataTypeForValue($value);

        // Style logic - strings
        if ($dataType === DATAPHPExcel_Cell_DataType::TYPE_STRING && !$value instanceof DATAPHPExcel_RichText) {
            //    Test for booleans using locale-setting
            if ($value == DATAPHPExcel_Calculation::getTRUE()) {
                $cell->setValueExplicit( TRUE, DATAPHPExcel_Cell_DataType::TYPE_BOOL);
                return true;
            } elseif($value == DATAPHPExcel_Calculation::getFALSE()) {
                $cell->setValueExplicit( FALSE, DATAPHPExcel_Cell_DataType::TYPE_BOOL);
                return true;
            }

            // Check for number in scientific format
            if (preg_match('/^'.DATAPHPExcel_Calculation::CALCULATION_REGEXP_NUMBER.'$/', $value)) {
                $cell->setValueExplicit( (float) $value, DATAPHPExcel_Cell_DataType::TYPE_NUMERIC);
                return true;
            }

            // Check for fraction
            if (preg_match('/^([+-]?) *([0-9]*)\s?\/\s*([0-9]*)$/', $value, $matches)) {
                // Convert value to number
                $value = $matches[2] / $matches[3];
                if ($matches[1] == '-') $value = 0 - $value;
                $cell->setValueExplicit( (float) $value, DATAPHPExcel_Cell_DataType::TYPE_NUMERIC);
                // Set style
                $cell->getWorksheet()->getStyle( $cell->getCoordinate() )
                    ->getNumberFormat()->setFormatCode( '??/??' );
                return true;
            } elseif (preg_match('/^([+-]?)([0-9]*) +([0-9]*)\s?\/\s*([0-9]*)$/', $value, $matches)) {
                // Convert value to number
                $value = $matches[2] + ($matches[3] / $matches[4]);
                if ($matches[1] == '-') $value = 0 - $value;
                $cell->setValueExplicit( (float) $value, DATAPHPExcel_Cell_DataType::TYPE_NUMERIC);
                // Set style
                $cell->getWorksheet()->getStyle( $cell->getCoordinate() )
                    ->getNumberFormat()->setFormatCode( '# ??/??' );
                return true;
            }

            // Check for percentage
            if (preg_match('/^\-?[0-9]*\.?[0-9]*\s?\%$/', $value)) {
                // Convert value to number
                $value = (float) str_replace('%', '', $value) / 100;
                $cell->setValueExplicit( $value, DATAPHPExcel_Cell_DataType::TYPE_NUMERIC);
                // Set style
                $cell->getWorksheet()->getStyle( $cell->getCoordinate() )
                    ->getNumberFormat()->setFormatCode( DATAPHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 );
                return true;
            }

            // Check for currency
            $currencyCode = DATAPHPExcel_Shared_String::getCurrencyCode();
            $decimalSeparator = DATAPHPExcel_Shared_String::getDecimalSeparator();
            $thousandsSeparator = DATAPHPExcel_Shared_String::getThousandsSeparator();
            if (preg_match('/^'.preg_quote($currencyCode).' *(\d{1,3}('.preg_quote($thousandsSeparator).'\d{3})*|(\d+))('.preg_quote($decimalSeparator).'\d{2})?$/', $value)) {
                // Convert value to number
                $value = (float) trim(str_replace(array($currencyCode, $thousandsSeparator, $decimalSeparator), array('', '', '.'), $value));
                $cell->setValueExplicit( $value, DATAPHPExcel_Cell_DataType::TYPE_NUMERIC);
                // Set style
                $cell->getWorksheet()->getStyle( $cell->getCoordinate() )
                    ->getNumberFormat()->setFormatCode(
                        str_replace('$', $currencyCode, DATAPHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD_SIMPLE )
                    );
                return true;
            } elseif (preg_match('/^\$ *(\d{1,3}(\,\d{3})*|(\d+))(\.\d{2})?$/', $value)) {
                // Convert value to number
                $value = (float) trim(str_replace(array('$',','), '', $value));
                $cell->setValueExplicit( $value, DATAPHPExcel_Cell_DataType::TYPE_NUMERIC);
                // Set style
                $cell->getWorksheet()->getStyle( $cell->getCoordinate() )
                    ->getNumberFormat()->setFormatCode( DATAPHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD_SIMPLE );
                return true;
            }

            // Check for time without seconds e.g. '9:45', '09:45'
            if (preg_match('/^(\d|[0-1]\d|2[0-3]):[0-5]\d$/', $value)) {
                // Convert value to number
                list($h, $m) = explode(':', $value);
                $days = $h / 24 + $m / 1440;
                $cell->setValueExplicit($days, DATAPHPExcel_Cell_DataType::TYPE_NUMERIC);
                // Set style
                $cell->getWorksheet()->getStyle( $cell->getCoordinate() )
                    ->getNumberFormat()->setFormatCode( DATAPHPExcel_Style_NumberFormat::FORMAT_DATE_TIME3 );
                return true;
            }

            // Check for time with seconds '9:45:59', '09:45:59'
            if (preg_match('/^(\d|[0-1]\d|2[0-3]):[0-5]\d:[0-5]\d$/', $value)) {
                // Convert value to number
                list($h, $m, $s) = explode(':', $value);
                $days = $h / 24 + $m / 1440 + $s / 86400;
                // Convert value to number
                $cell->setValueExplicit($days, DATAPHPExcel_Cell_DataType::TYPE_NUMERIC);
                // Set style
                $cell->getWorksheet()->getStyle( $cell->getCoordinate() )
                    ->getNumberFormat()->setFormatCode( DATAPHPExcel_Style_NumberFormat::FORMAT_DATE_TIME4 );
                return true;
            }

            // Check for datetime, e.g. '2008-12-31', '2008-12-31 15:59', '2008-12-31 15:59:10'
            if (($d = DATAPHPExcel_Shared_Date::stringToExcel($value)) !== false) {
                // Convert value to number
                $cell->setValueExplicit($d, DATAPHPExcel_Cell_DataType::TYPE_NUMERIC);
                // Determine style. Either there is a time part or not. Look for ':'
                if (strpos($value, ':') !== false) {
                    $formatCode = 'yyyy-mm-dd h:mm';
                } else {
                    $formatCode = 'yyyy-mm-dd';
                }
                $cell->getWorksheet()->getStyle( $cell->getCoordinate() )
                    ->getNumberFormat()->setFormatCode($formatCode);
                return true;
            }

            // Check for newline character "\n"
            if (strpos($value, "\n") !== FALSE) {
                $value = DATAPHPExcel_Shared_String::SanitizeUTF8($value);
                $cell->setValueExplicit($value, DATAPHPExcel_Cell_DataType::TYPE_STRING);
                // Set style
                $cell->getWorksheet()->getStyle( $cell->getCoordinate() )
                    ->getAlignment()->setWrapText(TRUE);
                return true;
            }
        }

        // Not bound yet? Use parent...
        return parent::bindValue($cell, $value);
    }
}

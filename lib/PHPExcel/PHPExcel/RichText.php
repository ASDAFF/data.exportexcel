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
 * @package    DATAPHPExcel_RichText
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    1.7.9, 2013-06-02
 */


/**
 * DATAPHPExcel_RichText
 *
 * @category   DATAPHPExcel
 * @package    DATAPHPExcel_RichText
 * @copyright  Copyright (c) 2006 - 2013 DATAPHPExcel (http://www.codeplex.com/DATAPHPExcel)
 */
class DATAPHPExcel_RichText implements DATAPHPExcel_IComparable
{
    /**
     * Rich text elements
     *
     * @var DATAPHPExcel_RichText_ITextElement[]
     */
    private $_richTextElements;

    /**
     * Create a new DATAPHPExcel_RichText instance
     *
     * @param DATAPHPExcel_Cell $pCell
     * @throws DATAPHPExcel_Exception
     */
    public function __construct(DATAPHPExcel_Cell $pCell = null)
    {
        // Initialise variables
        $this->_richTextElements = array();

        // Rich-Text string attached to cell?
        if ($pCell !== NULL) {
            // Add cell text and style
            if ($pCell->getValue() != "") {
                $objRun = new DATAPHPExcel_RichText_Run($pCell->getValue());
                $objRun->setFont(clone $pCell->getParent()->getStyle($pCell->getCoordinate())->getFont());
                $this->addText($objRun);
            }

            // Set parent value
            $pCell->setValueExplicit($this, DATAPHPExcel_Cell_DataType::TYPE_STRING);
        }
    }

    /**
     * Add text
     *
     * @param DATAPHPExcel_RichText_ITextElement $pText Rich text element
     * @throws DATAPHPExcel_Exception
     * @return DATAPHPExcel_RichText
     */
    public function addText(DATAPHPExcel_RichText_ITextElement $pText = null)
    {
        $this->_richTextElements[] = $pText;
        return $this;
    }

    /**
     * Create text
     *
     * @param string $pText Text
     * @return DATAPHPExcel_RichText_TextElement
     * @throws DATAPHPExcel_Exception
     */
    public function createText($pText = '')
    {
        $objText = new DATAPHPExcel_RichText_TextElement($pText);
        $this->addText($objText);
        return $objText;
    }

    /**
     * Create text run
     *
     * @param string $pText Text
     * @return DATAPHPExcel_RichText_Run
     * @throws DATAPHPExcel_Exception
     */
    public function createTextRun($pText = '')
    {
        $objText = new DATAPHPExcel_RichText_Run($pText);
        $this->addText($objText);
        return $objText;
    }

    /**
     * Get plain text
     *
     * @return string
     */
    public function getPlainText()
    {
        // Return value
        $returnValue = '';

        // Loop through all DATAPHPExcel_RichText_ITextElement
        foreach ($this->_richTextElements as $text) {
            $returnValue .= $text->getText();
        }

        // Return
        return $returnValue;
    }

    /**
     * Convert to string
     *
     * @return string
     */
    public function __toString()
    {
        return $this->getPlainText();
    }

    /**
     * Get Rich Text elements
     *
     * @return DATAPHPExcel_RichText_ITextElement[]
     */
    public function getRichTextElements()
    {
        return $this->_richTextElements;
    }

    /**
     * Set Rich Text elements
     *
     * @param DATAPHPExcel_RichText_ITextElement[] $pElements Array of elements
     * @throws DATAPHPExcel_Exception
     * @return DATAPHPExcel_RichText
     */
    public function setRichTextElements($pElements = null)
    {
        if (is_array($pElements)) {
            $this->_richTextElements = $pElements;
        } else {
            throw new DATAPHPExcel_Exception("Invalid DATAPHPExcel_RichText_ITextElement[] array passed.");
        }
        return $this;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        $hashElements = '';
        foreach ($this->_richTextElements as $element) {
            $hashElements .= $element->getHashCode();
        }

        return md5(
              $hashElements
            . __CLASS__
        );
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        $vars = get_object_vars($this);
        foreach ($vars as $key => $value) {
            if (is_object($value)) {
                $this->$key = clone $value;
            } else {
                $this->$key = $value;
            }
        }
    }
}

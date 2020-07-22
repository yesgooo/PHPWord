<?php
/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @see         https://github.com/PHPOffice/PHPWord
 * @copyright   2010-2018 PHPWord contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpWord\Writer\HTML\Element;

/**
 * Table element HTML writer
 *
 * @since 0.10.0
 */
class Table extends AbstractElement
{
    /**
     * Write table
     *
     * @return string
     */
    public function write()
    {
        if (!$this->element instanceof \PhpOffice\PhpWord\Element\Table) {
            return '';
        }

        $content = '';
        $rows = $this->element->getRows();
        $rowCount = count($rows);
        if ($rowCount > 0) {
            $content .= '<table' . self::getTableStyle($this->element->getStyle()) . '>' . PHP_EOL;
            for ($i = 0; $i < $rowCount; $i++) {
                /** @var $row \PhpOffice\PhpWord\Element\Row Type hint */
                $rowStyle = $rows[$i]->getStyle();
                $rowHeight = $rows[$i]->getHeight()/14;
                $tblHeader = $rowStyle->isTblHeader();
                $content .= '<tr style="height:'.$rowHeight.'px">' . PHP_EOL;
                $rowCells = $rows[$i]->getCells();
                $rowCellCount = count($rowCells);
                for ($j = 0; $j < $rowCellCount; $j++) {
                    $cellStyle = $rowCells[$j]->getStyle();
                    $hasBorder = $rowCells[$j]->getHasBorder();
                    $cellWidth = $cellStyle->getWidth()/14;
                    $cellBgColor = $cellStyle->getBgColor();
                    $cellFgColor = null;
                    if ($cellBgColor) {
                        $red = hexdec(substr($cellBgColor, 0, 2));
                        $green = hexdec(substr($cellBgColor, 2, 2));
                        $blue = hexdec(substr($cellBgColor, 4, 2));
                        $cellFgColor = (($red * 0.299 + $green * 0.587 + $blue * 0.114) > 186) ? null : 'ffffff';
                    }
                    $cellColSpan = $cellStyle->getGridSpan();
                    $cellRowSpan = 1;
                    $cellVMerge = $cellStyle->getVMerge();
                    // If this is the first cell of the vertical merge, find out how man rows it spans
                    if ($cellVMerge === 'restart') {
                        for ($k = $i + 1; $k < $rowCount; $k++) {
                            $kRowCells = $rows[$k]->getCells();
                            if (isset($kRowCells[$j])) {
                                if ($kRowCells[$j]->getStyle()->getVMerge() === 'continue') {
                                    $cellRowSpan++;
                                } else {
                                    break;
                                }
                            } else {
                                break;
                            }
                        }
                    }
                    // Ignore cells that are merged vertically with previous rows
                    // 边框 颜色
                    $borderSizes = $cellStyle->getBorderSize();
                    $borderColors = $cellStyle->getBorderStyle();
                    if ($cellVMerge !== 'continue') {
                        $cellTag = $tblHeader ? 'th' : 'td';
                        $cellColSpanAttr = (is_numeric($cellColSpan) && ($cellColSpan > 1) ? " colspan=\"{$cellColSpan}\"" : '');
                        $cellRowSpanAttr = ($cellRowSpan > 1 ? " rowspan=\"{$cellRowSpan}\"" : '');
                        $cellBgColorAttr = ((is_null($cellBgColor) || $cellBgColor == 'auto') ? '' : " bgcolor=\"#{$cellBgColor}\"");
                        $cellFgColorAttr = ((is_null($cellFgColor) || $cellFgColor == 'auto') ? '' : " color=\"#{$cellFgColor}\"");
                        $cellBorderColorAttr = ' style="width:'.$cellWidth.'px;';
                        $hasBorder && $cellBorderColorAttr .= 'border-top :'.($borderSizes[0]/4).'px solid '.((empty($borderColors[0]) || $borderColors[0] == 'auto') ? 'black' : '#'.$borderColors[0]);
                        $hasBorder &&  $cellBorderColorAttr .= ';border-left :'.($borderSizes[1]/4).'px solid '.((empty($borderColors[1]) || $borderColors[1] == 'auto') ? 'black' : '#'.$borderColors[1]);
                        $hasBorder && $cellBorderColorAttr .= ';border-right :'.($borderSizes[2]/4).'px solid '.((empty($borderColors[2]) || $borderColors[2] == 'auto') ? 'black' : '#'.$borderColors[2]);
                        $hasBorder && $cellBorderColorAttr .= ';border-bottom :'.($borderSizes[3]/4).'px solid '.((empty($borderColors[3]) || $borderColors[3] == 'auto') ? 'black' : '#'.$borderColors[3]);
                        $cellBorderColorAttr .= '"';
                        $content .= "<{$cellTag}{$cellColSpanAttr}{$cellRowSpanAttr}{$cellBgColorAttr}{$cellFgColorAttr}{$cellBorderColorAttr}>" . PHP_EOL;
                        $writer = new Container($this->parentWriter, $rowCells[$j]);
                        $content .= $writer->write();
                        if ($cellRowSpan > 1) {
                            // There shouldn't be any content in the subsequent merged cells, but lets check anyway
                            for ($k = $i + 1; $k < $rowCount; $k++) {
                                $kRowCells = $rows[$k]->getCells();
                                if (isset($kRowCells[$j])) {
                                    if ($kRowCells[$j]->getStyle()->getVMerge() === 'continue') {
                                        $writer = new Container($this->parentWriter, $kRowCells[$j]);
                                        $contents = $writer->write();
                                        if ($contents == '<p>&nbsp;</p>'.PHP_EOL) {
                                            break;
                                        }
                                        $content .= $contents;
                                    } else {
                                        break;
                                    }
                                } else {
                                    break;
                                }
                            }
                        }
                        $content .= "</{$cellTag}>" . PHP_EOL;
                    }
                }
                $content .= '</tr>' . PHP_EOL;
            }
            $content .= '</table>' . PHP_EOL;
        }

        return $content;
    }

    /**
     * Translates Table style in CSS equivalent
     *
     * @param string|\PhpOffice\PhpWord\Style\Table|null $tableStyle
     * @return string
     */
    private function getTableStyle($tableStyle = null)
    {
        if ($tableStyle == null) {
            return '';
        }
        if (is_string($tableStyle)) {
            $style = ' class="wordTable' . $tableStyle;
        } else {
            $style = ' style="';
            if ($tableStyle->getLayout() == \PhpOffice\PhpWord\Style\Table::LAYOUT_FIXED) {
                $style .= 'table-layout: fixed;';
            } elseif ($tableStyle->getLayout() == \PhpOffice\PhpWord\Style\Table::LAYOUT_AUTO) {
                $style .= 'table-layout: auto;';
            }
        }

        return $style . '"';
    }
}

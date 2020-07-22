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

namespace PhpOffice\PhpWord\Writer\HTML\Part;

use PhpOffice\PhpWord\Writer\HTML\Element\Container;
use PhpOffice\PhpWord\Writer\HTML\Element\TextRun as TextRunWriter;

/**
 * RTF body part writer
 *
 * @since 0.11.0
 */
class Body extends AbstractPart
{
    /**
     * Write part
     *
     * @return string
     */
    public function write()
    {
        $phpWord = $this->getParentWriter()->getPhpWord();

        $content = '';
        $sections = $phpWord->getSections();
        $height = $sections[0]->getStyle()->getPageSizeH()/17;
        $width = $sections[0]->getStyle()->getPageSizeW()/17;
        // 上下边距
        $topHeight = $sections[0]->getStyle()->getHeaderHeight()/17;
        $bottomHeight = $sections[0]->getStyle()->getFooterHeight()/17;
        // 获取表格css
        $css = $sections[0]->getTableCss();
        $content .= '<style>'.$css.'</style>';
        $content .= '<body style="'.'width: '.$width.'px;height:'.$height.'px">' . PHP_EOL;
        // 页眉
        $content .= '<div style="position: relative;margin-top: 100px;width: ' . $width .'">'.PHP_EOL;
        $header = $sections[0]->getHeaders();
        $writer = new Container($this->getParentWriter(), reset($header));
        $content .= $writer->write();
        $content .= '</div>';
        foreach ($sections as $section) {
            $writer = new Container($this->getParentWriter(), $section);
            $content .= $writer->write();
        }
        // 页脚
        $content .= '<div style="position: relative;width: ' . $width . '">'.PHP_EOL;
        $footer = $sections[0]->getFooters();
        $writer = new Container($this->getParentWriter(), reset($footer));
        $content .= $writer->write();
        $content .= '</div>';
        $content .= '</body>' . PHP_EOL;
        // 展示优化
        echo '&nbsp;';
        return $content;
    }

    /**
     * Write footnote/endnote contents as textruns
     *
     * @return string
     */
    private function writeNotes()
    {
        /** @var \PhpOffice\PhpWord\Writer\HTML $parentWriter Type hint */
        $parentWriter = $this->getParentWriter();
        $phpWord = $parentWriter->getPhpWord();
        $notes = $parentWriter->getNotes();

        $content = '';

        if (!empty($notes)) {
            $content .= '<hr />' . PHP_EOL;
            foreach ($notes as $noteId => $noteMark) {
                list($noteType, $noteTypeId) = explode('-', $noteMark);
                $method = 'get' . ($noteType == 'endnote' ? 'Endnotes' : 'Footnotes');
                $collection = $phpWord->$method()->getItems();

                if (isset($collection[$noteTypeId])) {
                    $element = $collection[$noteTypeId];
                    $noteAnchor = "<a name=\"note-{$noteId}\" />";
                    $noteAnchor .= "<a href=\"#{$noteMark}\" class=\"NoteRef\"><sup>{$noteId}</sup></a>";

                    $writer = new TextRunWriter($this->getParentWriter(), $element);
                    $writer->setOpeningText($noteAnchor);
                    $content .= $writer->write();
                }
            }
        }

        return $content;
    }
}

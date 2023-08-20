<?php

namespace MdhDigital\MdhExcel;

use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Common\Entity\Style\Border;
use OpenSpout\Common\Entity\Style\BorderPart;

trait MdhExcelStyles
{
      public function setStyles($styles)
      {
            $style = (new Style());

            $styles['bold']         != false    ? $style->setFontBold() : '';

            $styles['italic']       != false    ? $style->setFontItalic() : '';

            $styles['underline']    != false    ? $style->setFontUnderline() : '';

            $styles['wrap_text']    != false    ? $style->setShouldWrapText(true) : '';

            $styles['font_name']    != ''       ? $style->setFontName($styles['font_name']) : '';

            $styles['font_size']    != ''       ? $style->setFontSize((int)$styles['font_size']) : '';

            $styles['font_color']   != ''       ? $style->setFontColor($this->checkColor($styles['font_color'])) : '';

            $styles['bg_color']     != ''       ? $style->setBackgroundColor($this->checkColor($styles['bg_color'])) : '';

            $styles['alignment']    != ''       ? $style->setCellAlignment($styles['alignment']) : '';

            $styles['vertical_alignment']    != ''       ? $style->setCellVerticalAlignment($styles['vertical_alignment']) : '';


            return $style;
      }

      protected function checkColor($color)
      {

            if ($color == 'BLACK') {
                  return '000000';
            }

            if ($color == 'WHITE') {
                  return 'FFFFFF';
            }

            if ($color == 'RED') {
                  return 'FF0000';
            }

            if ($color == 'DARK_RED') {
                  return 'C00000';
            }

            if ($color == 'ORANGE') {
                  return 'FFC000';
            }

            if ($color == 'YELLOW') {
                  return 'FFFF00';
            }

            if ($color == 'LIGHT_GREEN') {
                  return '92D040';
            }

            if ($color == 'GREEN') {
                  return '00B050';
            }

            if ($color == 'LIGHT_BLUE') {
                  return '00B0E0';
            }

            if ($color == 'BLUE') {
                  return '0070C0';
            }

            if ($color == 'DARK_BLUE') {
                  return '002060';
            }

            if ($color == 'PURPLE') {
                  return '7030A0';
            }


            return $color;
      }
}

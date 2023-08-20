<?php

namespace MdhDigital\MdhExcel;

use OpenSpout\Writer\CSV\Options as CsvOptions;
use OpenSpout\Writer\XLSX\Options as XlsxOptions;

trait MdhExcelOptions
{
      /* CSV options */
      public static string $field_delimiter = ',';
      public static string $field_enclosure = '"';
      public static bool $should_add_bom = true;
      public static int $flush_threshold = 500;

      protected $customHeaderOptions;

      /* XLSX options */
      public static bool $should_use_inline_strings = true;

      public static function withoutInlineStrings(): self
      {
            self::$should_use_inline_strings = false;

            return self::getInstance();
      }

      public static function flushThreshold(int $threshold): self
      {
            self::$flush_threshold = $threshold;

            return self::getInstance();
      }

      public static function withoutBom(): self
      {
            self::$should_add_bom = false;

            return self::getInstance();
      }

      public static function useDelimiter(string $delimiter): self
      {
            self::$field_delimiter = $delimiter;

            return self::getInstance();
      }

      public function getCsvOptions(): CsvOptions
      {
            $options = new CsvOptions();
            $options->FIELD_DELIMITER = self::$field_delimiter;
            $options->FIELD_ENCLOSURE = self::$field_enclosure;
            $options->SHOULD_ADD_BOM = self::$should_add_bom;
            $options->FLUSH_THRESHOLD = self::$flush_threshold;

            $reset = new CsvOptions();
            self::$field_delimiter = $reset->FIELD_DELIMITER;
            self::$field_enclosure = $reset->FIELD_ENCLOSURE;
            self::$should_add_bom = $reset->SHOULD_ADD_BOM;
            self::$flush_threshold = $reset->FLUSH_THRESHOLD;

            return $options;
      }

      public function getXlsxOptions($header): XlsxOptions
      {
            $options = new XlsxOptions();
            $options->SHOULD_USE_INLINE_STRINGS = self::$should_use_inline_strings;

            $this->headerOptions($header);

            if ($this->customHeaderOptions !== false) {

                  $haderNumber = 1;

                  foreach ($this->customHeaderOptions as $header) {

                        (int)$header['width'] > 0 ? $options->setColumnWidth($header['width'], $haderNumber++) : $haderNumber++;
                  }
            }

            $reset = new XlsxOptions();
            self::$should_use_inline_strings = $reset->SHOULD_USE_INLINE_STRINGS;

            return $options;
      }

      public function headerOptions($header)
      {
            $header !== false ? $this->customHeaderOptions = $header : $this->customHeaderOptions = false;

            return $this;
      }
}

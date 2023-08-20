<?php

namespace MdhDigital\MdhExcel;

use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Writer\CSV\Options as CSVOptions;
use OpenSpout\Writer\CSV\Writer;
use OpenSpout\Writer\WriterInterface; 

/**
 * Basic MdhExcelCreation
 * 
 */

class MdhExcelProcess
{ 

      use MdhExcelOptions;
      
      private WriterInterface $creation;

      private bool $processHeader = true;

      private bool $processingFirstRow = true;

      private int $numberOfRows = 0;

      private ?Style $headerStyle = null;

      protected CSVOptions $csvOptions;

      public static function create(
            string $file,
            string $type = '',
            callable $configureWriter = null,
            ?string $delimiter = null,
            ?bool $shouldAddBom = null,
      ): static {
            $simpleExcelWriter = new static(
                  path: $file,
                  type: $type,
                  delimiter: $delimiter,
                  shouldAddBom: $shouldAddBom,
            );

            $writer = $simpleExcelWriter->getCreation();

            if ($configureWriter) {
                  $configureWriter($writer);
            }

            $writer->openToFile($file);

            return $simpleExcelWriter;
      }

      public static function createWithoutBom(string $file, string $type = ''): static
      {
            return static::create(
                  file: $file,
                  type: $type,
                  shouldAddBom: false,
            );
      }

      public static function stream(
            string $downloadName,
            string $type = '',
            callable $writerCallback = null,
            ?string $delimiter = null,
            ?bool $shouldAddBom = null,
      ): static {
            $simpleExcelWriter = new static($downloadName, $type, $delimiter, $shouldAddBom);

            $writer = $simpleExcelWriter->getCreation();

            if ($writerCallback) {
                  $writerCallback($writer);
            }

            $writer->openToBrowser($downloadName);

            return $simpleExcelWriter;
      } 

      protected function __construct(
            private string $path,
            string $type = '',
            ?string $delimiter = null,
            ?bool $shouldAddBom = null,
      ) {
            $this->csvOptions = new CSVOptions();

            $this->initWriter($path, $type);

            $this->addOptionsToWriter($path, $type, $delimiter, $shouldAddBom);
      }

      protected function initWriter(string $path, string $type, ?CSVOptions $options = null): void
      {
 
            $options = $this->getXlsxOptions(MdhExcelCreation::$customHeader);


            $this->creation = empty($type) ?
                  MdhExcelInitialization::createFromFile($path, $options) :
                  MdhExcelInitialization::createFromType($type, $options);
      } 

      protected function addOptionsToWriter(
            string $path,
            string $type = '',
            ?string $delimiter = null,
            ?bool $shouldAddBom = null,
      ): void {
            if (!$delimiter && $shouldAddBom) {
                  return;
            }

            if (!$this->creation instanceof Writer) {
                  return;
            }

            if ($delimiter !== null) {
                  $this->csvOptions->FIELD_DELIMITER = $delimiter;
            }

            if ($shouldAddBom !== null) {
                  $this->csvOptions->SHOULD_ADD_BOM = $shouldAddBom;
            }

            $this->initWriter($path, $type, $this->csvOptions);
      }

      public function getPath(): string
      {
            return $this->path;
      }

      public function getCreation(): WriterInterface
      {
            return $this->creation;
      }

      public function getNumberOfRows(): int
      {
            return $this->numberOfRows;
      }

      public function noHeaderRow(): static
      {
            $this->processHeader = false;

            return $this;
      }

      public function setHeaderStyle(Style $style): static
      {
            $this->headerStyle = $style;

            return $this;
      }

      public function addToRowExcel(Row|array $row, Style $style = null): static
      {
            if (is_array($row)) {
                  if ($this->processHeader && $this->processingFirstRow) {
                        $this->writeHeaderFromRow($row);
                  }

                  $row = Row::fromValues($row, $style);
            }

            $this->creation->addRow($row);
            $this->numberOfRows++;

            $this->processingFirstRow = false;

            return $this;
      }

      public function addToRowExcels(iterable $rows, Style $style = null): static
      {
            foreach ($rows as $row) {
                  $this->addToRowExcel($row, $style);
            }

            return $this;
      }

      public function addHeader(array $header): self
      {
            $headerRow = Row::fromValues($header, $this->headerStyle);

            $this->creation->addRow($headerRow);
            $this->numberOfRows++;

            $this->processingFirstRow = false;

            return $this;
      }

      protected function writeHeaderFromRow(array $row): void
      {
            $headerValues = array_keys($row);

            $headerRow = Row::fromValues($headerValues, $this->headerStyle);

            $this->creation->addRow($headerRow);
            $this->numberOfRows++;
      }


      public function toBrowser()
      {
            $this->creation->close();

            exit;
      }

      public function close(): void
      {
            $this->creation->close();
      }

      public function __destruct()
      {
            $this->close();
      }
}

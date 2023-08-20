<?php

namespace MdhDigital\MdhExcel;

/**
 * Basic MdhExcelCreation
 * 
 */

class MdhExcelCreation
{
      use MdhExcelOptions, MdhExcelStyles;

      protected object $collectionData;

      public static $customHeader;

      protected string $fileName;

      protected int $chunk;

      protected $creation;

      protected $styles;

      public function download(array $params)
      {

            $this->headerData($params['custom_header']);

            $this->customFileName($params['file_name']);

            $this->headerOptions($params['custom_header']);

            $this->headerData($params['custom_header']);

            $this->loadPerData($params['chunk']);

            $this->loadOurData($params['data']);

            $this->customHeaderStyle($params['header_style']);


            return $this->process();
      }


      private function process()
      {
 
            $this->creation = MdhExcelProcess::stream($this->fileName);

            $this->styles != false ?  $this->creation->setHeaderStyle($this->setStyles($this->styles)) : '';

            $this->processFromData();

            return $this->creation->toBrowser();
      }


      protected function customFileName($fileName)
      {
            $this->fileName = $fileName;

            return $this;
      }

      protected function headerData($header)
      {
            $header['status'] == false ? self::$customHeader = false : self::$customHeader = collect($header['header']);

            return $this;
      }

      protected function loadPerData(int $chunk)
      {
            $this->chunk            = $chunk;

            return $this;
      }

      protected function loadOurData(object $data)
      {
            $this->collectionData   = $data;
            return $this;
      }

      protected function processFromData()
      {
            $this->collectionData->chunk($this->chunk, function ($ourCollection) {
                  foreach ($ourCollection as $v) {
                        self::$customHeader == false ? $data = $this->getDataFromOriginal($v->getOriginal()) : $data = $this->getDataFromCustom($v);

                        $this->creation->addToRowExcel($data);
                  }
                  flush();
            });
      }

      protected function getDataFromOriginal(array $data)
      {

            foreach ($data as $d => $value) {
                  $i[$d]      = $value;
            }

            $listData = $i;

            return $listData;
      }

      protected function getDataFromCustom($data)
      {
            foreach (self::$customHeader as $d) {

                  if (count(explode('.', $d['value'])) > 1) {

                        $relationKey = $this->validationOfRelation($d['value']);
                        $value = eval('return $data->' . $relationKey . ' ?? "";');

                        $i[$d['label']]         = $d['type'] == 'int' ? (int)$value : $value;
                  }

                  if (count(explode('.', $d['value'])) == 1) {
                        $i[$d['label']]         = $d['type'] == 'int' ? (int)$data->getOriginal($d['value']) : $data->getOriginal($d['value']);
                  }
            }

            $listData   = $i;
            return $listData;
      }

      protected function validationOfRelation($value)
      {
            $parts = explode('.', $value);

            $parts_length = count($parts);

            $result = '';
            for ($i = 0; $i < $parts_length; $i++) {
                  if ($i == 0) {

                        $result .= $parts[$i];
                  } else {

                        $result .= '->' . $parts[$i];
                  }
            }

            return $result;
      }

      protected function customHeaderStyle($styles)
      {
            $styles['status'] == false ? $this->styles = false : $this->styles = collect($styles['attribut']);

            return $this;
      }
}

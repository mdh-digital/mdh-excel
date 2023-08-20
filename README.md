## MDHEXCEL - LARAVEL EXPORT EXCEL AND CSV FOR MANY DATA 
 

## Installation 

### Laravel
You can install it using the composer command

    composer require mdh-digital/mdh-excel

### Example Use Code
 

  ```
  use MdhDigital\MdhExcel\MdhExcelCreation;
$data = $this->query(); // Your Query or Data
$download   = new MdhExcelCreation();

        return $download->download([
            'file_name'         => 'download_md.xlsx',
            'data'              => $data,
            'chunk'             => 2500,
            'custom_header'     => array(
                'status'            => true,
                'header'            => [
                    array(
                        'label'     => 'Date',
                        'value'     => 'excel_date',
                        'type'      => 'string',
                        'width'     => 20,
                    ),
                    array(
                        'label'     => 'Ref No',
                        'value'     => 'transaction.ref_no',
                        'type'      => 'string',
                        'width'     => 20,
                    ),
                    array(
                        'label'     => 'Store',
                        'value'     => 'store.name',
                        'type'      => 'string',
                        'width'     => 20,
                    ),
                    array(
                        'label'     => 'Product Name',
                        'value'     => 'full_name',
                        'type'      => 'string',
                        'width'     => 20,
                    ),
                    array(
                        'label'     => 'Sell Price',
                        'value'     => 'excel_price',
                        'type'      => 'string',
                        'width'     => 20,
                    ),
                    array(
                        'label'     => 'Qty',
                        'value'     => 'qty',
                        'type'      => 'int',
                        'width'     => 20,
                    ),
                   
                    array(
                        'label'     => 'Subtotal',
                        'value'     => 'excel_subtotal',
                        'type'      => 'string',
                        'width'     => 20,
                    ),
                ]
            ),
            'header_style'  => array(
                'status'            => true,
                'attribut'          => array(
                    'bold'          => true,
                    'italic'        => false,
                    'underline'     => false,
                    'font_name'     => '',
                    'font_size'     => 15,
                    'font_color'    => 'WHITE',
                    'alignment'     => 'center',
                    'vertical_alignment'    => 'center',
                    'wrap_text'     => false,
                    'bg_color'      => 'LIGHT_BLUE'
                )
            ),
            'body_style' => array(
                'status'    => false
            )
        ]);
  ```
  
### License

This Mdhexcel Wrapper for Laravel is open-sourced software licensed under the [MIT license](http://opensource.org/licenses/MIT)

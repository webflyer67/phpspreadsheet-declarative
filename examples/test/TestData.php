<?php
/**
 * Класс TestData - тестовый массив с данными
 *
 * User: Alexander Kudryavtsev
 * Date: 22.10.2018
 */


namespace webflyer67\PhpspreadsheetDeclarative\Test;

use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;

class TestData
{

    /**  Возвращает массив с тестовыми данныи выборочного размера
     * @return array
     */
    public static function getData($size = 10)
    {
        $data = [];
        for ($i = 1; $i <= $size; $i++) {
            $data[$i] = [
                'number' => $i,
                'format' => '60x90',
                'href' => 'http://localhost',
                'count' => rand(1, 10),
                'type' => 'стикер  ',
            ];
            $data[$i]['price_our'] = rand(30, 100) * 100;
            $data[$i]['price'] = $data[$i]['price_our'] * 1.2;
            $data[$i]['total_price'] = $data[$i]['price'] * $data[$i]['count']; //'=SUM(E5:F5)';//'=SUM(RC[-2]:RC[-1])';
            $data[$i]['merge'] = (int)(($i - 1) / 10 + 1);
            if ($data[$i]['total_price'] < 20000) {
                $data[$i]['styles'] = 'green';
            }
            if ($data[$i]['total_price'] > 80000) {
                $data[$i]['styles'] = 'red';
            }
        }
        return $data;
    }

    /**  Возвращает массив с тестовыми  картинками
     * @return array
     */
    public static function getDataImages()
    {
        return [
            [
                'height' => 100,
                'img1' => [
                    'Path' => 'https://dummyimage.com/300x200/1291bb/fff.jpg&text=phpspreadsheet',
                    'Name' => 'Дерево',
                    'Description' => 'Дерево',
                    'OffsetX' => 0,
                    'OffsetY' => 0,
                    // 'Rotation' => -5,
                    'Hyperlink' => 'http://localhost',
                    'Height' => 130, // px
                    // 'Width' => 300,// px
                    //  'ResizeProportional' => false,
                    'Shadow' => ['Visible' => true, 'Direction' => 45, 'Distance' => 6],
                    //  'height' => 100,
                ],
                'img2' => [
                    'Path' => 'https://dummyimage.com/300x200/1291bb/fff.jpg&text=decrarative',
                    'Name' => 'Лес',
                    'Description' => 'Лес',
                    'OffsetX' => 0,
                    'OffsetY' => 0,
                    // 'Rotation' => 5,
                    'Hyperlink' => 'http://localhost',
                    'Height' => 130,// px
                    'Shadow' => ['Visible' => true, 'Direction' => 90,],
                ],
                'img3' => [
                    'Path' => 'https://dummyimage.com/300x200/1291bb/fff.jpg&text=phpspreadsheet',
                    'Name' => 'Олень',
                    'Description' => 'Олень',
                    'OffsetX' => 20,
                    'OffsetY' => 0,
                    // 'Rotation' => -5,
                    'Hyperlink' => 'http://localhost',
                    'Height' => 130,// px
                    'Shadow' => ['Visible' => true, 'Direction' => 135,],
                ],

            ],

            [
                'height' => 100,
                'img1' => [
                    'Path' => $_SERVER['DOCUMENT_ROOT'] . '/examples/images/3.jpg',
                    'Name' => 'Вокзал',
                    'Description' => 'Вокзал',
                    'OffsetX' => 0, // %
                    'OffsetY' => 0, // %
                    //'Rotation' => 5,// градусы
                    'Hyperlink' => 'http://localhost',
                    'Height' => 130,// px
                    'Shadow' => ['Visible' => true, 'Direction' => 315,],
                ],
                'img2' => [
                    'Path' => $_SERVER['DOCUMENT_ROOT'] . '/examples/images/4.jpg',
                    'Name' => 'Жук',
                    'Description' => 'Жук',
                    'OffsetX' => 0,
                    'OffsetY' => 0,
                    //  'Rotation' => -5,
                    'Hyperlink' => 'http://localhost',
                    'Height' => 130,// px
                    'Shadow' => ['Visible' => true, 'Direction' => 270,],
                ],
                'img3' => [
                    'Path' => $_SERVER['DOCUMENT_ROOT'] . '/examples/images/6.jpg',
                    'Name' => 'Памятник',
                    'Description' => 'Памятник',
                    'OffsetX' => 20,
                    'OffsetY' => 0,
                    //   'Rotation' => 5,
                    'Hyperlink' => 'http://localhost',
                    'Height' => 130,// px
                    'Shadow' => ['Visible' => true, 'Direction' => 225,],
                ],
            ],


        ];
    }

    public static function getMeta()
    {
        return [
            'Creator' => 'Vasilii Pupkin',
            'LastModifiedBy' => 'Vasilii Pupkin',
            'Title' => 'Test PhpspreadsheetDeclarative',
            'Subject' => 'Test PhpspreadsheetDeclarative',
            'Description' => 'Test PhpspreadsheetDeclarative',
            'Keywords' => 'PhpspreadsheetDeclarative, php, Phpspreadsheet, spreadsheet',
            'Category' => 'test spreadsheet',
            'Company' => 'webflyer67',
        ];
    }

    public static function getSetup()
    {
        return [
            'Orientation' => PageSetup::ORIENTATION_LANDSCAPE,
            'PaperSize' => PageSetup::PAPERSIZE_A4,
        ];
    }
}

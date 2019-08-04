<?php
/**
 * Добавление картинок, нет заголовка таблицы, пропуск столбцов
 */

require_once $_SERVER['DOCUMENT_ROOT'] . '/vendor/autoload.php';
use webflyer67\PhpspreadsheetDeclarative\Writer;

/** @var array Массив с табличными данными */
$images = [
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
        
        
      
        
        /** @var array Массив с шаблоном для генерации таблицы */
        $template =  [
            'sheetCaption' => 'Прайс',
            'tables' => [
                [             
                    'bindTable' => 'images',
                    'columns' => [
                        [
                            'body' => [
                                'width' => 40,
                                'bindImage' => 'img1', 
                                'bindHeight' => 'height'
                            ],
                        ],
                        [], [],
                        [
                            'body' => [
                                'width' => 40,
                                'bindImage' => 'img2',
                            ],
                        ],
                        [],
                        [
                            'body' => [
                                'width' => 40,
                                'bindImage' => 'img3',
                            ],
                        ]                        
                    ]
                ]
            ]
        ];
$fileName = 'example 11 ' . date("m.d.y H_i_s");
$fileNameFull = $_SERVER['DOCUMENT_ROOT'] . '/runtime/' . $fileName;
Writer::getWriter() // создание экземпляра объекта (новый xls документ)
    ->addData('images', $images) // привязка массива с данными
    ->addSheet($template, $pageSetup)   // добавление листа
    ->writeDocument($fileNameFull . '.xlsx') // сохранение на диск Word 2007
    ->writeDocument($fileNameFull . '.pdf'); // сохранение на диск PDF


require_once 'lib.php';
toLog('Файлы сохранены в '. $fileNameFull . '.*');

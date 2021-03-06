<?php
/**
 * Высота строк, ширина колонок
 */

require_once $_SERVER['DOCUMENT_ROOT'] . '/vendor/autoload.php';
use webflyer67\PhpspreadsheetDeclarative\Writer;

/** @var array Массив с табличными данными */
$users = [
    ['id' => 1, 'name' => 'Alex', 'age' => '15', 'group' => 'admin'],
    ['id' => 2, 'name' => 'John', 'age' => '45', 'group' => 'admin'],
    ['id' => 3, 'name' => 'Bill', 'age' => '16', 'group' => 'user', 'rowHeight' => 30,],
    ['id' => 4, 'name' => 'Jimm', 'age' => '31', 'group' => 'user', 'rowHeight' => 30,],
];

/** @var array Массив с шаблоном для генерации таблицы */
$template = [
    'sheetCaption' => 'Пользователи', // Название листа
    'tables' => [
        [
            'bindTable' => 'users', // название связанного массива с данными
            'styles' => [  // Глобальные стили для всей таблицы
                'all' => 'border', // стили для всей таблицы
            ],
            'columns' => [ // заголовки столбцов и привязанные к ним данные
                [
                    'head' => [// заголовок
                        [
                            'caption' => 'id пользователя', // текст в заголовке
                            'width' => 30, // ширина столбца(em)
                            'height' => 30, // высота строки(pt)
                        ],
                    ],
                    'body' => [ // тело
                        'bindColumn' => 'id' // привязанное значение из 'bindTable' => 'users'
                    ],
                ],
                [
                    'head' => [
                        [
                            'caption' => 'Имя пользователя'
                        ],
                    ],
                    'body' => [
                        'bindColumn' => 'name',
                        'width' => 30, // ширина столбца(em) 
                        'bindHeight' => 'rowHeight'// высота строки(pt)
                    ],
                ],
            ]
        ],
        
    ]
];

use PhpOffice\PhpSpreadsheet\Style\Border;
/** @var array Массив со стилями */
$styles = [
    'border' => [
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN],],
    ],
];

$fileName = 'example 06 ' . date("m.d.y H_i_s");
$fileNameFull = $_SERVER['DOCUMENT_ROOT'] . '/runtime/' . $fileName;
Writer::getWriter() // создание экземпляра объекта (новый xls документ)
    ->addData('users', $users) // привязка массива с данными
    ->addStyles($styles)// Добавление стилей
    ->addSheet($template, $pageSetup)   // добавление листа
    ->writeDocument($fileNameFull . '.xlsx'); // сохранение на диск Word 2007


require_once 'lib.php';
toLog('Файлы сохранены в '. $fileNameFull . '.*');

<?php
/**
 * Добавление стилей
 */

require_once $_SERVER['DOCUMENT_ROOT'] . '/vendor/autoload.php';
use webflyer67\PhpspreadsheetDeclarative\Writer;

/** @var array Массив с табличными данными */
$users = [
    ['id' => 1, 'name' => 'Alex', 'age' => '15', 'group' => 'admin'],
    ['id' => 2, 'name' => 'John', 'age' => '45', 'group' => 'admin', 'cellStyles' => 'right'],
    ['id' => 3, 'name' => 'Bill', 'age' => '16', 'group' => 'user'],
    ['id' => 4, 'name' => 'Jimm', 'age' => '31', 'group' => 'user'],
];

/** @var array Массив с шаблоном для генерации таблицы */
$template = [
    'sheetCaption' => 'Пользователи', // Название листа
    'tables' => [
        [
            'bindTable' => 'users', // название связанного массива с данными
            'styles' => [  // Глобальные стили для всей таблицы
                'all' => 'border', // стили для всей таблицы
                'head' => ['primary-bg-color', 'primary-font', 'center'], // стили для всех заголовков
                'body' => 'center' // стили для тела таблицы
            ],
            'columns' => [ // заголовки столбцов и привязанные к ним данные
                [
                    'head' => [// заголовок
                        [
                            'caption' => 'id пользователя', // текст в заголовке
                            'styles' => ['left'], // Стили для текущего заголовка
                            'width' => 30, // ширина столбца(em)
                        ],
                    ],
                    'body' => [ // тело
                        'bindColumn' => 'id' // привязанное значение из 'bindTable' => 'users'
                    ],
                ],
                [
                    'head' => [
                        [
                            'caption' => 'Имя пользователя',
                            'width' => 30, // ширина столбца(em)
                         ],
                    ],
                    'body' => [
                        'bindColumn' => 'name',
                        'styles' => ['left'], // Стили для текущего столбца
                        'bindStyles' => 'cellStyles' // привязка стилей для отдельных ячеек тела таблицы
                    ],
                ],
            ]
        ],
        
    ]
];

use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Font;
/** @var array Массив со стилями */
$styles = [
    'border' => [
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN],],
    ],
    'primary-bg-color' => [
        'fill' => ['fillType' => Fill::FILL_SOLID, 'color' => ['rgb' => '2980B9']],
    ],
    'primary-font' => [
        'font' => ['bold' => true, 'color' => ['rgb' => 'ffffff'], 'size' => 10, 'name' => 'Arial'],
    ],
    'center' => [
        'alignment' => ['horizontal' => 'center', 'vertical' => 'center', 'wrap' => true, 'shrinkToFit' => true],
    ],
    'left' => [
        'alignment' => ['horizontal' => 'left'],
    ],
    'right' => [
        'alignment' => ['horizontal' => 'right'],
    ],
];
$fileName = 'example 08 ' . date("m.d.y H_i_s");
$fileNameFull = $_SERVER['DOCUMENT_ROOT'] . '/runtime/' . $fileName;
Writer::getWriter() // создание экземпляра объекта (новый xls документ)
    ->addData('users', $users) // привязка массива с данными
    ->addStyles($styles)// Добавление стилей
    ->addSheet($template, $pageSetup)   // добавление листа
    ->writeDocument($fileNameFull . '.xlsx') // сохранение на диск Word 2007
    ->writeDocument($fileNameFull . '.pdf'); // сохранение на диск PDF


require_once 'lib.php';
toLog('Файлы сохранены в '. $fileNameFull . '.*');

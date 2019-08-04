<?php
/**
 * Объединение ячеек
 */

require_once $_SERVER['DOCUMENT_ROOT'] . '/vendor/autoload.php';
use webflyer67\PhpspreadsheetDeclarative\Writer;

/** @var array Массив с табличными данными */
$users = [
    ['id' => 1, 'name' => 'Alex', 'age' => '15', 'group' => 'admin'],
    ['id' => 2, 'name' => 'John', 'age' => '45', 'group' => 'admin'],
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
            ],
            'columns' => [ // заголовки столбцов и привязанные к ним данные
                [
                    'head' => [// заголовок
                        [
                            'caption' => 'id пользователя', // текст в заголовке
                            'width' => 30, // ширина столбца(em)
                            'height' => 30, // высота строки(pt)
                            'mergeId' => 'id'
                        ],
                        [
                            'mergeId' => 'id'
                        ],
                    ],
                    'body' => [ // тело
                        'bindColumn' => 'id' // привязанное значение из 'bindTable' => 'users'
                    ],
                ],
                [
                    'head' => [
                        [
                            'caption' => ' Пользователь',
                            'mergeId' => 'group', // идентификатор объединения
                        ],
                        [
                            'caption' => 'Имя',
                        ],
                    ],
                    'body' => [
                        'bindColumn' => 'name',
                        'width' => 30, // ширина столбца(em) 
                    ],
                ],
                [
                    'head' => [
                        [
                            'mergeId' => 'group',
                        ],
                        [
                            'caption' => 'Группа',
                        ],
                    ],
                    'body' => [
                        'bindColumn' => 'group',
                        'bindMerge' => 'group', // поле, которое является идентификатором объединения
                        'width' => 30, // ширина столбца(em)
                    ],
                ],
            ]
        ],
        
    ]
];

use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
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
];

$fileName = 'example 10 ' . date("m.d.y H_i_s");
$fileNameFull = $_SERVER['DOCUMENT_ROOT'] . '/runtime/' . $fileName;
Writer::getWriter() // создание экземпляра объекта (новый xls документ)
    ->addData('users', $users) // привязка массива с данными
    ->addStyles($styles)// Добавление стилей
    ->addSheet($template, $pageSetup)   // добавление листа
    ->writeDocument($fileNameFull . '.xlsx'); // сохранение на диск Word 2007


require_once 'lib.php';
toLog('Файлы сохранены в '. $fileNameFull . '.*');

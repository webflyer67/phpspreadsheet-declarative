<?php
/**
 * Добавление метаданных файла
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
    'sheetCaption' => 'Пользователи',
    'tables' => [
        [
            'bindTable' => 'users',
            'columns' => [
                [
                    'head' => [
                        ['caption' => 'id пользователя'],
                    ],
                    'body' => ['bindColumn' => 'id'],
                ],
                [
                    'head' => [
                        ['caption' => 'Имя пользователя'],
                    ],
                    'body' => ['bindColumn' => 'name'],
                ],
            ]
        ],
        
    ]
];
$meta = [
    'Creator' => 'Vasilii Pupkin',
    'LastModifiedBy' => 'Vasilii Pupkin',
    'Title' => 'Test PhpspreadsheetDeclarative',
    'Subject' => 'Test PhpspreadsheetDeclarative',
    'Description' => 'Test PhpspreadsheetDeclarative',
    'Keywords' => 'PhpspreadsheetDeclarative, php, Phpspreadsheet, spreadsheet',
    'Category' => 'test spreadsheet',
    'Company' => 'webflyer67',
];

$fileName = 'example 04 ' . date("m.d.y H_i_s");
$fileNameFull = $_SERVER['DOCUMENT_ROOT'] . '/runtime/' . $fileName;
Writer::getWriter() // создание экземпляра объекта (новый xls документ)
    ->setMeta($meta)// Добавление метаданных файла
    ->addData('users', $users) // привязка массива с данными
    ->addSheet($template, $pageSetup)   // добавление листа
    ->writeDocument($fileNameFull . '.xlsx') // сохранение на диск Word 2007
    ->writeDocument($fileNameFull . '.pdf'); // сохранение на диск PDF


require_once 'lib.php';
toLog('Файлы сохранены в '. $fileNameFull . '.*');

<?php
/**
 * Создание и сохранение документа в разных форматах, xlsx, xls, html, pdf(3 варианта)
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

$fileName = 'example 02 ' . date("m.d.y H_i_s");
$fileNameFull = $_SERVER['DOCUMENT_ROOT'] . '/runtime/' . $fileName;
Writer::getWriter() // создание экземпляра объекта (новый xls документ)
    ->addData('users', $users) // привязка массива с данными
    ->addSheet($template, $pageSetup)   // добавление листа

    ->writeDocument($fileNameFull . '.xlsx') // сохранение на диск Word 2007
    ->writeDocument($fileNameFull . '_m.pdf') // сохранение на диск PDF (MPDF)
    ->writeDocument($fileNameFull . '.xls') // сохранение на диск Word 2003
    ->writeDocument($fileNameFull . '.html') // сохранение на диск HTML
    ->writeDocument($fileNameFull . '_tc.pdf', 'tc') // сохранение на диск  PDF (TC PDF)
    ->writeDocument($fileNameFull . '_dom.pdf', 'dom'); // сохранение на диск  PDF (PDF Dom)



require_once 'lib.php';
toLog('Файлы сохранены в '. $fileNameFull . '.*');

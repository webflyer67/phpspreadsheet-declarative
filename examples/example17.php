<?php
/**
 * Доступ к объекту Spreadsheet и редактирование его напрямую
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
            'columns' => [ // заголовки столбцов и привязанные к ним данные
                [
                    'head' => [// заголовок
                        [
                            'caption' => 'id пользователя' // текст в заголовке
                        ],
                    ],
                    'body' => [ // тело
                        'bindColumn' => 'id' // привязанное значение из 'bindTable' => 'users'
                    ],
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

$fileName = 'example 17 ' . date("m.d.y H_i_s");
$fileNameFull = $_SERVER['DOCUMENT_ROOT'] . '/runtime/' . $fileName;
$spreadsheet = Writer::getWriter() // создание экземпляра объекта (новый xls документ)
    ->addData('users', $users) // привязка массива с данными
    ->addSheet($template);   // добавление листа
  
$spreadsheet->getDocument()
  ->getSheet(0)
  ->setCellValue('F1', 'Вставка данных через объект Spreadsheet')
  ->setCellValue('B6', 'Вставка данных через объект Spreadsheet');

$spreadsheet->writeDocument($fileNameFull . '.xlsx'); // сохранение на диск Word 2007
require_once 'lib.php';
toLog('Файлы сохранены в '. $fileNameFull . '.*');

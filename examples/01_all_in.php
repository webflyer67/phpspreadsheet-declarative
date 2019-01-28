<?php
/**
 * Создает документы разных форматов небольшого размера и сохраняет их на диск
 */
require_once 'lib.php';

use webflyer67\PhpspreadsheetDeclarative\Test\TestData;
use webflyer67\PhpspreadsheetDeclarative\Test\TestTemplates;
use webflyer67\PhpspreadsheetDeclarative\Test\TestStyles;

use webflyer67\PhpspreadsheetDeclarative\Writer;

$fileName = 'hello world ' . date("m.d.y H_i_s");
$fileNameFull = $_SERVER['DOCUMENT_ROOT'] . '/runtime/' . $fileName;
Writer::getWriter()// создание экземпляра FastXlsPdfHelper (новый xls документ)
->setMeta(TestData::getMeta())// Добавление метаданных файла
->addDatas([ // Добавление источников данных
    'prices' => TestData::getData(5),
    'images' => TestData::getDataImages()
])
    ->addStyles(TestStyles::getStyles())// Добавление стилей
    ->addSheet(TestTemplates::getTemplate1(), TestData::getSetup())//генерация первого листа
    ->addSheet(TestTemplates::getTemplate2(), TestData::getSetup())//генерация второго листа
// Созранение в разных форматах
    ->writeDocument($fileNameFull . '.xlsx')
    ->writeDocument($fileNameFull . '_m.pdf')
    ->writeDocument($fileNameFull . '.xls')
    ->writeDocument($fileNameFull . '.html')
    ->writeDocument($fileNameFull . '_tc.pdf', 'tc')
    ->writeDocument($fileNameFull . '_dom.pdf', 'dom');

toLog('Файлы сохранены в /runtime/...');



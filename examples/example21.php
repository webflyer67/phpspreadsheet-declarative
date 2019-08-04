<?php

require_once 'lib.php';

use webflyer67\PhpspreadsheetDeclarative\Test\TestData;
use webflyer67\PhpspreadsheetDeclarative\Test\TestTemplates;
use webflyer67\PhpspreadsheetDeclarative\Test\TestStyles;

use webflyer67\PhpspreadsheetDeclarative\Writer;


$spreadsheet = Writer::getWriter()// создание экземпляра FastXlsPdfHelper (новый xls документ)
->setMeta(TestData::getMeta())
//    ->addData('prices', TestData::getData(50))
//    ->addData('images', TestData::getDataImages())
    ->addDatas([
        'prices' => TestData::getData(3),
        'images' => TestData::getDataImages()
    ])
    ->addStyles(TestStyles::getStyles())
    ->addSheet(TestTemplates::getTemplate1(),TestData::getSetup());

$spreadsheet->getDocument()
    ->getActiveSheet()
    ->setCellValue('A1', 'Hello World !')
    ->setCellValue('B2', 'Hello World 2 !');
$spreadsheet->addSheet(TestTemplates::getTemplate2(),TestData::getSetup());
$fileName =  'example 21 ' . date("m.d.y H_i_s");
$fileNameFull = $_SERVER['DOCUMENT_ROOT'] . '/runtime/' . $fileName;
$spreadsheet->writeDocument($fileNameFull .  '.xlsx');
$spreadsheet->writeDocument($fileNameFull. '_m.pdf');


toLog('Файлы сохранены в '. $fileNameFull . '.*');


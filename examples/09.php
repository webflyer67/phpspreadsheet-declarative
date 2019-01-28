<?php

require_once 'lib.php';

use webflyer\PhpspreadsheetDeclarative\Test\TestData;
use webflyer\PhpspreadsheetDeclarative\Test\TestTemplates;
use webflyer\PhpspreadsheetDeclarative\Test\TestStyles;

use webflyer\PhpspreadsheetDeclarative\Writer;


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
$fileName = 'hello world ' . date("m.d.y H_i_s");

$spreadsheet->writeDocument($_SERVER['DOCUMENT_ROOT'] . '/runtime/' . $fileName . '.xlsx');
$spreadsheet->writeDocument($_SERVER['DOCUMENT_ROOT'] . '/runtime/' . $fileName . '_m.pdf');


//$spreadsheet->writeDocument($fileName . '.xls');
//$spreadsheet->writeDocument($_SERVER['DOCUMENT_ROOT'] . '/runtime/' . $fileName . '.html');
//$spreadsheet->writeDocument($_SERVER['DOCUMENT_ROOT'] . '/runtime/' . $fileName . '_tc.pdf','tc');
//$spreadsheet->writeDocument($_SERVER['DOCUMENT_ROOT'] . '/runtime/' . $fileName . '_dom.pdf','dom');
//$spreadsheet->sendDocument($fileName . '.pdf');

toLog([
    '01 !!!',

]);



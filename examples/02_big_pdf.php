<?php
/**
 * Создает документ xlsx большого размера и отправляет в браузер
 */
require_once 'lib.php';

use webflyer\PhpspreadsheetDeclarative\Test\TestData;
use webflyer\PhpspreadsheetDeclarative\Test\TestTemplates;
use webflyer\PhpspreadsheetDeclarative\Test\TestStyles;

use webflyer\PhpspreadsheetDeclarative\Writer;

$fileName = 'hello world ' . date("m.d.y H_i_s");

Writer::getWriter()// создание экземпляра FastXlsPdfHelper (новый xls документ)
->setMeta(TestData::getMeta())// Добавление метаданных файла
->addData('prices', TestData::getData(500))// Добавление источников данных
->addStyles(TestStyles::getStyles())// Добавление стилей
->addSheet(TestTemplates::getTemplate3(), TestData::getSetup())//генерация первого листа
->sendDocument($fileName . '.pdf');



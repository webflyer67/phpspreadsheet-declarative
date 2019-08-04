<?php
/**
 * Создает документ xlsx большого размера и отправляет в браузер
 */
require_once 'lib.php';

use webflyer67\PhpspreadsheetDeclarative\Test\TestData;
use webflyer67\PhpspreadsheetDeclarative\Test\TestTemplates;
use webflyer67\PhpspreadsheetDeclarative\Test\TestStyles;

use webflyer67\PhpspreadsheetDeclarative\Writer;

$fileName =  'example 20 ' . date("m.d.y H_i_s");

/** Writer   */
Writer::getWriter()// создание экземпляра FastXlsPdfHelper (новый xls документ)
->setMeta(TestData::getMeta())// Добавление метаданных файла
->addData('prices', TestData::getData(500))// Добавление источников данных
->addStyles(TestStyles::getStyles())// Добавление стилей
->addSheet(TestTemplates::getTemplate3(), TestData::getSetup())//генерация первого листа
->sendDocument($fileName . '.xlsx'); // Отправка в браузер




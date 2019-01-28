<?php
/**
 * Считывает xls в массив
 */
require_once 'lib.php';

use webflyer\PhpspreadsheetDeclarative\Reader;

$template = [
    'xls_link' => '^WWW$',
    'xls_code' => '^№$',
    'xls_side' => '^Сторона$',
];

/** @var string $fileNameFull Имя файла*/
$fileNameFull = $_SERVER['DOCUMENT_ROOT'] . '/examples/files/test.xls';


/** @var Reader $reader Парсер*/
$reader = Reader::getReader($fileNameFull);
$resultShort = $reader->parse($template);
$resultFull = $reader->parse($template, 'full');

toLog([
    'Парсинг завершен',
    '$resultShort' => $resultShort,
    '$resultFull' => $resultFull,
]);





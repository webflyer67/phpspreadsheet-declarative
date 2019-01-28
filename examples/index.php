<?php
require_once 'lib.php';
// формирование ссыло на все примеры
$fileList = glob("[0-9]*.php");
foreach ($fileList as $fileName) {
    echo '<a target="_blank" href="' . $fileName . '">' . $fileName . '</a><br>';
}
<?php
error_reporting(E_ALL);
ini_set('display_errors', 1);
ini_set("pcre.backtrack_limit", "5000000");

require_once $_SERVER['DOCUMENT_ROOT'] . '/vendor/autoload.php';

makeDir($_SERVER['DOCUMENT_ROOT'] . '/runtime');


function makeDir($dir)
{
    if (!is_dir($dir)) {
        mkdir($dir);
        chmod($dir, 777);
    }
}

/**
 * Пишет сообщение в лог  /runtime/__FILE__.log, можно передавать массив
 *
 * @param mixed $value Сообщение, или массив сообщений
 * @param int $die Флаг остановки скрипта
 * @param int $toHTML Флаг вывода в браузер
 */
function toLog($value, $die = 0, $toHTML = 1) //todo kill
{
    // if (!isset($_SESSION['auth']['user_id']) || $_SESSION['auth']['user_id'] != 79491)   return;
    if (!preg_match('/(\w+)\.php/is', __FILE__, $fn))
        $fn[1] = 'default_log';
    $msg = date("[d.m.Y  H:i:s]") . "[" . getmypid() . "]: " . print_r($value, 1);
    file_put_contents($_SERVER['DOCUMENT_ROOT'] . "/runtime/" . $fn[1] . ".log", $msg . "\n", FILE_APPEND);
    if ($toHTML)
        echo '<pre>' . $msg . "\n</pre>";
    if ($die) {
        if (defined('START_TIME_GLOBAL'))
            echo print_r('<pre>' . "\nВремя выполнения: " . round((microtime(true) - START_TIME_GLOBAL), 1), 1) . " сек.\n" . "\n</pre>";
        die();
    }
}


/**
 * Исправление проблем с совместимостью с Mpdf, перед генерацией документа вносим несклько изменений:
 * 1. В инлийн стилях убираем для границ "!important", иначе они не отрисовываются
 * 2. Для всех ссылок наследуем стили: цвет и подчеркивание, они применяются к родителю и по дефолту не наследуются
 */
/*$html = $this->generateHTMLHeader(false) .
    $this->generateSheetData() .
    $this->generateHTMLFooter();
$html = preg_replace('/(border[^<>;"\']+) !important/uis', '$1', $html);
$html = preg_replace('/<a /uis', '<a style="color: inherit;text-decoration: inherit;"', $html);
$pdf->WriteHTML($html);*/

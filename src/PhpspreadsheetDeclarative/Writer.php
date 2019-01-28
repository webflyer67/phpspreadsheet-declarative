<?php
/**
 * Класс Writer позволяет генерировать xls/pdf из шаблона и данных наподобие того как ангуляр строит страницу из шаблона и данных.
 * Пока может генерировать только простые таблицы в формате xls, позднее надо добавить поддержку pdf, и по xls тоже много чего добавить
 * XlsPdfHelper::buildDoc();// тестовая генерация
 *
 * User: Alexander Kudryavtsev
 * Date: 16.11.2018
 */

namespace webflyer67\PhpspreadsheetDeclarative;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\Hyperlink;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Writer\Html;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Dompdf;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Tcpdf;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


class  Writer
{
    /** @var Spreadsheet $document Объект PhpSpreadsheet */
    private $document;

    /** @var array $data Массив данных для таблиц */
    private $data = [];

    /** @var array $styles Массив стилей */
    private $styles = [];

    /** @var array $merge Массив с диапазонами ячеек для объединения */
    private $merge = [];

    /**
     * @var array $thousands Массив с координатами ячеек, у которых применен разделитель тысяч. Нужен для того чтоб при
     * сохранении в pdf разделить тысячи пробелом, поскольку pdf не понимает форматирование заданное через setFormatCode()
     */
    private $thousands = [];

    /**
     * Инициализирует и возвращает экземпляр данного класса
     *
     * @return $this
     */
    public static function getWriter()
    {
        $writer = new Writer();
        $writer->document = new Spreadsheet();
        return $writer;
    }

    /**
     * Устанавливает метаданные документа
     *
     * @param array $meta Метаданные документа: автор, название, описание и т.д.
     *
     * @return $this
     */
    public function setMeta($meta)
    {
        if (count($meta) > 0) {
            foreach ($meta as $prop => $val) {
                $method = 'set' . $prop;
                $this->document->getProperties()
                    ->$method($val);
            }
        }
        return $this;
    }

    /**
     * Добавляет массив из которого впоследствии будет сформировано тело таблицы
     *
     * @param string $name Имя массива
     * @param array $array Массив из элементо ключ-значение
     *
     * @return $this
     */
    public function addData($name, $array)
    {
        $this->data[$name] = $array;
        return $this;
    }

    /**
     * Обертка над addData, чтоб можно было одним массивом добавить несколько массивов данных
     *
     * @param array $array Массив массивов из элементов ключ-значение
     *
     * @return $this
     */
    public function addDatas($array)
    {
        foreach ($array as $dataName => $dataArray) {
            $this->addData($dataName, $dataArray);
        }
        return $this;
    }

    /**
     * Добавляет массива со стилями из которого впоследствии будут браться стили
     *
     * @param string $name Имя массива
     * @param array $array Массив из элементов в формате, котарый понимается методом applyFromArray()
     *
     * @return $this
     */
    public function addStyle($name, $array)
    {
        $this->styles[$name] = $array;
        return $this;
    }

    /**
     * Обертка над addStyle, чтоб можно было одним массивом добавить несколько массивов стилей
     *
     * @param array $array Массив массивов из элементов ключ-значение
     *
     * @return $this
     */
    public function addStyles($array)
    {
        foreach ($array as $styleName => $styleArray) {
            $this->addStyle($styleName, $styleArray);
        }
        return $this;
    }

    /**
     * Добавляет лист к документу. В переданном шаблоне установленв связи со стилями и данными, по этому шаблону строится лист
     *
     * @param array $template Шаблон генерируемого документа
     * @param array $setup Настройки листа: размер, ориентация
     *
     * @return $this
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function addSheet($template, $setup = [])
    {
        // Создание листов
        $this->document->setActiveSheetIndex($this->document->getSheetCount() - 1);
        $sheet = $this->document->getActiveSheet();

        $sheet->setShowGridLines(false);
        $this->setSetup($setup);

        // Формирование названия листа
        if (!empty($template['sheetCaption'])) {
            $sheet->setTitle($template['sheetCaption']);
        } else {
            $sheet->setTitle('Sheet ' . $this->document->getSheetCount());
        }
        if (is_array($template['tables']) && count($template['tables'])) {
            foreach ($template['tables'] as $tableTemplate) {
                $this->merge = [];
                // Создание таблиц
                $this->addTable($tableTemplate, $sheet);
                // Объединение ячеек
                $this->mergeCellsXLS($sheet);
            }
        }
        $sheet->setSelectedCell('A1');
        $this->document->createSheet();
        $this->document->setActiveSheetIndex($this->document->getSheetCount() - 1);
        return $this;
    }

    /**
     * Возвращает нативный объект для возможности вносить в него правки напрямую.
     * Для Writer - это Spreadsheet, для PdfWriter - это HTML
     *
     * @return Spreadsheet
     */
    public function getDocument()
    {
        return $this->document;
    }


    /**
     * Сохраняет файл на диск
     *
     * @param string $pFilename Полное имя файла
     * @param string $pdfType Выбрать какую библиотеку использовать для сохранения в pdf: 'm'|'dom'|'tc'
     *
     * @return $this
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @throws Exception
     */
    public function writeDocument($pFilename, $pdfType = 'm')
    {
        $this->document->setActiveSheetIndex(0);
        if ($this->document->getSheetCount() > 1) {
            $this->document->removeSheetByIndex($this->document->getSheetCount() - 1);
        }
        $ext = pathinfo($pFilename)['extension'];
        switch ($ext) {
            case 'xlsx':
                $writer = new Xlsx($this->document);
                break;
            case 'xls':
                $writer = new Xls($this->document);
                break;
            case 'html':
                $writer = new Html($this->document);
                break;
            case 'pdf':
                switch ($pdfType) {
                    case 'tc':
                        $writer = new Tcpdf($this->document);
                        break;
                    case 'dom':
                        $writer = new Dompdf($this->document);
                        break;
                    case 'm':
                    default:
                        $writer = new Mpdf($this->document);
                        break;
                }
                break;
            default:
                throw new Exception('Неверное расширение файла');
        }
        if ($ext == 'html' || $ext == 'pdf') {
            $this->correctThousands('pdf');
            $writer->writeAllSheets();
        }

        $writer->save($pFilename);

        if ($ext == 'html' || $ext == 'pdf') {
            $this->correctThousands('xls');
        }
        $this->document->createSheet();
        return $this;
    }

    /**
     * Отсылает файл в браузер
     *
     * @param string $filename - имя файла
     * @param string $pdfType Выбрать какую библиотеку использовать для сохранения в pdf: 'm'|'dom'|'tc'
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @throws Exception
     */
    public function sendDocument($filename, $pdfType = 'm')
    {
        $this->document->setActiveSheetIndex(0);
        if ($this->document->getSheetCount() > 1) {
            $this->document->removeSheetByIndex($this->document->getSheetCount() - 1);
        }
        $ext = pathinfo($filename)['extension'];
        switch ($ext) {
            case 'xlsx':
                $writer = new Xlsx($this->document);
                $mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
                break;
            case 'xls':
                $writer = new Xls($this->document);
                $mimeType = 'application/vnd.ms-excel';
                break;
            case 'html':
                $writer = new Html($this->document);
                $mimeType = 'text/html';
                break;
            case 'pdf':
                switch ($pdfType) {
                    case 'tc':
                        $writer = new Tcpdf($this->document);
                        break;
                    case 'dom':
                        $writer = new Dompdf($this->document);
                        break;
                    case 'm':
                    default:
                        $writer = new Mpdf($this->document);
                        break;
                }
                $mimeType = 'application/pdf';
                break;
            default:
                throw new Exception('Неверное расширение файла');
        }
        if ($ext == 'html' || $ext == 'pdf') {
            $this->correctThousands('pdf');
            $writer->writeAllSheets();
        }

        header('Content-Type: ' . $mimeType);
        header('Content-Disposition: attachment;filename="' . $filename . '"');
        header('Cache-Control: max-age=0');
        $writer->save('php://output');
        exit(1);
    }


    /**
     * Устанавливает настройки листа, размер бумаги, ориентацию
     *
     * @param array $setup Настройки
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function setSetup($setup)
    {
        if (!$setup && !is_array($setup)) {
            $setup = [
                'Orientation' => PageSetup::ORIENTATION_LANDSCAPE,
                'PaperSize' => PageSetup::PAPERSIZE_A4,
            ];
        }
        foreach ($setup as $prop => $val) {
            $method = 'set' . $prop;
            $this->document->getActiveSheet()->getPageSetup()
                ->$method($val);
        }
    }

    /**
     * Добавляет таблицу на лист
     *
     * @param array $template Шаблон листа
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet Лист - объект Worksheet
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function addTable($template, $sheet)
    {
        // Инициализация курсора и применение отступов
        // Отступ у колонок вседа с левого края
        $column = 1;
        if (!empty($template['marginLeft']) && $template['marginLeft'] > 0) {
            $column = $template['marginLeft'] + 1;
        }

        // Отступ у строк относительно уже занятых строк с учетом уже занятых трок и отступа
        $row = $sheet->getHighestDataRow();
        if ($row == 1 && $sheet->getCell($sheet->getHighestDataColumn() . '1')->getValue() == '') {
            $row = 0;
        }
        $row++;
        if (!empty($template['marginTop']) && $template['marginTop'] > 0) {
            $row += $template['marginTop'];
        }

        $bodyHeight = count($this->data[$template['bindTable']]);
        $tableHeight = $bodyHeight;
        $headHeight = 0;
        if (!empty($template['columns'][0]['head'])) {
            $headHeight = count($template['columns'][0]['head']);
        }

        $tableHeight += $headHeight;

        // Применение стилей stylesTableAll ко всей таблице
        if (!empty($template['styles']['all'])) {
            $cells = Coordinate::stringFromColumnIndex($column) . $row . ':'
                . Coordinate::stringFromColumnIndex($column + count($template['columns']) - 1) . ($row + $tableHeight - 1);
            $this->setStyles($cells, $template['styles']['all'], $sheet);
        }

        $startColumn = $column;
        $startRow = $row;

        // Создание шапки таблицы
        if (!empty($template['columns'][0]['head'])) {
            $this->addTableHead($template, $sheet, $headHeight, $startRow, $startColumn);
        }
        $startRow += $headHeight;

        // Создание тела таблицы
        $this->addTableBody($template, $sheet, $bodyHeight, $startRow, $startColumn);
    }

    /**
     * Добавляет заголовок таблицы
     *
     * @param array $template Шаблон листа
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet Лист - объект Worksheet
     * @param int $headHeight Высота заголовка
     * @param int $startRow Начальная строка
     * @param int $startColumn Начальный столбец
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function addTableHead($template, $sheet, $headHeight, $startRow, $startColumn)
    {
        // Шапка
        $row = $startRow;
        $column = $startColumn;
        // Применение общих стилей stylesTableHead ко всей шапке
        if (!empty($template['styles']['head'])) {
            $cells = Coordinate::stringFromColumnIndex($column) . $row . ':'
                . Coordinate::stringFromColumnIndex($column + count($template['columns']) - 1)
                . ($row + $headHeight - 1);
            $this->setStyles($cells, $template['styles']['head'], $sheet);
        }

        foreach ($template['columns'] as $colTemplate) {// Перебор шапки таблицы(колонки)
            $row = $startRow;
            foreach ($colTemplate['head'] as $headRow) {// Перебор шапки таблицы(ряды)

                $cells = Coordinate::stringFromColumnIndex($column) . $row;

                //Вывод заголовка
                if (!empty($headRow['caption'])) {
                    $sheet->setCellValue(Coordinate::stringFromColumnIndex($column) . $row, $headRow['caption']);
                }

                // Привязка ссылки, если есть
                if (!empty($headRow['url'])) {
                    $sheet->getCell($cells)->getHyperlink()->setUrl($headRow['url']);
                }

                // Применение уникальных для каждой ячейки шапки стилей stylesTableHead
                if (!empty($headRow['styles'])) {
                    $this->setStyles($cells, $headRow['styles'], $sheet);
                }

                // Применение ширины столбца
                if (!empty($headRow['width'])) {
                    $sheet->getColumnDimension(Coordinate::stringFromColumnIndex($column))->setWidth($headRow['width']);
                }

                // Применение высоты строки
                if (!empty($headRow['height'])) {
                    $sheet->getRowDimension($row)->setRowHeight($headRow['height']);
                }

                // Применение фильтров
                if (!empty ($headRow['filters'])) {
                    if (!is_array($headRow['filters'])) {
                        $headRow['filters'] = [$headRow['filters']];
                    }

                    // Форматирование суммы - разделитель тысяч
                    if (in_array('thousands', $headRow['filters'])) {
                        $sheet->getStyle($cells)->getNumberFormat()->setFormatCode('#,##0');
                        $this->thousands[$sheet->getParent()->getActiveSheetIndex()][] = $cells;
                    }
                }

                // Сбор данных об объединенных ячейках
                if (!empty($headRow['mergeId'])) {
                    $mergeId = $headRow['mergeId'];
                    $this->merge[$mergeId]['columns'][$column] = $column;
                    $this->merge[$mergeId]['rows'][$row] = $row;
                }
                $row++;
            }
            $column++;
        }
    }

    /**
     * Добавляет тело таблицы
     *
     * @param array $template Шаблон листа
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet Лист - объект Worksheet
     * @param int $bodyHeight Высота заголовка
     * @param int $startRow Начальная строка
     * @param int $startColumn Начальный столбец
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function addTableBody($template, $sheet, $bodyHeight, $startRow, $startColumn)
    {
        $row = $startRow;
        $column = $startColumn;
        // Применение общих стилей stylesTableBody ко всему телу
        if (!empty($template['styles']['body'])) {
            $cells = Coordinate::stringFromColumnIndex($column) . ($row) . ':' .
                Coordinate::stringFromColumnIndex($column + count($template['columns']) - 1) . ($row + $bodyHeight - 1);
            $this->setStyles($cells, $template['styles']['body'], $sheet);
        }

        // Создание тела таблицы
        foreach ($template['columns'] as $colTemplate) {// Перебор столбцов таблицы

            if (!empty($colTemplate['body'])) {


                $row = $startRow;

                $cells = Coordinate::stringFromColumnIndex($column) . $row . ':' . Coordinate::stringFromColumnIndex($column) . ($row + $bodyHeight - 1);

                // Применение стилей к телу таблицы, ко всему столбцу
                if (!empty($colTemplate['body']['styles'])) {
                    $this->setStyles($cells, $colTemplate['body']['styles'], $sheet);
                }

                // Применение ширины столбца
                if (!empty($colTemplate['body']['width'])) {
                    $sheet->getColumnDimension(Coordinate::stringFromColumnIndex($column))->setWidth($colTemplate['body']['width']);
                }

                // Применение фильтров
                if (!empty ($colTemplate['body']['filters'])) {
                    if (!is_array($colTemplate['body']['filters'])) {
                        $colTemplate['body']['filters'] = [$colTemplate['body']['filters']];
                    }

                    // Форматирование суммы - разделитель тысяч
                    if (in_array('thousands', $colTemplate['body']['filters'])) {
                        $sheet->getStyle($cells)->getNumberFormat()->setFormatCode('#,##0');
                        $this->thousands[$sheet->getParent()->getActiveSheetIndex()][] = $cells;
                    }
                }

                if (!empty($template['bindTable']) && !empty($this->data[$template['bindTable']])) {
                    foreach ($this->data[$template['bindTable']] as $dataRow) {// Перебор данных

                        $cell = Coordinate::stringFromColumnIndex($column) . $row;

                        //Вывод значения, или значения по умолчанию в тело таблицы
                        if (!empty($colTemplate['body']['bindColumn'])) {
                            if (isset($dataRow[$colTemplate['body']['bindColumn']])) {
                                $sheet->setCellValue($cell, $dataRow[$colTemplate['body']['bindColumn']]);
                            } elseif (isset($colTemplate['body']['defaultValue'])) {
                                $sheet->setCellValue($cell, $colTemplate['body']['defaultValue']);
                            }
                        }
                        if (!empty($colTemplate['body']['bindColumnImage'])) {
                            if (!empty($dataRow[$colTemplate['body']['bindColumnImage']])) {

                                $drawing = new Drawing();
                                $props = $dataRow[$colTemplate['body']['bindColumnImage']];
                                foreach ($props as $prop => $val) {
                                    $method = 'set' . $prop;
                                    if ($prop == 'Hyperlink') {
                                        $tooltip = '';
                                        if (!empty($props['Description'])) {
                                            $tooltip = $props['Description'];
                                        }
                                        $val = new Hyperlink($val, $tooltip);
                                    }
                                    if ($prop != 'Shadow') {
                                        $drawing->$method($val);
                                    } else {
                                        if (!empty($val)) {
                                            foreach ($val as $propShadow => $valShadow) {
                                                $method = 'set' . $propShadow;
                                                $drawing->getShadow()->$method($valShadow);
                                            }
                                        }
                                    }

                                }
                                $drawing->setCoordinates($cell);
                                $drawing->setWorksheet($sheet);
                            }
                        }
                        // Привязка ссылки, если есть
                        if (!empty($colTemplate['body']['bindColumnUrl']) && !empty($dataRow[$colTemplate['body']['bindColumnUrl']])) {
                            $sheet->getCell($cell)->getHyperlink()->setUrl($dataRow[$colTemplate['body']['bindColumnUrl']]);
                        }

                        // Применение высоты строки
                        if (!empty($colTemplate['body']['bindHeight']) && !empty($dataRow[$colTemplate['body']['bindHeight']])) {
                            // $sheet->getColumnDimension(Coordinate::stringFromColumnIndex($column))->setWidth($colTemplate['body']['width']);
                            $sheet->getRowDimension($row)->setRowHeight($dataRow[$colTemplate['body']['bindHeight']]);
                        }

                        // Применение стилей к одной ячейке тела таблицы
                        if (!empty($colTemplate['body']['bindStyles'])) {
                            if (!empty($dataRow[$colTemplate['body']['bindStyles']])) {
                                $this->setStyles(
                                    $cell,
                                    $dataRow[$colTemplate['body']['bindStyles']],
                                    $sheet
                                );
                            }
                        }

                        // Сбор данных об объединенных ячейках
                        if (!empty($colTemplate['body']['bindMerge'])) {
                            if (!empty($dataRow[$colTemplate['body']['bindMerge']])) {
                                $mergeId = $dataRow[$colTemplate['body']['bindMerge']];
                                $this->merge[$mergeId]['columns'][$column] = $column;
                                $this->merge[$mergeId]['rows'][$row] = $row;
                            }
                        }
                        $row++;
                    }
                }
            }
            $column++;
        }
    }

    /**
     * Объединяет ячейки, если некоторые обасти пересекаются, то не объединяет
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet Лист - объект Worksheet
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function mergeCellsXLS($sheet)
    {
        $mergedCells = [];
        foreach ($this->merge as $key => $value) {
            $errorMerge = false;
            $minColumn = min($value['columns']);
            $maxColumn = max($value['columns']);
            $minRow = min($value['rows']);
            $maxRow = max($value['rows']);

            if ($minRow > 0 && $maxRow > 0 && $minColumn > 0 && $maxColumn > 0) {
                // В $mergedCells хранятся адреса ячеек, которые уже объеденены в другие объединения.
                // Происходит проверка ячеек текущего объединения на вхождение в другие объединения, если они есть, то
                // текущее объединение не строится, иначе будет битый xls. Значит неправильно заданы mergeId.
                for ($i = $minColumn; $i <= $maxColumn; $i++) {
                    for ($j = $minRow; $j <= $maxRow; $j++) {
                        $cell = Coordinate::stringFromColumnIndex($i) . $j;
                        if (isset($mergedCells[$cell])) {
                            $errorMerge = true;
                        } else {
                            $mergedCells[$cell] = true;
                        }
                    }
                }
                if (!$errorMerge) {
                    $minColumn = Coordinate::stringFromColumnIndex($minColumn);
                    $maxColumn = Coordinate::stringFromColumnIndex($maxColumn);
                    $sheet->mergeCells($minColumn . $minRow . ':' . $maxColumn . $maxRow);
                }
            }
        }
    }

    /**
     * Применяет стили к ячейкам
     *
     * @param string $cells Ячейка или диапазон ячеек для применения стилей
     * @param array|string $styleNames Имена|имя стилей
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet Лист - объект Worksheet
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function setStyles($cells, $styleNames, $sheet)
    {
        if (is_string($styleNames)) {
            $styleNames = [$styleNames];
        }
        $mergedStyles = [];
        foreach ($styleNames as $styleName) {
            if (!empty($this->styles[$styleName])) {
                $mergedStyles = array_replace_recursive($mergedStyles, $this->styles[$styleName]);
            }
        }
        $sheet->getStyle($cells)->applyFromArray($mergedStyles);
    }

    /**
     * Преобразует разделители тысяч для pdf/html и обратно
     *
     * @param string $type Как преобразовывать 'pdf'|'xls'
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function correctThousands($type)
    {
        if (!empty($this->thousands)) {
            foreach ($this->thousands as $sheetIndex => $cellsArr) {
                /** @var array $selectedCells Ячейки в которых нужно скорректировать разделитель тысяч */
                $selectedCells = [];
                foreach ($cellsArr as $cells) {
                    if (preg_match('/(?<firstLetter>[A-Z]+)(?<firstDigit>[0-9]+)(?::(?<secondLetter>[A-Z]+)(?<secondDigit>[0-9]+))?/i', $cells, $result)) {
                        if (empty($result['secondDigit'])) {
                            $selectedCells[] = $result['firstLetter'] . $result['firstDigit'];
                        } else {
                            for ($i = $result['firstDigit']; $i <= $result['secondDigit']; $i++) {
                                $selectedCells[] = $result['firstLetter'] . $i;
                            }
                        }
                    }
                }

                $this->document->setActiveSheetIndex($sheetIndex);
                $sheet = $this->document->getActiveSheet();
                foreach ($selectedCells as $cell) {
                    $val = $sheet->getCell($cell)->getValue();
                    /* if ((substr($val, 0, 1) === '=') && (strlen($val) > 1)) {
                         $val = trim($sheet->getCell($cell)->getOldCalculatedValue());
                     }*/
                    if ($type == 'pdf') {
                        $val = number_format($val, 0, ',', ' ');
                    }
                    if ($type == 'xls') {
                        $val = (int)str_replace(' ', '', $val);
                    }
                    $sheet->setCellValue($cell, $val);
                }

            }
        }
    }
}

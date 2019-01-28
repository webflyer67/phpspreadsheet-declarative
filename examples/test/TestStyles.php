<?php
/**
 * Класс TestStyles - тестовый массив со стилями
 *
 * User: Alexander Kudryavtsev
 * Date: 22.10.2018
 */


namespace webflyer67\PhpspreadsheetDeclarative\Test;

use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Font;

class TestStyles
{


    /**  Возвращает массив с тестовыми стилями
     * @return array
     */
    public static function getStyles()
    {
        return [
            'border' => [
                'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN],],
            ],
            'mainCaption' => [
                'font' => ['bold' => true, 'color' => ['rgb' => '000000'], 'size' => 20, 'name' => 'Arial'
                ],
            ],
            'mainHead' => [
                'font' => ['bold' => true, 'color' => ['rgb' => 'ffffff'], 'size' => 10, 'name' => 'Arial'],
                'fill' => ['fillType' => Fill::FILL_SOLID, 'color' => ['rgb' => '538dd5']],
                'alignment' => ['horizontal' => 'center', 'vertical' => 'center', 'wrap' => true, 'shrinkToFit' => true],
            ],
            'mainBody' => [
                'alignment' => ['horizontal' => 'center', 'vertical' => 'center', 'wrap' => true, 'shrinkToFit' => true],
            ],
            'mainLink' => [
                'font' => ['bold' => true, 'color' => ['rgb' => '0000FF'], 'name' => 'Arial', 'underline' => Font::UNDERLINE_NONE],
            ],
            'wrap' => [
                'alignment' => ['vertical' => 'center', 'wrap' => true, 'shrinkToFit' => true],
            ],
            'right' => [
                'alignment' => ['horizontal' => 'right'],
            ],
            'left' => [
                'alignment' => ['horizontal' => 'left'],
            ],
            'red' => [
                'font' => ['bold' => true, 'color' => ['rgb' => 'aa0000']],
            ],
            'green' => [
                'font' => ['bold' => true, 'color' => ['rgb' => '00aa00']],
            ],
        ];
    }


}

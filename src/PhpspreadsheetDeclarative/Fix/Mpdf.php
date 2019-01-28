<?php

namespace webflyer67\PhpspreadsheetDeclarative\Fix;


use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf as PhpOfficeMpdf;

/**
 * Class Mpdf Переопределение PhpOffice\PhpSpreadsheet\Writer\Pdf для исправления неточностей при генерации pdf
 * @package webflyer67\PhpspreadsheetDeclarative\Fix
 */
class Mpdf extends PhpOfficeMpdf
{
    /**
     * Save Spreadsheet to file.
     * @param string $pFilename Name of the file to save as
     * @throws \Mpdf\MpdfException
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function save($pFilename)
    {
        $fileHandle = parent::prepareForSave($pFilename);

        //  Default PDF paper size
        $paperSize = 'LETTER'; //    Letter    (8.5 in. by 11 in.)

        //  Check for paper size and page orientation
        if (null === $this->getSheetIndex()) {
            $orientation = ($this->spreadsheet->getSheet(0)->getPageSetup()->getOrientation()
                == PageSetup::ORIENTATION_LANDSCAPE) ? 'L' : 'P';
            $printPaperSize = $this->spreadsheet->getSheet(0)->getPageSetup()->getPaperSize();
        } else {
            $orientation = ($this->spreadsheet->getSheet($this->getSheetIndex())->getPageSetup()->getOrientation()
                == PageSetup::ORIENTATION_LANDSCAPE) ? 'L' : 'P';
            $printPaperSize = $this->spreadsheet->getSheet($this->getSheetIndex())->getPageSetup()->getPaperSize();
        }
        $this->setOrientation($orientation);

        //  Override Page Orientation
        if (null !== $this->getOrientation()) {
            $orientation = ($this->getOrientation() == PageSetup::ORIENTATION_DEFAULT)
                ? PageSetup::ORIENTATION_PORTRAIT
                : $this->getOrientation();
        }
        $orientation = strtoupper($orientation);

        //  Override Paper Size
        if (null !== $this->getPaperSize()) {
            $printPaperSize = $this->getPaperSize();
        }

        if (isset(self::$paperSizes[$printPaperSize])) {
            $paperSize = self::$paperSizes[$printPaperSize];
        }

        //  Create PDF
        $config = ['tempDir' => $this->tempDir];

        $pdf = $this->createExternalWriterInstance($config);
        $ortmp = $orientation;
        $pdf->_setPageSize(strtoupper($paperSize), $ortmp);
        $pdf->DefOrientation = $orientation;
        $pdf->AddPageByArray([
            'orientation' => $orientation,
            'margin-left' => $this->inchesToMm($this->spreadsheet->getActiveSheet()->getPageMargins()->getLeft()),
            'margin-right' => $this->inchesToMm($this->spreadsheet->getActiveSheet()->getPageMargins()->getRight()),
            'margin-top' => $this->inchesToMm($this->spreadsheet->getActiveSheet()->getPageMargins()->getTop()),
            'margin-bottom' => $this->inchesToMm($this->spreadsheet->getActiveSheet()->getPageMargins()->getBottom()),
        ]);

        //  Document info
        $pdf->SetTitle($this->spreadsheet->getProperties()->getTitle());
        $pdf->SetAuthor($this->spreadsheet->getProperties()->getCreator());
        $pdf->SetSubject($this->spreadsheet->getProperties()->getSubject());
        $pdf->SetKeywords($this->spreadsheet->getProperties()->getKeywords());
        $pdf->SetCreator($this->spreadsheet->getProperties()->getCreator());
        $pdf->WriteHTML($this->generateHTMLHeader(false));


        $html = $this->generateSheetData();
        foreach (\array_chunk(\explode(PHP_EOL, $html), 1000) as $lines) {
            /**
             * Исправление проблем с совместимостью с Mpdf, перед генерацией документа вносим несклько изменений:
             * 1. В инлийн стилях убираем для границ "!important", иначе они не отрисовываются
             * 2. Для всех ссылок наследуем стили: цвет и подчеркивание, они применяются к родителю и по дефолту не наследуются
             */
            $lines = preg_replace('/(border[^<>;"\']+) !important/uis', '$1', $lines);
            $lines = preg_replace('/<a /uis', '<a style="color: inherit;text-decoration: inherit;"', $lines);

            $pdf->WriteHTML(\implode(PHP_EOL, $lines));
        }
        $pdf->WriteHTML($this->generateHTMLFooter());

        //  Write to file
        fwrite($fileHandle, $pdf->Output('', 'S'));

        parent::restoreStateAfterSave($fileHandle);
    }

    /**
     * Convert inches to mm.
     * @param float $inches
     * @return float
     */
    private function inchesToMm($inches)
    {
        return $inches * 25.4;
    }
}

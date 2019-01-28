<?php
/**
 * Класс TestTemplates - тестовый массив с шаблонами
 *
 * User: Alexander Kudryavtsev
 * Date: 22.10.2018
 */


namespace webflyer\PhpspreadsheetDeclarative\Test;

class TestTemplates
{

    /**  Возвращает тестовый шаблон 1
     * @return array
     */
    public static function getTemplate1()
    {
        return [
            'sheetCaption' => 'Электропоезда Экспресс',
            'tables' => [
                self::getTable1(),
                self::getTable2(),
            ]
        ];
    }

    /**  Возвращает тестовый шаблон 2
     * @return array
     */
    public static function getTemplate2()
    {
        return [
            'sheetCaption' => 'Электропоезда',
            'tables' => [
                self::getTable2(),
                self::getTable1(),
            ]
        ];
    }

    /**  Возвращает тестовый шаблон 2
     * @return array
     */
    public static function getTemplate3()
    {
        return [
            'sheetCaption' => 'Цены',
            'tables' => [
                self::getTable1(),
            ]
        ];
    }

    public static function getTable1()
    {
        return [
            'styles' => [
                'all' => 'border',
                'head' => 'mainHead',
                'body' => 'mainBody'
            ],

            'marginTop' => 0,
            'marginLeft' => 0,
            'bindTable' => 'prices',
            'columns' => [
                [
                    'head' => [
                        ['caption' => 'Номер', 'mergeId' => 'Номер', 'width' => 10, 'height' => 50],
                        ['mergeId' => 'Номер', 'height' => 50],
                    ],
                    'body' => ['bindColumn' => 'number'],
                ],
                [
                    'head' => [
                        ['caption' => 'Тип', 'mergeId' => 'Тип'],
                        ['mergeId' => 'Тип'],
                    ],
                    'body' => ['bindColumn' => 'type', 'width' => 15, 'bindColumnUrl' => 'url', 'styles' => ['mainLink']],
                ],
                [
                    'head' => [
                        ['caption' => 'Формат', 'styles' => ['left'], 'mergeId' => 'Формат'],
                        ['mergeId' => 'Формат'],
                    ],
                    'body' => ['bindColumn' => 'format', 'width' => 15, 'styles' => ['left'], 'bindMerge' => 'merge'],
                ],
                [
                    'head' => [
                        ['caption' => 1000000, 'mergeId' => 'Количество', 'filters' => 'thousands'],
                        ['mergeId' => 'Количество'],
                    ],
                    'body' => ['bindColumn' => 'count', 'width' => 15],
                ],
                [
                    'head' => [
                        ['caption' => 'Цены', 'mergeId' => 'Цены'],
                        ['caption' => 'Цена нам'],
                    ],
                    'body' => ['bindColumn' => 'price_our', 'width' => 20, 'defaultValue' => 'по запросу'],
                ],
                [
                    'head' => [
                        ['mergeId' => 'Цены'],
                        ['caption' => 'Цена клиенту', 'url' => 'http://localhost'],
                    ],
                    'body' => ['bindColumn' => 'price', 'width' => 20, 'filters' => 'thousands', 'defaultValue' => 'по запросу'],
                ],
                [
                    'head' => [
                        ['mergeId' => 'Цены', 'url' => 'http://localhost'],
                        ['caption' => 'Цена итого',],
                    ],
                    'body' => ['bindColumn' => 'total_price', 'width' => 20, 'filters' => 'thousands', 'defaultValue' => 'по запросу', 'bindStyles' => 'styles'],
                ],
            ]
        ];
    }

    public static function getTable2()
    {
        return [
            'marginTop' => 1,
            'marginLeft' => 0,
            'bindTable' => 'images',
            'columns' => [
                [
                    'body' => ['bindColumnImage' => 'img1', 'bindHeight' => 'height'],
                ],
                [], [],
                [
                    'body' => ['bindColumnImage' => 'img2',],
                ],
                [],
                [
                    'body' => ['bindColumnImage' => 'img3',],
                ]

            ]
        ];
    }
}

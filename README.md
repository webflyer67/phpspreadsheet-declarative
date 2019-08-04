# phpspreadsheet-declarative

PhpSpreadsheet Declarative - декларативное создание таблиц. Позволяет быстро и просто создавать таблицы через привязку стилей и массива данных к шаблону. Сохранение результаты в форматах xlsx, xls, pdf, html.


## Установка

Используйте [composer](https://getcomposer.org) чтобы установить PhpSpreadsheet Declarative в проект:

```sh
composer require webflyer67/phpspreadsheet-declarative
```

## Руководство

Все примеры находятся в /examples, результаты их работы в /runtime

### example01.php Быстрый старт - создание и сохранение документа
 
Writer::getWriter() - создание экземпляра объекта (новый xls документ)
addData('users', $users)- привязка массива с данными
addSheet($template) - добавление листа, в метод передается шаблон. 
Шаблон представляет ассоциативный массив установленного формата.
sheetCaption - название листа
tables - массив из шаблонов листов
tables.bindTable - имя привязанного массива с данными
tables.columns - массив шаблонов столбцов
tables.columns.head - свойства заголовка
tables.columns.body - свойства тела таблицы
tables.columns.head.caption - текст в заголовке
tables.columns.body.bindColumn -  имя привязанного свойства из привязанного массива с данными

writeDocument($fileNameFull . '.xlsx') - сохранение документа на диск 
```php
require_once $_SERVER['DOCUMENT_ROOT'] . '/vendor/autoload.php';
use webflyer67\PhpspreadsheetDeclarative\Writer;

/** @var array Массив с табличными данными */
$users = [
    ['id' => 1, 'name' => 'Alex', 'age' => '15', 'group' => 'admin'],
    ['id' => 2, 'name' => 'John', 'age' => '45', 'group' => 'admin'],
    ['id' => 3, 'name' => 'Bill', 'age' => '16', 'group' => 'user'],
    ['id' => 4, 'name' => 'Jimm', 'age' => '31', 'group' => 'user'],
];

/** @var array Массив с шаблоном для генерации таблицы */
$template = [
    'sheetCaption' => 'Пользователи', // Название листа
    'tables' => [
        [
            'bindTable' => 'users', // название связанного массива с данными
            'columns' => [ // заголовки столбцов и привязанные к ним данные
                [
                    'head' => [// заголовок
                        [
                            'caption' => 'id пользователя' // текст в заголовке
                        ],
                    ],
                    'body' => [ // тело
                        'bindColumn' => 'id' // привязанное значение из 'bindTable' => 'users'
                    ],
                ],
                [
                    'head' => [
                        ['caption' => 'Имя пользователя'],
                    ],
                    'body' => ['bindColumn' => 'name'],
                ],
            ]
        ],        
    ]
];

$fileName = 'example 01 ' . date("m.d.y H_i_s");
$fileNameFull = $_SERVER['DOCUMENT_ROOT'] . '/runtime/' . $fileName;
Writer::getWriter() // создание экземпляра объекта (новый xls документ)
    ->addData('users', $users) // привязка массива с данными
    ->addSheet($template, $pageSetup)   // добавление листа
    ->writeDocument($fileNameFull . '.xlsx'); // сохранение на диск Word 2007
```


### example02.php Создание и сохранение документа в разных форматах, xlsx, xls, html, pdf(3 варианта)

writeDocument() - сохраняет документ на диск. Первым аргументом передается полный путь к файлу. По расширению файла автоматически выбирается нужный Writer. Для pdf существуют 3 Writer'а. Для кириллицы лучше всего работает MPdfWriter, он же установлен по умолчанию. Выбрать pdf-writer можно вторым аргументом функции ('m','tc','dom'). Можно сохранять несколько раз в разные форматы и между сохранениями добавлять листы.

### example03.php Создание и отправка документа в браузер

sendDocument() - отправляет документ в браузер. Первым аргументом передается имя файла. По расширению файла автоматически выбирается нужный Writer и MIME-типы. После отправки происходит остановка работы.

### example04.php Добавление метаданных файла

setMeta() - добавляет метаданные файла. В метод передается ассоциативный массив метаданных, в котором ключами являются названия методов Phpspreadsheet для управления метаданными.
```php
$meta = [
    'Creator' => 'Vasilii Pupkin',
    'LastModifiedBy' => 'Vasilii Pupkin',
    'Title' => 'Test PhpspreadsheetDeclarative',
    'Subject' => 'Test PhpspreadsheetDeclarative',
    'Description' => 'Test PhpspreadsheetDeclarative',
    'Keywords' => 'PhpspreadsheetDeclarative, php, Phpspreadsheet, spreadsheet',
    'Category' => 'test spreadsheet',
    'Company' => 'webflyer67',
];

$fileName = 'example 04 ' . date("m.d.y H_i_s");
$fileNameFull = $_SERVER['DOCUMENT_ROOT'] . '/runtime/' . $fileName;
Writer::getWriter() // создание экземпляра объекта (новый xls документ)
    ->setMeta($meta)// Добавление метаданных файла
    ->addData('users', $users) // привязка массива с данными
    ->addSheet($template, $pageSetup)   // добавление листа
    ->writeDocument($fileNameFull . '.xlsx') // сохранение на диск Word 2007
    ->writeDocument($fileNameFull . '.pdf'); // сохранение на диск PDF
```
### example05.php Добавление формата листа

addSheet($template, $pageSetup)  - вторым аргументом передаются настройки листа, такие же как в Phpspreadsheet

```php
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
$pageSetup = [
    'Orientation' => PageSetup::ORIENTATION_LANDSCAPE,
    'PaperSize' => PageSetup::PAPERSIZE_A4,
];
->addSheet($template, $pageSetup)   // добавление листа
```

### example06.php Высота строк, ширина колонок
Ширину столбца можно задать как в заголовке(рекомендуется), так и в теле
tables.columns.head.width - ширина столбца(em)
tables.columns.body.width - ширина столбца(em)
Высота строки заголовка задаётся в 
tables.columns.head.height - высота строки(pt)
Для установки высоты строки в теле таблицы следует определить в массиве с табличными данными специальное свойство, и указать его имя в 
tables.columns.body.bindHeight - имя привязанного свойства из привязанного массива с данными, который содержит высоту строки(pt)

### example07.php Отступы
Можно задать отступы таблицы(в количествах ячеек)
tables.marginTop - отступ сверху(строк) от предыдущей  таблицы
tables.marginLeft - отступ слева(столбцов) от левого края документа

### example08.php Добавление стилей

addStyles($styles) - Добавление стилей. $styles - представляет собой массив ключ-значение, где ключ - имя стиля, значение - массив стилей, такой же как используется в phpspreadsheet. Стили могут применяться ко всей таблице, ко всему заголовку, ко всему телу, к ячейке заголовка, к столбцу тела или к ячейке тела. Стили задаются в виде массива имен стилей(если стиль один - можно задать строкой) из привязанного массива со стилями. 

tables.styles.all - стили для всей таблицы
tables.styles.head - стили для всех заголовков
tables.styles.body - стили для тела таблицы

tables.columns.head.styles - стили для текущего заголовка
tables.columns.body.styles - стили для текущего столбца
tables.columns.body.bindStyles - имя привязанного свойства из привязанного массива с данными, который содержит стили

```php
/** @var array Массив с табличными данными */
$users = [
    ['id' => 1, 'name' => 'Alex', 'age' => '15', 'group' => 'admin'],
    ['id' => 2, 'name' => 'John', 'age' => '45', 'group' => 'admin', 'cellStyles' => 'right'],
    ['id' => 3, 'name' => 'Bill', 'age' => '16', 'group' => 'user'],
    ['id' => 4, 'name' => 'Jimm', 'age' => '31', 'group' => 'user'],
];

/** @var array Массив с шаблоном для генерации таблицы */
$template = [
    'sheetCaption' => 'Пользователи', // Название листа
    'tables' => [
        [
            'bindTable' => 'users', // название связанного массива с данными
            'styles' => [  // Глобальные стили для всей таблицы
                'all' => 'border', // стили для всей таблицы
                'head' => ['primary-bg-color', 'primary-font', 'center'], // стили для всех заголовков
                'body' => 'center' // стили для тела таблицы
            ],
            'columns' => [ // заголовки столбцов и привязанные к ним данные
                [
                    'head' => [// заголовок
                        [
                            'caption' => 'id пользователя', // текст в заголовке
                            'styles' => ['left'], // Стили для текущего заголовка
                            'width' => 30, // ширина столбца(em)
                        ],
                    ],
                    'body' => [ // тело
                        'bindColumn' => 'id' // привязанное значение из 'bindTable' => 'users'
                    ],
                ],
                [
                    'head' => [
                        [
                            'caption' => 'Имя пользователя',
                            'width' => 30, // ширина столбца(em)
                         ],
                    ],
                    'body' => [
                        'bindColumn' => 'name',
                        'styles' => ['left'], // Стили для текущего столбца
                        'bindStyles' => 'cellStyles' // привязка стилей для отдельных ячеек тела таблицы
                    ],
                ],
            ]
        ],
        
    ]
];
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Font;
/** @var array Массив со стилями */
$styles = [
    'border' => [
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN],],
    ],
    'primary-bg-color' => [
        'fill' => ['fillType' => Fill::FILL_SOLID, 'color' => ['rgb' => '2980B9']],
    ],
    'primary-font' => [
        'font' => ['bold' => true, 'color' => ['rgb' => 'ffffff'], 'size' => 10, 'name' => 'Arial'],
    ],
    'center' => [
        'alignment' => ['horizontal' => 'center', 'vertical' => 'center', 'wrap' => true, 'shrinkToFit' => true],
    ],
    'left' => [
        'alignment' => ['horizontal' => 'left'],
    ],
    'right' => [
        'alignment' => ['horizontal' => 'right'],
    ],
];
Writer::getWriter() // создание экземпляра объекта (новый xls документ)
    ->addData('users', $users) // привязка массива с данными
    ->addStyles($styles)// Добавление стилей
    ->addSheet($template, $pageSetup)   // добавление листа
    ->writeDocument($fileNameFull . '.xlsx') // сохранение на диск Word 2007
    ->writeDocument($fileNameFull . '.pdf'); // сохранение на диск PDF

```
### example09.php Многострочный заголовок

Возможно сделать заголовок из нескольких строк, нужно в tables.columns.head добавить несколько значений. 

### example10.php Объединение ячеек

Чтобы объединить ячейки в заголовке нужно задать идентификатор в свойстве tables.columns.head.mergeId. Ячейки с одинаковым идентификатором объединятся.
Чтобы объединить ячейки в теле нужно задать
tables.columns.body.bindMerge - имя привязанного свойства из привязанного массива с данными, который содержит идентификатор объединения

### example11.php Добавление картинок, нет заголовка таблицы, пропуск столбцов

 tables.columns.body.bindImage - имя привязанного свойства из привязанного массива с данными, который содержит настройки для Drawing.
 Можно задавать все параметры, характерные для Drawing, а также гиперссылку Hyperlink. В свойстве Path задаётся путь к картинке, на диске или url.

 ```php
 /** @var array Массив с шаблоном для генерации таблицы */
$template =  [
    'sheetCaption' => 'Прайс',
    'tables' => [
        [             
            'bindTable' => 'images',
            'columns' => [
                [
                    'body' => [
                        'width' => 40,
                        'bindImage' => 'img1', 
                        'bindHeight' => 'height'
                    ],
                ],
                [], [],
                [
                    'body' => [
                        'width' => 40,
                        'bindImage' => 'img2',
                    ],
                ],
                [],
                [
                    'body' => [
                        'width' => 40,
                        'bindImage' => 'img3',
                    ],
                ]                        
            ]
        ]
    ]
];
```

### example12.php Гиперссылки

Для добавления гиперссылке в заголовке нужно добавить tables.columns.head.href 
В теле tables.columns.body.href - имя привязанного свойства из привязанного массива с данными, который содержит url для гиперссылки 

### example13.php Применение фильтров

Пока доступен только thousands - разделитель тысяч.
tables.columns.head.filters 
tables.columns.body.filters 

### example14.php Несколько листов

Для добавления нескольких листов нужно несколько раз вызвать addSheet()

### example15.php Несколько таблиц

Для добавления нескольких таблиц на одном листе нужно в tables добавить несколько элементов

### example16.php Имитация заполнения статическими данными без привязки к таблицам

Если требуется  добавить статические данные можно просто сделать заголовок таблицы без тела

### example17.php Доступ к объекту Spreadsheet и редактирование его напрямую

getDocument() - Возвращает объект Spreadsheet для возможности вносить в него правки напрямую

```php
$spreadsheet = Writer::getWriter() // создание экземпляра объекта (новый xls документ)
    ->addData('users', $users) // привязка массива с данными
    ->addSheet($template);   // добавление листа
  
$spreadsheet->getDocument()
  ->getSheet(0)
  ->setCellValue('F1', 'Вставка данных через объект Spreadsheet')
  ->setCellValue('B6', 'Вставка данных через объект Spreadsheet');

$spreadsheet->writeDocument($fileNameFull . '.xlsx'); // сохранение на диск Word 2007
```
### example18.php Генерация сложного документа(демонстрация почти всех возможностей библиотеки)

В данном примере показана работа почти всех описываемых возможностей вместе.

### Справочник
```php

```
#### Методы

getWriter() - Инициализирует и возвращает экземпляр данного класса

setMeta($meta) - Устанавливает метаданные документа

addData($name, $array) - Добавляет массив из которого впоследствии будет сформировано тело таблицы

addDatas($array) - Обертка над addData, чтоб можно было одним массивом добавить несколько массивов данных

addStyle($name, $array) - Добавляет массива со стилями из которого впоследствии будут браться стили

addStyles($array) - Обертка над addStyle, чтоб можно было одним массивом добавить несколько массивов стилей

addSheet($template, $setup) - Добавляет лист к документу. В переданном шаблоне установлены связи со стилями и данными, по этому шаблону строится лист

getDocument() - Возвращает объект Spreadsheet для возможности вносить в него правки напрямую

writeDocument($pFilename, $pdfType) - Сохраняет файл на диск

sendDocument($filename, $pdfType) - Отсылает файл в браузер

#### Структура шаблона

tables - массив из шаблонов листов
tables.bindTable - имя привязанного массива с данными
tables.marginTop - отступ сверху(строк) от предыдущей  таблицы
tables.marginLeft - отступ слева(столбцов) от левого края документа
tables.styles.all - стили для всей таблицы
tables.styles.head - стили для всех заголовков
tables.styles.body - стили для тела таблицы
          
tables.columns - массив шаблонов столбцов
tables.columns.head - свойства заголовка
tables.columns.head.caption - текст в заголовке
tables.columns.head.width - ширина столбца(em)
tables.columns.head.height - высота строки(pt)
tables.columns.head.styles - стили для текущего заголовка
tables.columns.head.mergeId - идентификатор объединения ячеек
tables.columns.head.href - url для гиперссылки
tables.columns.head.filters - фильтр(пока доступен только thousands - разделитель тысяч)
 
tables.columns.body - свойства тела таблицы
tables.columns.body.bindColumn -  имя привязанного свойства из привязанного массива с данными
tables.columns.body.width - ширина столбца(em)
tables.columns.body.bindHeight - имя привязанного свойства из привязанного массива с данными, который содержит высоту строки(pt)
tables.columns.body.styles - стили для текущего столбца
tables.columns.body.bindStyles - имя привязанного свойства из привязанного массива с данными, который содержит стили
tables.columns.body.bindMerge - имя привязанного свойства из привязанного массива с данными, который содержит идентификатор объединения
tables.columns.body.bindImage - имя привязанного свойства из привязанного массива с данными, который содержит настройки для Drawing.
tables.columns.body.href - имя привязанного свойства из привязанного массива с данными, который содержит url для гиперссылки
tables.columns.body.filters - фильтр(пока доступен только thousands - разделитель тысяч)

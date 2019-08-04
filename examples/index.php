<?php
require_once 'lib.php';
// формирование ссылок на все примеры
$fileList = [
    [
        'name' => 'example01.php',
        'desc' => 'Быстрый старт - создание и сохранение документа',
    ],
    [
        'name' => 'example02.php',
        'desc' => 'Создание и сохранение документа в разных форматах, xlsx, xls, html, pdf(3 варианта)',
    ],
    [
        'name' => 'example03.php',
        'desc' => 'Создание и отправка документа в браузер',
    ],
    [
        'name' => 'example04.php',
        'desc' => 'Добавление метаданных файла',
    ],
    [
        'name' => 'example05.php',
        'desc' => 'Добавление формата листа',
    ],
    [
        'name' => 'example06.php',
        'desc' => 'Высота строк, ширина колонок',
    ],
    [
        'name' => 'example07.php',
        'desc' => 'Отступы',
    ],
    [
        'name' => 'example08.php',
        'desc' => 'Добавление стилей',
    ],
    [
        'name' => 'example09.php',
        'desc' => 'Многострочный заголовок',
    ],
    [
        'name' => 'example10.php',
        'desc' => 'Объединение ячеек',
    ],
    [
        'name' => 'example11.php',
        'desc' => 'Добавление картинок',
    ],
    [
        'name' => 'example12.php',
        'desc' => 'Гиперссылки',
    ],
    [
        'name' => 'example13.php',
        'desc' => 'Применение фильтров',
    ],
    [
        'name' => 'example14.php',
        'desc' => 'Несколько листов',
    ],
    [
        'name' => 'example15.php',
        'desc' => 'Несколько таблиц',
    ],
    [
        'name' => 'example16.php',
        'desc' => 'Имитация заполнения статическими данными без привязки к таблицам',
    ],
    [
        'name' => 'example17.php',
        'desc' => 'Доступ к объекту Spreadsheet и редактирование его напрямую',
    ],
    [
        'name' => 'example18.php',
        'desc' => 'Генерация сложного документа(демонстрация почти всех возможностей библиотеки)',
    ],
    [
        'name' => 'example19.php',
        'desc' => 'Генерация сложного документа(pdf большого размера)',
    ],
    [
        'name' => 'example20.php',
        'desc' => 'Генерация сложного документа(xlsx большого размера)',
    ],     
];
foreach ($fileList as $file) {
    echo '<a target="_blank" href="' . $file['name'] . '">' . $file['name'] . ' - ' . $file['desc'] . '</a><br>';
}

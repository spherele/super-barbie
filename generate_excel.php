<?php
set_include_path(__DIR__);
require 'vendor/autoload.php';

// JSON-данные
$jsonFilePath = 'src/data.json';
$jsonData = file_get_contents($jsonFilePath);

$phpArray = json_decode($jsonData, true);

if ($phpArray === null && json_last_error() !== JSON_ERROR_NONE) {
    die('Ошибка при чтении JSON-файла');
}

// Создание объекта Excphp и генерация документа
$obj = new Excphp();
$obj->genDoc($jsonData);

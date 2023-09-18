<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Excphp {

    public function genDoc($jsonData)
    {
        // Декодируем JSON-данные
        $data = json_decode($jsonData, true);

        if ($data === null && json_last_error() !== JSON_ERROR_NONE) {
            die('Ошибка при чтении JSON-файла');
        }

        // Создаем новый Spreadsheet объект
        $spreadsheet = new Spreadsheet();

        // Подключение к активной таблице
        $sheet = $spreadsheet->getActiveSheet();

        // Объединяем ячейки от A1:C1
        $sheet->mergeCells('A1:C1');

        // Устанавливаем значение ячейке A1
        $sheet->setCellValue('A1', 'Данные стран');

        // Применяем стиль для выравнивания текста по центру
        $style = [
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'color' => ['rgb' => 'CCCCCC'], // Серый цвет
            ],
        ];

        $sheet->getStyle('A1')->applyFromArray($style);

        // Установка значений в шапку таблицы
        $sheet->setCellValue('A2', '№ п/п');
        $sheet->setCellValue('B2', 'ID страны');
        $sheet->setCellValue('C2', 'Название');

        // Создаем стиль для четных строк
        $evenRowStyle = [
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'color' => ['rgb' => 'FFFF00'], // Желтый цвет
            ],
        ];

        // Создаем стиль для нечетных строк
        $oddRowStyle = [
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'color' => ['rgb' => 'CCCCCC'], // Серый цвет
            ],
        ];

        // Получаем номер последней строки с записью и прибавляем 1, чтобы писать на следующей строке
        $highestRow = $sheet->getHighestRow() + 1;

        // Цикл по данным из JSON
        foreach ($data['result']['items'] as $key => $item) {
            $sheet->setCellValue("A$highestRow", $key + 1);
            $sheet->setCellValue("B$highestRow", $item['ID']);
            $sheet->setCellValue("C$highestRow", $item['NAME']);

            // Применяем стиль к текущей строке
            $style = $highestRow % 2 == 0 ? $evenRowStyle : $oddRowStyle;
            $sheet->getStyle("A$highestRow:C$highestRow")->applyFromArray($style);

            // Увеличиваем значение последней строки
            $highestRow++;
        }

        // Увеличим ширину ячейки B
        $sheet->getColumnDimension('B')->setWidth(70);

        // Увеличим ширину ячейки C
        $sheet->getColumnDimension('C')->setWidth(30);

        // Получение даты, которая будет использоваться в имени файла
        $dt = date('h:i:s');

        // Создание объекта Xlsx
        $writer = new Xlsx($spreadsheet);

        // Отправка файла в браузер
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header("Content-Disposition: attachment; filename=file-$dt.xlsx");
        $writer->save('php://output');
    }
}


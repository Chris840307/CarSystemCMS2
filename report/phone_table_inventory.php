<?php
require '../vendor/autoload.php';
require 'styles.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("手持使用行動電話統計表(清冊)");

// 日期變數
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(12);

// 表頭
$headers = [
    'A1' => '序',
    'B1' => '舉發單號',
    'C1' => '車號',
    'D1' => '車種',
    'E1' => '填單日期',
    'F1' => '違規日期',
    'G1' => '違規時間',
    'H1' => '違規地點',
    'I1' => '法條',
    'J1' => '違規事實',
    'K1' => '法條',
    'L1' => '違規事實',
    'M1' => '舉發單位',
    'N1' => '員警',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}

// 內容資料
$data = [
    ["E0YD33476", "NUH-7668", "機車", "1141002", "1141002", "1731", "竹市科學園路1號", "1610102", "變更異動不依規定申報登記", "", "", "新竹市警察局交通隊", "黃振浩"],
    ["E0YD33476", "NUH-7668", "機車", "1141002", "1141002", "1731", "竹市科學園路1號", "1610102", "變更異動不依規定申報登記", "", "", "新竹市警察局交通隊", "黃振浩"],
    ["E0YD33476", "NUH-7668", "機車", "1141002", "1141002", "1731", "竹市科學園路1號", "1610102", "變更異動不依規定申報登記", "", "", "新竹市警察局交通隊", "黃振浩"],
];

$seq = 1; //序
$row = 2;
foreach ($data as $d) {
    $sheet->setCellValue("A{$row}", $seq++);  //序
    $sheet->setCellValue("B{$row}", $d[0]);
    $sheet->setCellValue("C{$row}", $d[1]);
    $sheet->setCellValue("D{$row}", $d[2]);
    $sheet->setCellValue("E{$row}", $d[3]);
    $sheet->setCellValue("F{$row}", $d[4]);
    $sheet->setCellValue("G{$row}", $d[5]);
    $sheet->setCellValue("H{$row}", $d[6]);
    $sheet->setCellValue("I{$row}", $d[7]);
    $sheet->setCellValue("J{$row}", $d[8]);
    $sheet->setCellValue("K{$row}", $d[9]);
    $sheet->setCellValue("L{$row}", $d[10]);
    $sheet->setCellValue("M{$row}", $d[11]);
    $sheet->setCellValue("N{$row}", $d[12]);

    $row++;
}

// 全表格框線
$sheet->getStyle("A1:N" . ($row - 1))->applyFromArray($thinBorder);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(4);
$sheet->getColumnDimension('B')->setWidth(11);
$sheet->getColumnDimension('C')->setWidth(10);
$sheet->getColumnDimension('D')->setWidth(6);
$sheet->getColumnDimension('E')->setWidth(9);
$sheet->getColumnDimension('F')->setWidth(9);
$sheet->getColumnDimension('G')->setWidth(9);
$sheet->getColumnDimension('H')->setWidth(40);
$sheet->getColumnDimension('I')->setWidth(9);
$sheet->getColumnDimension('J')->setWidth(70);
$sheet->getColumnDimension('K')->setWidth(9);
$sheet->getColumnDimension('L')->setWidth(70);
$sheet->getColumnDimension('M')->setWidth(24);
$sheet->getColumnDimension('N')->setWidth(9);

// 檔名
$filename = "{$republicStartDate}-{$republicEndDate}-手持使用行動電話統計表(清冊).xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

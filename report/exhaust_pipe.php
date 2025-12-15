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
$sheet->setTitle("取締排氣管");

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
    'D1' => '填單日期',
    'E1' => '違規日期',
    'F1' => '違規時間',
    'G1' => '違規地點',
    'H1' => '法條',
    'I1' => '違規事實',
    'J1' => '法條',
    'K1' => '違規事實',
    'L1' => '舉發單位',
    'M1' => '員警',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}

// 內容資料
$data = [
    ["E0YD33476", "NUH-7668", "1141002", "1141002", "1731", "竹市科學園路1號", "1610102", "變更異動不依規定申報登記", "", "", "新竹市警察局交通隊", "黃振浩"],
    ["E0YD33476", "NUH-7668", "1141002", "1141002", "1731", "竹市科學園路1號", "1610102", "變更異動不依規定申報登記", "", "", "新竹市警察局交通隊", "黃振浩"],
    ["E0YD33476", "NUH-7668", "1141002", "1141002", "1731", "竹市科學園路1號", "1610102", "變更異動不依規定申報登記", "", "", "新竹市警察局交通隊", "黃振浩"],
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

    $row++;
}

// 全表格框線
$sheet->getStyle("A1:M" . ($row - 1))->applyFromArray($thinBorder);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(4);
$sheet->getColumnDimension('B')->setWidth(11);
$sheet->getColumnDimension('C')->setWidth(10);
$sheet->getColumnDimension('D')->setWidth(9);
$sheet->getColumnDimension('E')->setWidth(9);
$sheet->getColumnDimension('F')->setWidth(9);
$sheet->getColumnDimension('G')->setWidth(30);
$sheet->getColumnDimension('H')->setWidth(9);
$sheet->getColumnDimension('I')->setWidth(30);
$sheet->getColumnDimension('J')->setWidth(9);
$sheet->getColumnDimension('K')->setWidth(30);
$sheet->getColumnDimension('L')->setWidth(24);
$sheet->getColumnDimension('M')->setWidth(9);

// 檔名
$filetype = (isset($_GET['type']) && $_GET['type'] == 'week') ? '週報-' : '';
$filename = "{$filetype}{$republicStartDate}-{$republicEndDate}-取締排氣管.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

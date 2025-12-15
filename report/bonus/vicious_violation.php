<?php
require '../../vendor/autoload.php';
require '../styles.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("惡性違規核對表");

// 變數
$name = $_GET['name'];   //員警名稱
$id = $_GET['id'];   //員警編號
$total = "";  //開立件數
$point = "";  //點數

// 日期變數
$makeDate  = "11407";
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(12);

// 標題
$sheet->mergeCells('A1:F1');
$sheet->setCellValue('A1', "{$name} {$makeDate}月合計開立{$total}件，{$point}點");
$sheet->getStyle('A1')->applyFromArray($center);
$sheet->getStyle('A1')->getFont()->setSize(16);

// 表頭
$headers = [
    'A2' => '項次',
    'B2' => '違規名稱',
    'C2' => '攔停逕舉',
    'D2' => '件數',
    'E2' => '點數',
    'F2' => '小計',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}

// 內容資料
$data = [
    ["酒駕", "無", "0", "150", "0"],
];

$seq = 1; //序
$row = 3;
foreach ($data as $d) {
    $sheet->setCellValue("A{$row}", $seq++);  //序
    $sheet->setCellValue("B{$row}", $d[0]);
    $sheet->setCellValue("C{$row}", $d[1]);
    $sheet->setCellValue("D{$row}", $d[2]);
    $sheet->setCellValue("E{$row}", $d[3]);
    $sheet->setCellValue("F{$row}", $d[4]);
    $row++;
}

// 全表格框線
$sheet->getStyle("A1:F" . ($row - 1))->applyFromArray($thinBorder);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(6);
$sheet->getColumnDimension('B')->setWidth(50);
$sheet->getColumnDimension('C')->setWidth(10);
$sheet->getColumnDimension('D')->setWidth(6);
$sheet->getColumnDimension('E')->setWidth(6);
$sheet->getColumnDimension('F')->setWidth(6);

// 檔名
$filename = "{$name}({$id}).xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

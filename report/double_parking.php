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
$sheet->setTitle("並排停車舉發統計");

// 日期變數
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(16);

// 標題
$sheet->mergeCells('A1:D1');
$sheet->setCellValue('A1', "{$republicStartDate}-{$republicEndDate}違規停車舉發統計");
$sheet->getStyle('A1')->applyFromArray($center);
$sheet->getStyle('A1')->getFont()->setSize(16);

// 表頭
$headers = [
    'A2' => '舉發單位',
    'B2' => '倂排違停',
    'C2' => '一般違停',
    'D2' => '合計',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}
$sheet->getStyle('B2:D2')->applyFromArray($right);
$sheet->getStyle('A1:D1')->applyFromArray($thinBorder);
$sheet->getStyle('B1:D2')->getFont()->setSize(14);

// 內容資料
$data = [
    ['五組', '0', '0', '0'],
    ['西門所', '0', '0', '0'],
    ['北門所', '0', '0', '0'],
    ['樹林頭', '0', '0', '0'],
    ['南寮所', '0', '0', '0'],
    ['湳雅所', '0', '0', '0'],
    ['警備隊', '0', '0', '0'],
    ['小計', '0', '0', '0'],
    ['五組', '0', '0', '0'],
    ['東門所', '0', '0', '0'],
    ['東勢所', '0', '0', '0'],
    ['埔頂所', '0', '0', '0'],
    ['關東橋', '0', '0', '0'],
    ['文華所', '0', '0', '0'],
    ['警備隊', '0', '0', '0'],
    ['小計', '0', '0', '0'],
    ['五組', '0', '0', '0'],
    ['南門所', '0', '0', '0'],
    ['青草湖', '0', '0', '0'],
    ['香山所', '0', '0', '0'],
    ['朝山所', '0', '0', '0'],
    ['中華所', '0', '0', '0'],
    ['警備隊', '0', '0', '0'],
    ['小計', '0', '0', '0'],
    ['保安隊', '0', '0', '0'],
    ['交通隊', '0', '0', '0'],
    ['合計', '0', '0', '0'],
];

// 從第3列開始填資料
$row = 3;
foreach ($data as $d) {
    $col = 'A';
    foreach ($d as $val) {
        $sheet->setCellValue($col . $row, $val);
        $col++;
    }
    $row++;
}

$sheet->getStyle('A2:D' . ($row - 1))->applyFromArray($thinBorder);
$sheet->getStyle("B" . ($row - 1) . ":D" . ($row - 1))->applyFromArray($right);
$sheet->getStyle('B2:D' . ($row - 1))->getFont()->setSize(16);
for ($i = 3; $i <= 29; $i++) {
    $sheet->getStyle("A{$i}")->getFont()->setSize(14);
}

// 橘底
$sheet->getStyle('A2:D2')->applyFromArray($orangeFill2);
// 黃底
$sheet->getStyle('A10:D10')->applyFromArray($yellowFill2);
$sheet->getStyle('A26:D26')->applyFromArray($yellowFill2);
$sheet->getStyle('A29:D29')->applyFromArray($yellowFill2);
// 綠底
for ($i = 3; $i <= 9; $i++) {
    $sheet->getStyle("A{$i}")->applyFromArray($greenFill2);
}
for ($i = 11; $i <= 17; $i++) {
    $sheet->getStyle("A{$i}")->applyFromArray($greenFill2);
}
for ($i = 19; $i <= 25; $i++) {
    $sheet->getStyle("A{$i}")->applyFromArray($greenFill2);
}
$sheet->getStyle('A27')->applyFromArray($greenFill2);
$sheet->getStyle('A28')->applyFromArray($greenFill2);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(12);
foreach (range('B', 'D') as $col) {
    $sheet->getColumnDimension($col)->setWidth(10);
}

// 欄高
$sheet->getRowDimension(1)->setRowHeight(50);

// 檔名
$filetype = (isset($_GET['type']) && $_GET['type'] == 'week') ? '週報-' : '';
$filename = "{$filetype}{$republicStartDate}-{$republicEndDate}-並排停車舉發統計.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

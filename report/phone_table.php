<?php
require '../vendor/autoload.php';
require 'styles.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Borders;
use PhpOffice\PhpSpreadsheet\Style\Fill;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("手持使用行動電話統計表");

// 日期變數
$startDate = "114年10月01日";
$endDate   = "114年10月31日";
$makeDate  = "中華民國 114年11月04日";
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(14);

// 標題
$sheet->mergeCells('A1:K1');
$sheet->setCellValue('A1', '新竹市警察辦理取締以手持方式使用行動電話、電腦或其他相類功能裝置違規件數統計表');
$sheet->getStyle('A1')->applyFromArray($center);
$sheet->getStyle('A1')->getFont()->setSize(16);
$sheet->getStyle('A1')->getFont()->setBold(true);
$sheet->getStyle('A1')->getAlignment()->setWrapText(true);

// 統計期間
$sheet->mergeCells('A2:K2');
$sheet->setCellValue('A2', "統計期間:自{$startDate}起至{$endDate}止");
$sheet->getStyle('A2')->getFont()->setBold(true)->setSize(10);
$sheet->getStyle('A2')->applyFromArray($right);

// 表頭
$headers = [
    'A3' => '項目',
    'C3' => '以手持方式使用行動電話件數',
    'E3' => '以手持方式使用電腦件數',
    'G3' => '以手持方式使用其他相類功能裝置',
    'I3' => '合計',
    'K3' => '總計',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->mergeCells('A3:B4');
$sheet->mergeCells('C3:D4');
$sheet->mergeCells('E3:F4');
$sheet->mergeCells('G3:H4');
$sheet->mergeCells('I3:J4');
$sheet->mergeCells('K3:K4');
$sheet->getStyle('A3')->applyFromArray($centerTop);
$sheet->getStyle('C3:K4')->applyFromArray($center);
$sheet->getStyle('A3:K4')->getFont()->setBold(true)->setSize(12);
$sheet->getStyle('A1:K4')->applyFromArray($thinBorder);
$sheet->getStyle('A3:B4')->applyFromArray($diagonal)->getBorders()->setDiagonalDirection(Borders::DIAGONAL_DOWN);

$headers = [
    'A5' => '單位',
    'C5' => '汽車',
    'D5' => '機車',
    'E5' => '汽車',
    'F5' => '機車',
    'G5' => '汽車',
    'H5' => '機車',
    'I5' => '汽車',
    'J5' => '機車',
    'K5' => '汽車、機車',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->mergeCells('A5:B5');
$sheet->getStyle('A5')->applyFromArray($left);
$sheet->getStyle('C5:K5')->applyFromArray($center);
$sheet->getStyle('A5')->getFont()->setBold(true);
$sheet->getStyle('C5:K5')->getFont()->setBold(true)->setSize(12);

// 內容資料
$sheet->setCellValue('A6', "第\n一\n分\n局");
$sheet->mergeCells('A6:A13');
$sheet->getStyle('A6')->getAlignment()->setWrapText(true);
$sheet->setCellValue('A14', "第\n二\n分\n局");
$sheet->mergeCells('A14:A24');
$sheet->getStyle('A14')->getAlignment()->setWrapText(true);
$sheet->setCellValue('A25', "第\n三\n分\n局");
$sheet->mergeCells('A25:A29');
$sheet->getStyle('A25')->getAlignment()->setWrapText(true);

$data = [
    ['北門派出所', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['西門派出所', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['湳雅派出所', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['樹林頭派出所', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['南寮派出所', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['警備隊', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['民眾檢舉', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['合計', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['東門派出所', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['東勢派出所', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['埔頂派出所', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['文華派出所', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['關東橋派出所', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['警備隊', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['民眾檢舉', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['合計', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['香山派出所', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['南門派出所', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['青草湖派出所', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['朝山派出所', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['中華派出所', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['警備隊', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['民眾檢舉', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['合計', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
];

$row = 6;
foreach ($data as $d) {
    $col = 'B';
    foreach ($d as $val) {
        $sheet->setCellValue($col . $row, $val);
        $col++;
    }
    $row++;
}

$sheet->setCellValue('A30', '交通隊');
$sheet->setCellValue('A31', '保安隊');
$sheet->setCellValue('A32', '總計');
$data = [
    ['0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['0', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['0', '0', '0', '0', '0', '0', '0', '0', '0'],
];
foreach ($data as $d) {
    $col = 'C';
    foreach ($d as $val) {
        $sheet->setCellValue($col . $row, $val);
        $col++;
    }
    $row++;
}

$sheet->mergeCells('A30:B30');
$sheet->mergeCells('A31:B31');
$sheet->mergeCells('A32:B32');
$sheet->getStyle('A6:B32')->getFont()->setBold(true);
$sheet->getStyle('A4:A' . $row)->applyFromArray($left);
$sheet->getStyle('A5:K' . $row)->applyFromArray($center);
$sheet->getStyle('A3:K' . ($row - 1))->applyFromArray($thinBorder);
$sheet->getStyle('C13:K13')->getFont()->setBold(true);
$sheet->getStyle('C21:K21')->getFont()->setBold(true);
$sheet->getStyle('C29:K29')->getFont()->setBold(true);
$sheet->getStyle('C32:K32')->getFont()->setBold(true);

// 簽名列
$sheet->setCellValue('A' . ($row + 1), '承辦人：                        單位主管：');
$sheet->getStyle('A' . ($row + 1))->getFont()->setBold(true);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(6);
$sheet->getColumnDimension('B')->setWidth(14);
foreach (range('C', 'J') as $col) {
    $sheet->getColumnDimension($col)->setWidth(4);
}
$sheet->getColumnDimension('K')->setWidth(8);

// 欄高
$sheet->getRowDimension(1)->setRowHeight(40);
$sheet->getRowDimension(3)->setRowHeight(25);
$sheet->getRowDimension(4)->setRowHeight(25);

// 檔名
$filename = "{$republicStartDate}-{$republicEndDate}-手持使用行動電話統計表.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

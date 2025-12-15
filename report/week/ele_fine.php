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
$sheet->setTitle("電子罰單-疑似缺漏號一覽表");

// 分局變數
$organize = "";
switch ($_GET['organize']) {
    case 'one':
        $organize = "一分局";
        break;
    case 'two':
        $organize = "二分局";
        break;
    case 'three':
        $organize = "三分局";
        break;
    case 'traffic':
        $organize = "交通隊";
        break;
    case 'all':
        $organize = "全局";
        break;
}

// 日期變數
$startDate = "113/07/01";
$endDate   = "114/11/19";

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(12);

// 標題
$sheet->mergeCells('A1:I1');
$sheet->setCellValue('A1', "{$startDate}~{$endDate}缺漏號一覽表");
$sheet->getStyle('A1')->applyFromArray($center);
$sheet->getStyle('A1')->getFont()->setSize(16);
$sheet->getStyle("A1:I1")->applyFromArray($thinBorderMedium);

// 表頭
$headers = [
    'A2' => '缺漏單號',
    'B2' => '上單舉發單位',
    'C2' => '開單日',
    'D2' => '上單員警',
    'E2' => '下單舉發單位',
    'F2' => '開單日',
    'G2' => '下單員警',
    'H2' => '註記',
    'I2' => '清查結果',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}

// 內容資料
$data = [
    ["E20127356", "第一分局西門派出所", "1140704", "翁岑睿", "第一分局西門派出所", "1141106", "翁岑睿", "", ""],
];

$row = 3;
foreach ($data as $d) {
    $sheet->setCellValue("A{$row}", $d[0]);
    $sheet->setCellValue("B{$row}", $d[1]);
    $sheet->setCellValue("C{$row}", $d[2]);
    $sheet->setCellValue("D{$row}", $d[3]);
    $sheet->setCellValue("E{$row}", $d[4]);
    $sheet->setCellValue("F{$row}", $d[5]);
    $sheet->setCellValue("G{$row}", $d[6]);
    $sheet->setCellValue("H{$row}", $d[7]);
    $sheet->setCellValue("I{$row}", $d[8]);

    $row++;
}

// 全表格框線
$sheet->getStyle("A1:I" . ($row - 1))->applyFromArray($thinBorderMedium);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(12);
$sheet->getColumnDimension('B')->setWidth(20);
$sheet->getColumnDimension('C')->setWidth(10);
$sheet->getColumnDimension('D')->setWidth(12);
$sheet->getColumnDimension('E')->setWidth(20);
$sheet->getColumnDimension('F')->setWidth(10);
$sheet->getColumnDimension('G')->setWidth(10);
$sheet->getColumnDimension('H')->setWidth(10);
$sheet->getColumnDimension('I')->setWidth(22);

// 檔名
$filename = "{$organize}-電子罰單-疑似缺漏號一覽表.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

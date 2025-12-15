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
$sheet->setTitle("排氣管改裝取締報表");

// 分局變數
$organize = "";
switch ($_GET['organize']) {
    case 'one':
        $organize = "第一分局";
        break;
    case 'two':
        $organize = "第二分局";
        break;
    case 'three':
        $organize = "第三分局";
        break;
    case 'traffic':
        $organize = "交通隊";
        break;
    case 'all':
        $organize = "全局";
        break;
}

// 日期變數
$makeDate  = "114年7月";

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(14);

// 標題
$sheet->mergeCells('A1:D1');
$sheet->setCellValue('A1', "新竹市警察局{$organize}{$makeDate}\n執行排氣管改裝交通稽查取締統計表");
$sheet->getStyle('A1')->applyFromArray($center);
$sheet->getStyle('A1')->getFont()->setSize(16);
$sheet->getStyle('A1')->getAlignment()->setWrapText(true);

// 表頭
$headers = [
    'A2' => '單位',
    'B2' => '姓名',
    'C2' => '取締件數',
    'D2' => '備考',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}

// 內容資料
$data = [
    ["第一分局樹林頭派出所", "李元宏", "1", ""],
];

$row = 3;
foreach ($data as $d) {
    $sheet->setCellValue("A{$row}", $d[0]);
    $sheet->setCellValue("B{$row}", $d[1]);
    $sheet->setCellValue("C{$row}", $d[2]);
    $sheet->setCellValue("D{$row}", $d[3]);

    $row++;
}

$data = [
    ["小計", "", ""],
];

$row;
foreach ($data as $d) {
    $sheet->setCellValue("A{$row}", $d[0]);
    $sheet->setCellValue("B{$row}", $d[1]);
    $sheet->setCellValue("C{$row}", $d[2]);

    $row++;
}
$sheet->mergeCells("A" . ($row - 1) . ":B" . ($row - 1));
$sheet->getStyle("A2:D" . ($row - 1))->applyFromArray($center);

// 全表格框線
$sheet->getStyle("A1:D" . ($row - 1))->applyFromArray($thinBorder);

// 欄高
$sheet->getRowDimension(1)->setRowHeight(50);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(24);
$sheet->getColumnDimension('D')->setWidth(16);

// 檔名
$filename = "{$organize}{$makeDate}排氣管改裝取締件數.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

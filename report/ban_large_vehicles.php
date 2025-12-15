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
$sheet->setTitle("統計表");

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
$startDate = "114年10月01日";
$endDate   = "114年10月31日";
$makeDate  = "中華民國 114年11月04日";
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(12);

// 標題
$sheet->mergeCells('A1:F1');
$sheet->setCellValue('A1', '新竹市警察局取締大型車(裝載砂石,土方)違規績效統計表');
$sheet->getStyle('A1')->applyFromArray($center);
$sheet->getStyle('A1')->getFont()->setSize(14);
$sheet->getStyle('A1')->applyFromArray($yellowFill2);

// 表頭
$headers = [
    'A2' => '項次',
    'B2' => '統計項目',
    'C2' => '營大貨車',
    'D2' => '自大貨車',
    'E2' => '曳引車',
    'F2' => '合計'
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}
$sheet->getStyle('A2')->applyFromArray($center);
$sheet->getStyle('B2')->applyFromArray($left);
$sheet->getStyle('C2:F2')->applyFromArray($center);

// 橘底
$sheet->getStyle('A2:F2')->applyFromArray($orangeFill2);

// 內容資料
$data = [
    [1, "超載", "", "", "", 0],
    [2, "超速", "", "", "", 0],
    [3, "闖紅燈", 4, 2, "", 6],
    [4, "酒醉駕車", "", "", "", 0],
    [5, "無照駕駛", "", "", "", 0],
    [6, "未使用專用車廂", "", "", "", 0],
    [7, "車斗不合規定", "", "", "", 0],
    [8, "爭道行駛", "", "", "", 1],
    [9, "未依規定裝設行車記錄器", "", "", "", 1],
    [10, "違反管制規定", "", "", "", 0],
];

$row = 3;
foreach ($data as $d) {
    $sheet->setCellValue("A{$row}", $d[0]);
    $sheet->setCellValue("B{$row}", $d[1]);
    $sheet->setCellValue("C{$row}", $d[2]);
    $sheet->setCellValue("D{$row}", $d[3]);
    $sheet->setCellValue("E{$row}", $d[4]);
    $sheet->setCellValue("F{$row}", $d[5]);

    // 置中
    $sheet->getStyle("A{$row}:F{$row}")->applyFromArray($center);
    // B 靠左
    $sheet->getStyle("B{$row}")->applyFromArray($left);

    $row++;
}

// 合計列
$sheet->setCellValue("B{$row}", "合計");
$sheet->setCellValue("C{$row}", 4);
$sheet->setCellValue("D{$row}", 3);
$sheet->setCellValue("E{$row}", 1);
$sheet->setCellValue("F{$row}", 8);

$sheet->getStyle("A{$row}:F{$row}")->applyFromArray($center);

// A1~F13 加細邊框
$sheet->getStyle('A1:F13')->applyFromArray($thinBorder);

// 統計範圍
$row++;
$sheet->mergeCells("A{$row}:F{$row}");
$sheet->setCellValue("A{$row}", "統計範圍：{$startDate} 到 {$endDate}");
$sheet->getStyle("A{$row}")->applyFromArray($newMingCenter);
$sheet->getStyle("A{$row}:F{$row}")->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_NONE);

// 編製日期
$row++;
$sheet->mergeCells("A{$row}:F{$row}");
$sheet->setCellValue("A{$row}", "{$makeDate} 編製");
$sheet->getStyle("A{$row}")->applyFromArray($newMingRight);
$sheet->getStyle("A{$row}:F{$row}")->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_NONE);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(8);
$sheet->getColumnDimension('B')->setWidth(36);
$sheet->getColumnDimension('C')->setWidth(12);
$sheet->getColumnDimension('D')->setWidth(12);
$sheet->getColumnDimension('E')->setWidth(12);
$sheet->getColumnDimension('F')->setWidth(12);

// 檔名
$filename = "{$republicStartDate}-{$republicEndDate}-{$organize}-取締大型車.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

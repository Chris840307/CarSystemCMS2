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
$sheet->setTitle("淨牌專案統計清冊");

// 日期變數
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(12);

// 標題
$sheet->mergeCells('A1:K1');
$sheet->setCellValue('A1', '淨牌專案統計清冊');
$sheet->getStyle('A1')->applyFromArray($center);
$sheet->getStyle('A1')->getFont()->setSize(16);

// 表頭
$headers = [
    'A2' => '序',
    'B2' => '舉發單位',
    'C2' => '員警代碼',
    'D2' => '舉發員警',
    'E2' => '單號',
    'F2' => '法條',
    'G2' => '違規事實',
    'H2' => '車種',
    'I2' => '填單日期',
    'J2' => '違規日期',
    'K2' => '扣件內容',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}

// 內容資料
$data = [
    ["新竹市警察局交通隊", "H10600", "高慶妤", "E0YD81315", "1210501", "牌照借供他車使用", "機車", "1140927", "1140915", "無, 無, 無"],
    ["新竹市警察局交通隊", "H20130", "劉兆祖", "E0YD81367", "1210408", "使用註銷之牌照", "機車", "1141028", "1141028", "無, 無, 無"],
    ["新竹市警察局交通隊", "I15700", "李佑翰", "E2017734", "1210408", "使用註銷之牌照", "機車", "1141030", "1141028", "牌一面, 無, 無"],
    ["新竹市警察局交通隊", "I40760", "林家曄", "E0YE01390", "1210408", "使用註銷之牌照", "機車", "1141014", "1141013", "無, 無, 無"],
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
    $sheet->setCellValue("G{$row}", $d[5]);
    $sheet->setCellValue("H{$row}", $d[6]);
    $sheet->setCellValue("I{$row}", $d[7]);
    $sheet->setCellValue("J{$row}", $d[8]);
    $sheet->setCellValue("K{$row}", $d[9]);

    $row++;
}

// 全表格框線
$sheet->getStyle("A1:K" . ($row - 1))->applyFromArray($thinBorder);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(4);
$sheet->getColumnDimension('B')->setWidth(24);
$sheet->getColumnDimension('C')->setWidth(10);
$sheet->getColumnDimension('D')->setWidth(10);
$sheet->getColumnDimension('E')->setWidth(10);
$sheet->getColumnDimension('F')->setWidth(9);
$sheet->getColumnDimension('G')->setWidth(60);
$sheet->getColumnDimension('H')->setWidth(6);
$sheet->getColumnDimension('I')->setWidth(10);
$sheet->getColumnDimension('J')->setWidth(10);
$sheet->getColumnDimension('K')->setWidth(16);

// 檔名
$filetype = (isset($_GET['type']) && $_GET['type'] == 'week') ? '週報-' : '';
$filename = "{$filetype}{$republicStartDate}-{$republicEndDate}-淨牌專案統計清冊.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

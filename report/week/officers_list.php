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
$sheet->setTitle("員警舉發統計清冊");

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
    'A2' => '單位',
    'B2' => '員警',
    'C2' => '代碼',
    'D2' => '民眾檢舉',
    'E2' => '科技執法',
    'F2' => '其他',
    'G2' => '總件數',
    'H2' => '超速',
    'I2' => '手開單',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}

// 內容資料
$data = [
    ["新竹市警察局交通隊", "王道遠", "M49190", "0", "0", "11", "11", "0", "10"],
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
$sheet->getColumnDimension('A')->setWidth(20);
$sheet->getColumnDimension('B')->setWidth(10);
$sheet->getColumnDimension('C')->setWidth(10);
$sheet->getColumnDimension('D')->setWidth(10);
$sheet->getColumnDimension('E')->setWidth(10);
$sheet->getColumnDimension('F')->setWidth(6);
$sheet->getColumnDimension('G')->setWidth(8);
$sheet->getColumnDimension('H')->setWidth(6);
$sheet->getColumnDimension('I')->setWidth(8);

// 檔名
$filename = "交通隊員警舉發清冊.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

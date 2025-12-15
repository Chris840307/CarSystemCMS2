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
$sheet->setTitle("取締43條");

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
    'D1' => '違規人',
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
    ["E20136286", "MDN-6105", "", "1141014", "1141013", "0137", "竹市民權路與民權路56巷口", "4310101", "在道路上蛇行", "", "", "第二分局東門派出所", "馮堯祖"],
    ["E20136286", "MDN-6105", "", "1141014", "1141013", "0137", "竹市民權路與民權路56巷口", "4310101", "在道路上蛇行", "", "", "第二分局東門派出所", "馮堯祖"],
    ["E20136286", "MDN-6105", "", "1141014", "1141013", "0137", "竹市民權路與民權路56巷口", "4310101", "在道路上蛇行", "", "", "第二分局東門派出所", "馮堯祖"],
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
$sheet->getColumnDimension('B')->setWidth(12);
$sheet->getColumnDimension('C')->setWidth(10);
$sheet->getColumnDimension('D')->setWidth(9);
$sheet->getColumnDimension('E')->setWidth(10);
$sheet->getColumnDimension('F')->setWidth(9);
$sheet->getColumnDimension('G')->setWidth(9);
$sheet->getColumnDimension('H')->setWidth(26);
$sheet->getColumnDimension('I')->setWidth(12);
$sheet->getColumnDimension('J')->setWidth(80);
$sheet->getColumnDimension('K')->setWidth(12);
$sheet->getColumnDimension('L')->setWidth(80);
$sheet->getColumnDimension('M')->setWidth(24);
$sheet->getColumnDimension('N')->setWidth(12);

// 檔名
$filename = "{$republicStartDate}-{$republicEndDate}-取締43條.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

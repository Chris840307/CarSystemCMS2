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
$sheet->setTitle("標示單清冊");

// 日期變數
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(12);

// 表頭
$headers = [
    'A1' => '序',
    'B1' => '警政署下載檔',
    'C1' => '單號',
    'D1' => '車號',
    'E1' => '車種',
    'F1' => '違規日期',
    'G1' => '違規時間',
    'H1' => '舉發單位',
    'I1' => '員警',
    'J1' => '註記',
    'K1' => '入案日期',
    'L1' => '入案結果',
    'M1' => '對照單號',
    'N1' => '備註',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}

// 內容資料
$data = [
    ["TQ_BF000_20251001100042", "E1O4B0219", "BQV-5522", "自用小客車", "1141001", "0112", "第一分局南寮派出所", "吳雅茹", "", "1141003", "在道路上蛇行", "正常", ""],
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
$sheet->getColumnDimension('A')->setWidth(6);
$sheet->getColumnDimension('B')->setWidth(26);
$sheet->getColumnDimension('C')->setWidth(16);
$sheet->getColumnDimension('D')->setWidth(9);
$sheet->getColumnDimension('E')->setWidth(20);
$sheet->getColumnDimension('F')->setWidth(9);
$sheet->getColumnDimension('G')->setWidth(9);
$sheet->getColumnDimension('H')->setWidth(26);
$sheet->getColumnDimension('I')->setWidth(10);
$sheet->getColumnDimension('J')->setWidth(30);
$sheet->getColumnDimension('K')->setWidth(10);
$sheet->getColumnDimension('L')->setWidth(22);
$sheet->getColumnDimension('M')->setWidth(18);
$sheet->getColumnDimension('N')->setWidth(6);

// 檔名
$filename = "{$republicStartDate}-{$republicEndDate}-標示單完整清冊.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

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
$sheet->setTitle("電腦撤掣單");

// 日期變數
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(12);

// 表頭
$headers = [
    'A1' => '單號',
    'B1' => '車號',
    'C1' => '違規日期',
    'D1' => '填單日期',
    'E1' => '移送日期',
    'F1' => '應到案日期',
    'G1' => '應到案日稽核',
    'H1' => '入案日期',
    'I1' => '入案日稽核',
    'J1' => '法條1',
    'K1' => '違規事實1',
    'L1' => '法條2',
    'M1' => '違規事實2',
    'N1' => '違規地點',
    'O1' => '移送監理站',
    'P1' => '員警',
    'Q1' => '單位',
    'R1' => '貼條號碼',
    'S1' => '註記',
    'T1' => 'IIDNO',
    'U1' => 'INAME',
    'V1' => 'DRIVER',
    'W1' => 'OWNERX',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}

// 內容資料
$data = [
    ["E0Y1G0001", "AFB-0516", "1141010", "1141016", "1141016", "1141130", "45", "1141016", "0", "5610402", "在設有禁止停車標線之處所停車", "", "", "新竹市北區中華路二段445號", "苗栗監理站", "范逸軒", "新竹市警察局交通隊", "597855", "正常", "", "", "", ""],
    ["E0Y1G0001", "AFB-0516", "1141010", "1141016", "1141016", "1141130", "45", "1141016", "0", "5610402", "在設有禁止停車標線之處所停車", "", "", "新竹市北區中華路二段445號", "苗栗監理站", "范逸軒", "新竹市警察局交通隊", "597855", "正常", "", "", "", ""],
    ["E0Y1G0001", "AFB-0516", "1141010", "1141016", "1141016", "1141130", "45", "1141016", "0", "5610402", "在設有禁止停車標線之處所停車", "", "", "新竹市北區中華路二段445號", "苗栗監理站", "范逸軒", "新竹市警察局交通隊", "597855", "正常", "", "", "", ""],
];

$row = 2;
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
    $sheet->setCellValue("J{$row}", $d[9]);
    $sheet->setCellValue("K{$row}", $d[10]);
    $sheet->setCellValue("L{$row}", $d[11]);
    $sheet->setCellValue("M{$row}", $d[12]);
    $sheet->setCellValue("N{$row}", $d[13]);
    $sheet->setCellValue("O{$row}", $d[14]);
    $sheet->setCellValue("P{$row}", $d[15]);
    $sheet->setCellValue("Q{$row}", $d[16]);
    $sheet->setCellValue("R{$row}", $d[17]);
    $sheet->setCellValue("S{$row}", $d[18]);
    $sheet->setCellValue("T{$row}", $d[19]);
    $sheet->setCellValue("U{$row}", $d[20]);
    $sheet->setCellValue("V{$row}", $d[21]);
    $sheet->setCellValue("W{$row}", $d[22]);

    $row++;
}

// 全表格框線
$sheet->getStyle("A1:W" . ($row - 1))->applyFromArray($thinBorder);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(12);
$sheet->getColumnDimension('B')->setWidth(11);
$sheet->getColumnDimension('C')->setWidth(10);
$sheet->getColumnDimension('D')->setWidth(10);
$sheet->getColumnDimension('E')->setWidth(10);
$sheet->getColumnDimension('F')->setWidth(12);
$sheet->getColumnDimension('G')->setWidth(13);
$sheet->getColumnDimension('H')->setWidth(10);
$sheet->getColumnDimension('I')->setWidth(12);
$sheet->getColumnDimension('J')->setWidth(10);
$sheet->getColumnDimension('K')->setWidth(80);
$sheet->getColumnDimension('L')->setWidth(10);
$sheet->getColumnDimension('M')->setWidth(80);
$sheet->getColumnDimension('N')->setWidth(28);
$sheet->getColumnDimension('O')->setWidth(12);
$sheet->getColumnDimension('P')->setWidth(8);
$sheet->getColumnDimension('Q')->setWidth(22);
$sheet->getColumnDimension('R')->setWidth(10);
$sheet->getColumnDimension('S')->setWidth(6);
$sheet->getColumnDimension('T')->setWidth(10);
$sheet->getColumnDimension('U')->setWidth(10);
$sheet->getColumnDimension('V')->setWidth(10);
$sheet->getColumnDimension('W')->setWidth(10);

// 檔名
$filename = "{$republicStartDate}-{$republicEndDate}-建檔-電腦撤掣單.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

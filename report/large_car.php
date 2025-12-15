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
$sheet->setTitle("取締大型車");

// 日期變數
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(12);

// 標題
$sheet->mergeCells('A1:O1');
$sheet->setCellValue('A1', '大型車違規舉發清冊');
$sheet->getStyle('A1')->applyFromArray($center);
$sheet->getStyle('A1')->getFont()->setSize(16);

// 表頭
$headers = [
    'A2' => '序',
    'B2' => '舉發單位',
    'C2' => '員警代碼',
    'D2' => '舉發員警',
    'E2' => '單號',
    'F2' => '車號',
    'G2' => '法條',
    'H2' => '違規事實',
    'I2' => '車種',
    'J2' => '填單日期',
    'K2' => '違規日期',
    'L2' => '違規時間',
    'M2' => '違規地點',
    'N2' => '車主姓名',
    'O2' => '車主姓名',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}

// 內容資料
$data = [
    ["新竹市警察局交通隊", "A98230", "詹淳卉", "E20176450", "KES-8551", "5310001", "駕車行經有燈光號誌管制之交岔路口闖紅燈", "自用大貨車", "1141001", "1140910", "2236", "竹市公道五路二段228號前", "金象強商行", "金象強商行"],
    ["新竹市警察局交通隊", "A98230", "詹淳卉", "E20176450", "KES-8551", "5310001", "駕車行經有燈光號誌管制之交岔路口闖紅燈", "自用大貨車", "1141001", "1140910", "2236", "竹市公道五路二段228號前", "金象強商行", "金象強商行"],
    ["新竹市警察局交通隊", "A98230", "詹淳卉", "E20176450", "KES-8551", "5310001", "駕車行經有燈光號誌管制之交岔路口闖紅燈", "自用大貨車", "1141001", "1140910", "2236", "竹市公道五路二段228號前", "金象強商行", "金象強商行"],
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
    $sheet->setCellValue("L{$row}", $d[10]);
    $sheet->setCellValue("M{$row}", $d[11]);
    $sheet->setCellValue("N{$row}", $d[12]);
    $sheet->setCellValue("O{$row}", $d[13]);

    $row++;
}

// 全表格框線
$sheet->getStyle("A1:O" . ($row - 1))->applyFromArray($thinBorder);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(4);
$sheet->getColumnDimension('B')->setWidth(20);
$sheet->getColumnDimension('C')->setWidth(10);
$sheet->getColumnDimension('D')->setWidth(9);
$sheet->getColumnDimension('E')->setWidth(10);
$sheet->getColumnDimension('F')->setWidth(9);
$sheet->getColumnDimension('G')->setWidth(9);
$sheet->getColumnDimension('H')->setWidth(90);
$sheet->getColumnDimension('I')->setWidth(12);
$sheet->getColumnDimension('J')->setWidth(10);
$sheet->getColumnDimension('K')->setWidth(10);
$sheet->getColumnDimension('L')->setWidth(10);
$sheet->getColumnDimension('M')->setWidth(26);
$sheet->getColumnDimension('N')->setWidth(12);
$sheet->getColumnDimension('O')->setWidth(12);

// 檔名
$filetype = (isset($_GET['type']) && $_GET['type'] == 'week') ? '週報-' : '';
$filename = "{$filetype}{$republicStartDate}-{$republicEndDate}-取締大型車.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

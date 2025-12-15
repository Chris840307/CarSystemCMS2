<?php
require '../../vendor/autoload.php';
require '../styles.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Borders;
use PhpOffice\PhpSpreadsheet\Style\Fill;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("績效新表");

// 日期變數
$startDate = "114年11月20日";
$endDate   = "114年11月26日";

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(16);

// 標題
$sheet->mergeCells('A1:S1');
$sheet->setCellValue('A1', "新竹市警察局交通隊勤務組取締交通違規每週績效統計表");
$sheet->getStyle('A1')->applyFromArray($center);
$sheet->getStyle('A1')->getFont()->setSize(22);
$sheet->getStyle('A1')->getFont()->setBold(true);

// 統計期間
$sheet->mergeCells('A2:S2');
$sheet->setCellValue('A2', "{$startDate}至{$endDate}");
$sheet->getStyle('A2')->applyFromArray($right);
$sheet->getStyle('A2')->getFont()->setSize(14);
$sheet->getStyle('A2')->getFont()->setBold(true);

// 表頭
$sheet->mergeCells('A3:B4');
$sheet->getStyle('A3:B4')->applyFromArray($diagonal)->getBorders()->setDiagonalDirection(Borders::DIAGONAL_DOWN);
$sheet->mergeCells('C3:D3');
$sheet->mergeCells('E3:K3');
foreach (range('L', 'S') as $col) {
    $sheet->mergeCells("{$col}3:{$col}4");
}
$headers = [
    'C3' => '違規停車',
    'E3' => '重大違規',
    'C4' => '併排停車',
    'D4' => '一般違規停車',
    'E4' => '酒駕駕車',
    'F4' => '闖紅燈',
    'G4' => '嚴重超速',
    'H4' => '闖單行道',
    'I4' => '轉彎未依規定',
    'J4' => '未禮讓行人',
    'K4' => '惡意逼車',
    'L3' => '取締排氣管',
    'M3' => '取締大型車超載',
    'N3' => '未戴安全帽',
    'O3' => '淨牌',
    'P3' => '行人違規',
    'Q3' => '其他',
    'R3' => '合計',
    'S3' => '權利車登錄',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}
$sheet->getStyle('A3:S3')->applyFromArray($center);
 $sheet->getStyle('A3:S4')->getFont()->setBold(true);

// 內容資料
$data = [
    ["第一小隊", "林祥仁", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["第一小隊", "林祥仁", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["總計", "", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
];

$row = 5;
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

    $sheet->getStyle("A{$row}:S{$row}")->applyFromArray($center);
    $sheet->getStyle("A{$row}:B{$row}")->getFont()->setSize(14);
    $sheet->getStyle("A{$row}:S{$row}")->getFont()->setBold(true);

    $row++;
}
$sheet->mergeCells("A" . ($row - 1) . ":B" . ($row - 1));

// 全表格框線
$sheet->getStyle("A1:S" . ($row - 1))->applyFromArray($thinBorderMedium);

// 黃底
$sheet->getStyle('C4:K4')->applyFromArray($yellowFill3);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(10);

// 欄高
$sheet->getRowDimension(3)->setRowHeight(40);
$sheet->getRowDimension(4)->setRowHeight(70);

// 檔名
$filename = "績效新表.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

<?php
require '../vendor/autoload.php';
require 'styles.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Borders;
use PhpOffice\PhpSpreadsheet\Style\Fill;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("統計表");

// 日期變數
$startDate = "114年10月01日";
$endDate   = "114年10月31日";
$makeDate  = "中華民國 114年11月04日";
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(12);

// 標題
$sheet->mergeCells('A1:G1');
$sheet->setCellValue('A1', '新竹市警察局擴大取締路口不停讓行人執法專案成效統計表');
$sheet->getStyle('A1')->applyFromArray($center);
$sheet->getStyle('A1')->getFont()->setSize(14);
$sheet->getStyle('A1')->getFont()->setBold(true);

// 統計期間
$sheet->mergeCells('A2:G2');
$sheet->setCellValue('A2', "統計期間:自{$startDate}起至{$endDate}止");
$sheet->getStyle('A2')->getFont()->setBold(true)->setSize(10);
$sheet->getStyle('A2')->applyFromArray($right);

// 表頭
$headers = [
    'A3' => '取締項目',
    'B3' => '路口不停讓行人',
    'C3' => '非號誌化路口未依標誌、標線、號誌停車再開',
    'D3' => '人行道違規停車與違規臨時停車',
    'E3' => '取締道路障礙',
    'F3' => '行人違規',
    'G3' => '合計'
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->getStyle('A3')->applyFromArray($centerTop);
$sheet->getStyle('A4:G3')->applyFromArray($center);
$sheet->getStyle('A3:F3')->getFont()->setBold(true);
$sheet->getStyle('A1:G3')->applyFromArray($thinBorder);
$sheet->getStyle('A3')->applyFromArray($diagonal)->getBorders()->setDiagonalDirection(Borders::DIAGONAL_DOWN);

// 單位名稱
$sheet->setCellValue('A4', '單位名稱');
$sheet->getStyle('A4')->getFont()->setBold(true);
$sheet->getStyle('A4')->applyFromArray($left);
// 法條內容
$sheet->setCellValue('B4', '第44條2項及第3項');
$sheet->getStyle('B4')->getAlignment()->setWrapText(true);
$sheet->setCellValue('C4', '第60條第2項第3款及第45條第1項第18款');
$sheet->getStyle('C4')->getAlignment()->setWrapText(true);
$sheet->setCellValue('D4', '第55條第1項第1款及第56條第1項第1款');
$sheet->getStyle('D4')->getAlignment()->setWrapText(true);
$sheet->setCellValue('E4', '第82條第1項各款');
$sheet->getStyle('E4')->getAlignment()->setWrapText(true);
$sheet->setCellValue('F4', '第78條');
$sheet->getStyle('F4')->getAlignment()->setWrapText(true);
// 合計
$sheet->mergeCells('G3:G4');
$sheet->setCellValue('G3', '合計');
$sheet->getStyle('G3')->applyFromArray($center);
$sheet->getStyle('B3:F3')->applyFromArray($center);
$sheet->getStyle('B4:F4')->applyFromArray($center);
$sheet->getStyle('A3:F3')->getFont()->setBold(true);
$sheet->getStyle('A3:G4')->applyFromArray($thinBorder);

// 內容資料
$data = [
    ['新竹市警察局總計', '526', '3', '340', '175', '0', '1044'],
    ['保安隊', '0', '0', '0', '0', '0', '0'],
    ['交通隊', '18', '3', '16', '0', '0', '37'],
    ['一分局各所合計', '71', '0', '17', '7', '0', '95'],
    ['西門派出所', '5', '0', '12', '1', '0', '18'],
    ['北門派出所', '0', '0', '2', '6', '0', '8'],
    ['樹林頭派出所', '2', '0', '3', '0', '0', '5'],
    ['南寮派出所', '1', '0', '0', '0', '0', '1'],
    ['湳雅派出所', '55', '0', '0', '0', '0', '55'],
    ['警備隊', '0', '0', '0', '0', '0', '0'],
    ['五組', '8', '0', '0', '0', '0', '8'],
    ['二分局各所合計', '41', '0', '279', '138', '0', '458'],
    ['東門派出所', '17', '0', '4', '0', '0', '21'],
    ['東勢派出所', '3', '0', '167', '1', '0', '171'],
    ['埔頂派出所', '3', '0', '42', '2', '0', '47'],
    ['關東橋派出所', '4', '0', '69', '2', '0', '75'],
    ['文華派出所', '13', '0', '1', '129', '0', '143'],
    ['警備隊', '0', '0', '0', '0', '0', '0'],
    ['五組', '1', '0', '0', '0', '0', '1'],
    ['三分局各所合計', '41', '0', '279', '138', '0', '458'],
    ['南門派出所', '17', '0', '4', '0', '0', '21'],
    ['青草湖派出所', '3', '0', '167', '1', '0', '171'],
    ['香山派出所', '3', '0', '42', '2', '0', '47'],
    ['朝山派出所', '4', '0', '69', '2', '0', '75'],
    ['中華派出所', '13', '0', '1', '129', '0', '143'],
    ['警備隊', '0', '0', '0', '0', '0', '0'],
    ['五組', '1', '0', '0', '0', '0', '1'],
    ['日平均數', '17', '0', '11', '6', '0', '34'],
];

$row = 5;
foreach ($data as $d) {
    $col = 'A';
    foreach ($d as $val) {
        $sheet->setCellValue($col . $row, $val);
        $col++;
    }
    $row++;
}

$sheet->getStyle('A4:A' . $row)->applyFromArray($left);
$sheet->getStyle('A5:G' . $row)->applyFromArray($center);
$sheet->getStyle('A3:G' . $row)->applyFromArray($thinBorder);

// 橘底
$sheet->getStyle('A5:G5')->applyFromArray($orangeFill);
$sheet->getStyle('A32:G32')->applyFromArray($orangeFill);
// 籃底
$sheet->getStyle('A6:G6')->applyFromArray($blueFill);
// 黃底
$sheet->getStyle('A7:G7')->applyFromArray($yellowFill);
// 綠底
$sheet->getStyle('A8:G8')->applyFromArray($greenFill);
$sheet->getStyle('A16:G16')->applyFromArray($greenFill);
$sheet->getStyle('A24:G24')->applyFromArray($greenFill);

// 備註列
$sheet->setCellValue("A{$row}", '備註:各單位應於每週一12時前將上週執行成效，依附表－取締件數統計表，函報本署彙辦。');
$sheet->mergeCells("A{$row}:G{$row}");
$sheet->getStyle("A{$row}")->applyFromArray($left);
$sheet->getStyle("A{$row}")->getFont()->setSize(14);
$sheet->getStyle("A{$row}")->getAlignment()->setWrapText(true);

// 簽名列
$sheet->setCellValue('A' . ($row + 1), '             承辦人:                                     單位主官:						');
$sheet->getStyle('A' . ($row + 1))->getFont()->setBold(true);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(26);
foreach (range('B', 'G') as $col) {
    $sheet->getColumnDimension($col)->setWidth(10);
}

// 欄高
$sheet->getRowDimension(3)->setRowHeight(90);
$sheet->getRowDimension(4)->setRowHeight(80);
$sheet->getRowDimension(33)->setRowHeight(50);

// 檔名
$filetype = (isset($_GET['type']) && $_GET['type'] == 'week') ? '週報-' : '';
$filename = "{$filetype}{$republicStartDate}-{$republicEndDate}-取締路口不停讓行人.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

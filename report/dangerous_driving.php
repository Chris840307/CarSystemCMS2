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
$sheet->setTitle("防制危險駕車績效成果表");

// 日期變數
$startDate = "114年10月01日";
$endDate   = "114年10月31日";
$makeDate  = "114年09月";
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(14);

// 標題
$sheet->mergeCells('A1:K1');
$sheet->setCellValue('A1', "新竹市警察局{$makeDate}防制危險駕車績效成果表");
$sheet->getStyle('A1')->applyFromArray($center);
$sheet->getStyle('A1')->getFont()->setSize(16);
$sheet->getStyle('A1')->getFont()->setBold(true);

// 統計期間
$sheet->mergeCells('A2:K2');
$sheet->setCellValue('A2', "{$startDate}00時至{$endDate}24時");
$sheet->getStyle('A2')->applyFromArray($right);
$sheet->getStyle('A2')->getFont()->setSize(16);
$sheet->getStyle('A2')->getFont()->setBold(true);

// 表頭
$sheet->mergeCells('B3:B4');
$sheet->mergeCells('C3:C4');
$sheet->mergeCells('D3:D4');
$sheet->mergeCells('E3:E4');
$sheet->mergeCells('F3:F4');
$sheet->mergeCells('G3:G4');
$sheet->mergeCells('H3:H4');
$sheet->mergeCells('I3:I4');
$sheet->mergeCells('J3:K4');
$headers = [
    'A3' => '內容',
    'B3' => '違規內容',
    'C3' => '法條',
    'D3' => '車種',
    'E3' => '第一分局',
    'F3' => '第二分局',
    'G3' => '第三分局',
    'H3' => '保安隊',
    'I3' => '交通隊',
    'J3' => '件數',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->setCellValue("A4", '項目');
$sheet->getStyle('A4')->applyFromArray($thinBorder);
$sheet->getStyle('A3:K4')->applyFromArray($center);
$sheet->getStyle('A1:K4')->applyFromArray($thinBorder);
$sheet->getStyle('A3')->applyFromArray($diagonal)->getBorders()->setDiagonalDirection(Borders::DIAGONAL_DOWN);


// 內容資料
$data = [
    ['1', '蛇行或危險方式駕車', '第43條第1項第1款'],
    ['2', '車速超過規定最高速度40公里', '第43條第1項第2款'],
    ['3', '以迫近、驟然變換車道或不當方式，迫使他車讓道', '第43條第1項第3款'],
    ['4', '行駛中任意驟然減速、煞車或暫停', '第43條第1項第4款'],
    ['5', '拆除消音器', '第43條第1項第5款前段'],
    ['6', '以其他方式造成噪音', '第43條第1項第5款後段'],
    ['7', '在高速公路或快速公路迴車、倒車、逆向行駛', '第43條第1項第6款'],
    ['8', '二輛以上共同違反第一項規定', '第43條第3項'],
    ['9', '吊扣牌照車輛再次違規', '第43條第4項'],
    ['10', '車身、引擎及底盤等重要設備變更或調換', '第18條'],
    ['11', '除頭燈外之燈光、雨刷等設備不全或損壞', '第16條第1項第2款前段'],
    ['12', '無照駕車', '第21條'],
];
$row = 5;
foreach ($data as $d) {
    $col = 'A';
    foreach ($d as $val) {
        $sheet->setCellValue($col . $row, $val);
        $sheet->getStyle($col . $row)->getAlignment()->setWrapText(true);

        $mer_row = $row + 1;
        $sheet->mergeCells("A{$row}:A{$mer_row}");
        $sheet->mergeCells("B{$row}:B{$mer_row}");
        $sheet->mergeCells("C{$row}:C{$mer_row}");

        $col++;
    }
    $row += 2;
}

$data = [
    ['機車', '0', '0', '0', '1', '0', '1'],
    ['汽車', '0', '0', '0', '1', '0', '1'],
    ['機車', '0', '0', '0', '1', '0', '1'],
    ['汽車', '0', '0', '0', '1', '0', '1'],
    ['機車', '0', '0', '0', '1', '0', '1'],
    ['汽車', '0', '0', '0', '1', '0', '1'],
    ['機車', '0', '0', '0', '1', '0', '1'],
    ['汽車', '0', '0', '0', '1', '0', '1'],
    ['機車', '0', '0', '0', '1', '0', '1'],
    ['汽車', '0', '0', '0', '1', '0', '1'],
    ['機車', '0', '0', '0', '1', '0', '1'],
    ['汽車', '0', '0', '0', '1', '0', '1'],
    ['機車', '0', '0', '0', '1', '0', '1'],
    ['汽車', '0', '0', '0', '1', '0', '1'],
    ['機車', '0', '0', '0', '1', '0', '1'],
    ['汽車', '0', '0', '0', '1', '0', '1'],
    ['機車', '0', '0', '0', '1', '0', '1'],
    ['汽車', '0', '0', '0', '1', '0', '1'],
    ['機車', '0', '0', '0', '1', '0', '1'],
    ['汽車', '0', '0', '0', '1', '0', '1'],
    ['機車', '0', '0', '0', '1', '0', '1'],
    ['汽車', '0', '0', '0', '1', '0', '1'],
    ['機車', '0', '0', '0', '1', '0', '1'],
    ['汽車', '0', '0', '0', '1', '0', '1'],
];
$row = 5;
foreach ($data as $d) {
    $col = 'D';
    foreach ($d as $val) {
        $sheet->setCellValue($col . $row, $val);
        $sheet->mergeCells("J" . $row . ":K" . $row);
        $col++;
    }
    $row++;
}

$data = [
    'B29' => '移送法辦（刑法185條）',
    'B31' => '違反社會秩序維護法案件（第72、74條）',
    'B33' => '具學生身分',
    'B34' => '未滿十八歲',
];
foreach ($data as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->mergeCells("A29:A30");
$sheet->mergeCells("A31:A32");
$sheet->mergeCells("B29:D30");
$sheet->mergeCells("B31:D32");
$sheet->mergeCells("B33:D33");
$sheet->mergeCells("B34:D34");

$data = [
    ['0', '0', '0', '0', '0', '0', '件'],
    ['0', '0', '0', '0', '0', '0', '人'],
    ['0', '0', '0', '0', '0', '0', '件'],
    ['0', '0', '0', '0', '0', '0', '人'],
    ['0', '0', '0', '0', '0', '0', '人'],
    ['0', '0', '0', '0', '0', '0', '人'],
];
$row = 29;
foreach ($data as $d) {
    $col = 'E';
    foreach ($d as $val) {
        $sheet->setCellValue($col . $row, $val);
        $col++;
    }
    $row++;
}


$sheet->setCellValue("B35", '動用警力情形');
$sheet->mergeCells("A35:A38");
$sheet->mergeCells("B35:B38");
$data = [
    ['制服', '0', '0', '0', '0', '0', '0', '人次'],
    ['便服', '0', '0', '0', '0', '0', '0', '人次'],
    ['蒐證', '0', '0', '0', '0', '0', '0', '人次'],
    ['合計', '0', '0', '0', '0', '0', '0', '人次'],
];
$row = 35;
foreach ($data as $d) {
    $col = 'D';
    foreach ($d as $val) {
        $sheet->setCellValue($col . $row, $val);
        $col++;
    }
    $row++;
}

$sheet->setCellValue("B39", '受理他轄（110報案系統）');
$sheet->setCellValue("B40", '受理他轄（其他）');
$sheet->setCellValue("B41", '受理本轄（110報案系統）');
$sheet->setCellValue("B42", '受理本轄（其他）');
$data = [
    ['0', '件'],
    ['0', '件'],
    ['0', '件'],
    ['0', '件'],
];
$row = 39;
foreach ($data as $d) {
    $col = 'J';
    foreach ($d as $val) {
        $sheet->setCellValue($col . $row, $val);
        $col++;
    }
    $row++;
}



$sheet->getStyle('A5:K43')->applyFromArray($center);
$sheet->getStyle('B5:B26')->applyFromArray($left);
$sheet->getStyle('C5:C26')->applyFromArray($left);
$sheet->getStyle('A5:K44')->applyFromArray($thinBorder);

// 備考列
$sheet->mergeCells("B{$row}:K{$row}");
$sheet->setCellValue("A{$row}", '備考');
$sheet->setCellValue("B{$row}", '本案移送法辦或違反社會秩序維護法案件之統計，係危險駕駛行為之相關案件。');

// 簽名列
$sheet->mergeCells("A" . ($row + 1) . ":K" . ($row + 1));
$sheet->setCellValue("A" . ($row + 1), "製表：                   審核：                      單位主管：										");
$sheet->getStyle("A" . ($row + 1))->applyFromArray($left);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(8);
$sheet->getColumnDimension('B')->setWidth(25);
$sheet->getColumnDimension('C')->setWidth(25);
$sheet->getColumnDimension('J')->setWidth(5);
$sheet->getColumnDimension('K')->setWidth(5);
foreach (range('D', 'I') as $col) {
    $sheet->getColumnDimension($col)->setWidth(6);
}

// 欄高
$sheet->getRowDimension(3)->setRowHeight(24);
$sheet->getRowDimension(43)->setRowHeight(40);
$sheet->getRowDimension(44)->setRowHeight(30);

// 檔名
$filename = "{$republicStartDate}-{$republicEndDate}-防制危險駕車績效成果表.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

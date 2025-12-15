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
$sheet->mergeCells('A1:J1');
$sheet->setCellValue('A1', "新竹市警察局全國同步擴大執行防制危險駕車專案績效月報表");
$sheet->getStyle('A1')->applyFromArray($center);
$sheet->getStyle('A1')->getFont()->setSize(16);
$sheet->getStyle('A1')->getFont()->setBold(true);

// 統計期間
$sheet->mergeCells('A2:J2');
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
$sheet->mergeCells('I3:J4');
$headers = [
    'A3' => '內容',
    'B3' => '違規內容',
    'C3' => '車種',
    'D3' => '第一分局',
    'E3' => '第二分局',
    'F3' => '第三分局',
    'G3' => '保安隊',
    'H3' => '交通隊',
    'I3' => '件數',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->setCellValue("A4", '項目');
$sheet->getStyle('A4')->applyFromArray($thinBorder);
$sheet->getStyle('A3:J4')->applyFromArray($center);
$sheet->getStyle('A1:J4')->applyFromArray($thinBorder);
$sheet->getStyle('A3')->applyFromArray($diagonal)->getBorders()->setDiagonalDirection(Borders::DIAGONAL_DOWN);

// 內容資料
$data = [
    ['1', '蛇行或危險方式駕車(第43條第1項第1款)'],
    ['2', '以迫近、驟然變換車道或不當方式，迫使他車讓道（第43條第1項第3款）'],
    ['3', '行駛中任意驟然減速、煞車或暫停（第43條第1項第4款）'],
    ['4', '車速超過規定最高速度40公里以上(第43條第1項第2款)'],
    ['5', '拆除消音器（第43條第1項第5款前段）'],
    ['6', '以其他方式造成噪音（第43條第1項第5款後段）'],
    ['7', '二輛以上共同違反第一項規定(第43條第3項)'],
    ['8', '吊扣牌照車輛再次違規(第43條第4項)'],
    ['9', '車身、引擎、底盤及電系等重要設備變更或調換(第18條)'],
    ['10', '燈光等設備不全或損壞(第16條第1項第2款前段)'],
    ['11', '無照駕駛'],
];
$row = 5;
foreach ($data as $d) {
    $col = 'A';
    foreach ($d as $val) {
        $sheet->setCellValue($col . $row, $val);
        $sheet->getStyle($col . $row)->getAlignment()->setWrapText(true);
        $sheet->getStyle("B{$row}")->getAlignment()->setWrapText(true);

        $mer_row = $row + 1;
        $sheet->mergeCells("A{$row}:A{$mer_row}");
        $sheet->mergeCells("B{$row}:B{$mer_row}");

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
];
$row = 5;
foreach ($data as $d) {
    $col = 'C';
    foreach ($d as $val) {
        $sheet->setCellValue($col . $row, $val);
        $sheet->mergeCells("I" . $row . ":J" . $row);
        $col++;
    }
    $row++;
}

$data = [
    'B27' => '移送法辦（刑法185條）',
    'B29' => '違反社會秩序維護法案件（第72、74條）',
    'B31' => '具學生身分',
    'B32' => '未滿十八歲',
];
foreach ($data as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->mergeCells("A27:A28");
$sheet->mergeCells("A29:A30");
$sheet->mergeCells("B27:C28");
$sheet->mergeCells("B29:C30");
$sheet->mergeCells("B31:C31");
$sheet->mergeCells("B32:C32");

$data = [
    ['0', '0', '0', '0', '0', '0', '件'],
    ['0', '0', '0', '0', '0', '0', '人'],
    ['0', '0', '0', '0', '0', '0', '件'],
    ['0', '0', '0', '0', '0', '0', '人'],
    ['0', '0', '0', '0', '0', '0', '人'],
    ['0', '0', '0', '0', '0', '0', '人'],
];
$row = 27;
foreach ($data as $d) {
    $col = 'D';
    foreach ($d as $val) {
        $sheet->setCellValue($col . $row, $val);
        $col++;
    }
    $row++;
}


$sheet->setCellValue("B33", '動用警力情形');
$sheet->mergeCells("A33:A36");
$sheet->mergeCells("B33:B36");
$data = [
    ['制服', '0', '0', '0', '0', '0', '0', '人次'],
    ['便服', '0', '0', '0', '0', '0', '0', '人次'],
    ['蒐證', '0', '0', '0', '0', '0', '0', '人次'],
    ['合計', '0', '0', '0', '0', '0', '0', '人次'],
];
$row = 33;
foreach ($data as $d) {
    $col = 'C';
    foreach ($d as $val) {
        $sheet->setCellValue($col . $row, $val);
        $col++;
    }
    $row++;
}

$sheet->setCellValue("B37", '受理他轄（110報案系統）');
$sheet->setCellValue("B38", '受理他轄（其他）');
$sheet->setCellValue("B39", '受理本轄（110報案系統）');
$sheet->setCellValue("B40", '受理本轄（其他）');
$data = [
    ['0', '件'],
    ['0', '件'],
    ['0', '件'],
    ['0', '件'],
];
$row = 37;
foreach ($data as $d) {
    $col = 'I';
    foreach ($d as $val) {
        $sheet->setCellValue($col . $row, $val);
        $col++;
    }
    $row++;
}
$sheet->getStyle('A5:J40')->applyFromArray($center);
$sheet->getStyle('B5:B26')->applyFromArray($left);
$sheet->getStyle('A5:J42')->applyFromArray($thinBorder);

// 備考列
$sheet->mergeCells("B{$row}:J{$row}");
$sheet->setCellValue("A{$row}", '備考');
$sheet->setCellValue("B{$row}", 'ㄧ、本案移送法辦或違反社會秩序維護法案件之統計，係危險駕駛行為之相關案件。
二、本表於勤務結束翌日上午9時前傳真本署交通組彙辦。
傳真電話：02-23219861；警用：722-2256');
$sheet->getStyle('A41')->applyFromArray($center);
$sheet->getStyle('B41:J41')->applyFromArray($top);
$sheet->getStyle('B41:J41')->getAlignment()->setWrapText(true);

// 簽名列
$sheet->mergeCells("A" . ($row + 1) . ":J" . ($row + 1));
$sheet->setCellValue("A" . ($row + 1), "製表：                   審核：                      單位主管：										");
$sheet->getStyle("A" . ($row + 1))->applyFromArray($left);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(8);
$sheet->getColumnDimension('B')->setWidth(30);
$sheet->getColumnDimension('J')->setWidth(5);
foreach (range('C', 'I') as $col) {
    $sheet->getColumnDimension($col)->setWidth(6);
}

// 欄高
$sheet->getRowDimension(3)->setRowHeight(24);
$sheet->getRowDimension(41)->setRowHeight(80);
$sheet->getRowDimension(42)->setRowHeight(30);

// 檔名
$filename = "週報-{$republicStartDate}-{$republicEndDate}-防制危險駕車績效成果表.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

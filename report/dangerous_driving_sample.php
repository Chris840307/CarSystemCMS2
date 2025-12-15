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
$sheet->setTitle("防制危險駕車績效成果表");

// 日期變數
$makeDate  = "114年10月";
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(10);

// 標題
$sheet->mergeCells('A1:AN1');
$sheet->setCellValue('A1', "內政部警政署");
$sheet->getStyle('A1')->applyFromArray($center);
$sheet->getStyle('A1')->getFont()->setSize(19);
$sheet->getStyle('A1')->getFont()->setBold(true);
$sheet->getStyle('A1:AN1')->applyFromArray($thinBorder);

$sheet->mergeCells('A2:AN2');
$sheet->setCellValue('A2', "{$makeDate} 防制危險駕車績效成果表");
$sheet->getStyle('A2')->applyFromArray($center);
$sheet->getStyle('A2')->getFont()->setSize(19);
$sheet->getStyle('A2')->getFont()->setBold(true);
$sheet->getStyle('A2:AN2')->applyFromArray($thinBorder);

// 第一層表頭
$sheet->mergeCells('A3:A4');
$sheet->mergeCells('B3:C3');
$sheet->mergeCells('D3:E3');
$sheet->mergeCells('F3:G3');
$sheet->mergeCells('H3:I3');
$sheet->mergeCells('J3:K3');
$sheet->mergeCells('L3:M3');
$sheet->mergeCells('N3:O3');
$sheet->mergeCells('P3:Q3');
$sheet->mergeCells('R3:S3');
$sheet->mergeCells('T3:U3');
$sheet->mergeCells('V3:W3');
$sheet->mergeCells('X3:Y3');

$sheet->mergeCells('Z3:AA4');
$sheet->mergeCells('AB3:AC4');
$sheet->mergeCells('AD3:AD4');
$sheet->mergeCells('AE3:AE4');
$sheet->mergeCells('AF3:AG4');
$sheet->mergeCells('AH3:AJ4');

$sheet->mergeCells('AK3:AN3');
$sheet->mergeCells('AK4:AL4');
$sheet->mergeCells('AM4:AN4');
$headers = [
    'A3' => '項目',
    'B3' => '蛇行或危險方式駕車',
    'D3' => '車速超過規定最高速度40公里',
    'F3' => '以迫近、驟然變換車道或不當方式，迫使他車讓道',
    'H3' => '行駛中任意驟然減速、煞車或暫停',
    'J3' => '拆除消音器',
    'L3' => '以其他方式造成噪音',
    'N3' => '在高速公路或快速公路迴車、倒車、逆向行駛',
    'P3' => '二輛以上共同違反第一項規定',
    'R3' => '吊扣牌照車輛再次違規',
    'T3' => '車身、引擎及底盤等重要設備變更或調換',
    'V3' => '除頭燈外之燈光、雨刷等設備不全或損壞',
    'X3' => '無照駕車',

    'Z3' => '移送法辦',
    'AB3' => '違反社會秩序維護法案件',
    'AD3' => '具學生生份',
    'AE3' => '未滿十八歲',
    'AF3' => '傷亡人數',
    'AH3' => '動用警力情形',

    'AK3' => '受理民眾報案件數',
    'AK4' => '他轄案件',
    'AM4' => '本轄案件',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}

// 第二層表頭
$sheet->mergeCells('B4:C4');
$sheet->mergeCells('D4:E4');
$sheet->mergeCells('F4:G4');
$sheet->mergeCells('H4:I4');
$sheet->mergeCells('J4:K4');
$sheet->mergeCells('L4:M4');
$sheet->mergeCells('N4:O4');
$sheet->mergeCells('P4:Q4');
$sheet->mergeCells('R4:S4');
$sheet->mergeCells('T4:U4');
$sheet->mergeCells('V4:W4');
$sheet->mergeCells('X4:Y4');
$headers = [
    'B4' => '第43條第1項第1款',
    'D4' => '第43條第1項第2款',
    'F4' => '第43條第1項第3款',
    'H4' => '第43條第1項第4款',
    'J4' => '第43條第1項第5款前段',
    'L4' => '第43條第1項第5款後段',
    'N4' => '第43條第1項第6款',
    'P4' => '第43條第3項',
    'R4' => '第43條第4項',
    'T4' => '第18條',
    'V4' => '第16條第1項第2款前段',
    'X4' => '第21條',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}

// 第三層表頭
$headers = [
    'A5' => '單位',
    'B5' => '汽車',
    'C5' => '機車',
    'D5' => '汽車',
    'E5' => '機車',
    'F5' => '汽車',
    'G5' => '機車',
    'H5' => '汽車',
    'I5' => '機車',
    'J5' => '汽車',
    'K5' => '機車',
    'L5' => '汽車',
    'M5' => '機車',
    'N5' => '汽車',
    'O5' => '機車',
    'P5' => '汽車',
    'Q5' => '機車',
    'R5' => '汽車',
    'S5' => '機車',
    'T5' => '汽車',
    'U5' => '機車',
    'V5' => '汽車',
    'W5' => '機車',
    'X5' => '汽車',
    'Y5' => '機車',
    'Z5' => '件數',
    'AA5' => '人數',
    'AB5' => '件數',
    'AC5' => '人數',
    'AD5' => '人數',
    'AE5' => '人數',
    'AF5' => '死亡',
    'AG5' => '受傷',
    'AH5' => '制服',
    'AI5' => '便服',
    'AJ5' => '蒐證',
    'AK5' => '110系統報案方式',
    'AL5' => '其他報案方式',
    'AM5' => '110系統報案方式',
    'AN5' => '其他報案方式',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}

// 第三列
$sheet->setCellValue('A6', '新竹市警察局');

// 第四列
$sheet->setCellValue('A7', '備考');

$sheet->getStyle('A3:AN5')->applyFromArray($center);
$sheet->getStyle('A3:AN5')->applyFromArray($thinBorder);
$sheet->getStyle('A6:A7')->applyFromArray($center);
$sheet->getStyle('A6')->applyFromArray($thinBorder);
$sheet->mergeCells('B7:AN7');
$sheet->getStyle('A7:AN7')->applyFromArray($thinBorder);

// 內容資料
$data = [
    ['0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'],
];
$row = 6;
foreach ($data as $d) {
    $col = 'B';
    foreach ($d as $val) {
        $sheet->setCellValue($col . $row, $val);
        $col++;
    }
    $row++;
}
$sheet->getStyle('B6:AN' . ($row - 1))->applyFromArray($thinBorder);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(18);

// 欄高
$sheet->getRowDimension(3)->setRowHeight(60);
$sheet->getRowDimension(4)->setRowHeight(60);
$sheet->getRowDimension(5)->setRowHeight(40);

// 檔名
$filename = "{$republicStartDate}-{$republicEndDate}-R12防制危險駕車績效成果表上傳範本.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

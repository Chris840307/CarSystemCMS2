<?php
require '../vendor/autoload.php';
require 'styles.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("員警舉發-整理交通秩序成果統計比較表");

// 日期變數
$startDate = "114年10月01日";
$endDate   = "114年10月31日";
$makeDate  = "中華民國 114年11月04日";
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：新細明體
$spreadsheet->getDefaultStyle()->getFont()->setName('新細明體')->setSize(12);

// 標題
$sheet->mergeCells('A1:AE1');
$sheet->setCellValue('A1', "新竹市警察局{$startDate} 至 {$endDate} 整理交通秩序成果統計比較表");
$sheet->getStyle('A1')->applyFromArray($center);
$sheet->getStyle('A1')->getFont()->setSize(12);

// 表頭
$headers = [
    'A2' => '項目',
    'C2' => '加強取締重大交通違規專案',
    'O2' => '清道',
    'P2' => '自行車',
    'Q2' => '淨牌',
    'R2' => '違規停車',
    'V2' => '驗收成效統計',
    'Z2' => '攔檢大型車',
    'AE2' => '總計',

    'A3' => '承辦人',

    'A4' => '條文',
    'C4' => '1',
    'E4' => '2',
    'G4' => '3',
    'J4' => '4',
    'K4' => '5',
    'M4' => '6',
    'N4' => '7',
    'O4' => '9',
    'P4' => '10',
    'Q4' => '11',
    'R4' => '12',
    'V4' => '13',
    'W4' => '14',
    'X4' => '15',
    'Y4' => '16',
    'Z4' => '17',
    'AA4' => '18',
    'AB4' => '19',
    'AC4' => '20',
    'AD4' => '21',

    'A5' => '',
    'C5' => '酒醉駕車',
    'E5' => '闖紅燈',
    'G5' => '一般',
    'H5' => '四十公里超速未達嚴重超速',
    'I5' => '嚴重超速',
    'J5' => '闖單行道',
    'K5' => '轉彎未依規定',
    'M5' => '不停讓行人',
    'N5' => '危險駕駛',
    'O5' => '攤架取締沒入',
    'P5' => '自行車違規',
    'Q5' => '淨牌',
    'R5' => '',
    'V5' => '無照駕駛',
    'W5' => '未戴安全帽',
    'X5' => '未繫安全帶',
    'Y5' => '不依規撥接行動電話',
    'Z5' => '超載',
    'AA5' => '載運危險物品不依規定',
    'AB5' => '違規超車',
    'AC5' => '汽車佔用停等區',
    'AD5' => '其他違規',

    'A6' => '',
    'C6' => '條項款 35',
    'E6' => '條項款 53.1.',
    'G6' => '條項款 33.1.1 40. 43.01.02',
    'H6' => '',
    'I6' => '超速60公里以上',
    'J6' => '條項款 45.1 45.3',
    'K6' => '條項款 48.1.2 48.1.4',
    'M6' => '條項款 44.2 44.3',
    'N6' => '條項款 43.1.1 43.1.3 43.1.4',
    'O6' => '條項款 78至84',
    'P6' => '條項款 69至76',
    'Q6' => '條項款 12 13.  1',
    'R6' => '條項款 55.01.01. 至 56.01.10.',
    'V6' => '條項款 21.01.01',
    'W6' => '條項款 31.6.',
    'X6' => '條項款 31.1',
    'Y6' => '條項款 31之1.1.',
    'Z6' => '條項款 29',
    'AA6' => '條項款 30',
    'AB6' => '條項款 47',
    'AC6' => '條項款 60.02.03.',
    'AD6' => '條項款',

    'A7' => '車種',
    'C7' => '汽、機',
    'E7' => '汽、機',
    'G7' => '',
    'H7' => '',
    'I7' => '',
    'J7' => '汽、機',
    'K7' => '汽車',
    'L7' => '機車',
    'M7' => '機車',
    'N7' => '汽、機',
    'O7' => '汽、機',
    'P7' => '自行車',
    'Q7' => '汽、機',
    'R7' => '汽、機',
    'V7' => '汽、機',
    'W7' => '未戴安全帽',
    'X7' => '汽車',
    'Y7' => '汽、機',
    'Z7' => '汽、機',
    'AA7' => '汽、機',
    'AB7' => '汽、機',
    'AC7' => '汽車',
    'AD7' => '汽、機',

    'A8' => '單位',
    'C8' => '舉發',
    'D8' => '公危',
    'E8' => '欄舉',
    'F8' => '逕舉',
    'G8' => '件',
    'H8' => '',
    'I8' => '件',
    'J8' => '件',
    'K8' => '件',
    'L8' => '件',
    'M8' => '件',
    'N8' => '件',
    'O8' => '件',
    'P8' => '件',
    'Q8' => '件',
    'R8' => '攔(併排、其他)',
    'T8' => '逕(併排、其他)',
    'V8' => '件',
    'W8' => '件',
    'X8' => '件',
    'Y8' => '件',
    'Z8' => '件',
    'AA8' => '件',
    'AB8' => '件',
    'AC8' => '件',
    'AD8' => '件',
    'AE8' => '件',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}
$sheet->mergeCells('A1:AE1');
$sheet->mergeCells('A2:B2');
$sheet->mergeCells('C2:N2');
$sheet->mergeCells('C3:N3');
$sheet->mergeCells('R2:U2');
$sheet->mergeCells('R3:U3');
$sheet->mergeCells('V2:Y2');
$sheet->mergeCells('V3:Y3');
$sheet->mergeCells('Z2:AA2');
$sheet->mergeCells('Z3:AA3');
$sheet->mergeCells('AE2:AE6');
$sheet->mergeCells('A3:B3');

$sheet->mergeCells('C4:D4');
$sheet->mergeCells('E4:F4');
$sheet->mergeCells('G4:I4');
$sheet->mergeCells('K4:L4');
$sheet->mergeCells('R4:U4');

$sheet->mergeCells('A4:B6');
$sheet->mergeCells('C5:D5');
$sheet->mergeCells('E5:F5');
$sheet->mergeCells('K5:L5');
$sheet->mergeCells('R5:U5');

$sheet->mergeCells('C6:D6');
$sheet->mergeCells('E6:F6');
$sheet->mergeCells('K6:L6');
$sheet->mergeCells('R6:U6');

$sheet->mergeCells('A7:B7');
$sheet->mergeCells('C7:D7');
$sheet->mergeCells('E7:F7');
$sheet->mergeCells('R7:U7');

$sheet->mergeCells('A8:B8');
$sheet->mergeCells('R8:S8');
$sheet->mergeCells('T8:U8');

$sheet->getStyle('A3:AE8')->getFont()->setSize(8);
$sheet->getStyle('AE2')->getFont()->setSize(12);
$sheet->getStyle('C6:AD6')->getFont()->setSize(6);
$sheet->getStyle('C7:AE7')->getFont()->setSize(7);
$sheet->getStyle('C8:AE8')->getFont()->setSize(8);
$sheet->getStyle("A1:AE8")->applyFromArray($thinBorder);
$sheet->getStyle('A1:AE4')->applyFromArray($center);
$sheet->getStyle('A5')->applyFromArray($center);
$sheet->getStyle('C5:AE5')->applyFromArray($centerTop);
$sheet->getStyle('C6:AE6')->applyFromArray($top);
$sheet->getStyle('A7:AE8')->applyFromArray($center);
$sheet->getStyle('A1:AE6')->getAlignment()->setWrapText(true);
$sheet->getStyle('C6:AD7')->getFont()->getColor()->setARGB('FF808080');


$headers = [
    'A9' => '第一分局',
    'A11' => '第二分局',
    'A13' => '第三分局',
    'A15' => '保安隊',
    'A17' => '交通隊',
    'A19' => '合計',
    'B21' => '說明',
    'C21' => '一、本表如有未盡事宜、得隨時修正、補充之。二、本表起用:(114.01.01日修正)。',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}
$sheet->mergeCells('A9:A10');
$sheet->mergeCells('A11:A12');
$sheet->mergeCells('A13:A14');
$sheet->mergeCells('A15:A16');
$sheet->mergeCells('A17:A18');
$sheet->mergeCells('A19:A20');
$sheet->getStyle('A8:AE8')->applyFromArray($thinBottomMedium);


// 內容資料-本期
$data = [
    ["本期", "4", "7", "38", "559", "4", "", "", "524", "691", "127", "71", "", "15", "10", "23", "17", "35", "27", "2451", "17", "68", "", "9", "", "", "5", "246", "1412", ""],
    ["本期", "4", "7", "38", "559", "4", "", "", "524", "691", "127", "71", "", "15", "10", "23", "17", "35", "27", "2451", "17", "68", "", "9", "", "", "5", "246", "1412", ""],
    ["本期", "4", "7", "38", "559", "4", "", "", "524", "691", "127", "71", "", "15", "10", "23", "17", "35", "27", "2451", "17", "68", "", "9", "", "", "5", "246", "1412", ""],
    ["本期", "4", "7", "38", "559", "4", "", "", "524", "691", "127", "71", "", "15", "10", "23", "17", "35", "27", "2451", "17", "68", "", "9", "", "", "5", "246", "1412", ""],
    ["本期", "4", "7", "38", "559", "4", "", "", "524", "691", "127", "71", "", "15", "10", "23", "17", "35", "27", "2451", "17", "68", "", "9", "", "", "5", "246", "1412", ""],
    ["本期", "4", "7", "38", "559", "4", "", "", "524", "691", "127", "71", "", "15", "10", "23", "17", "35", "27", "2451", "17", "68", "", "9", "", "", "5", "246", "1412", ""],
];

$row = 9;
foreach ($data as $d) {
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
    $sheet->setCellValue("P{$row}", $d[14]);
    $sheet->setCellValue("Q{$row}", $d[15]);
    $sheet->setCellValue("R{$row}", $d[16]);
    $sheet->setCellValue("S{$row}", $d[17]);
    $sheet->setCellValue("T{$row}", $d[18]);
    $sheet->setCellValue("U{$row}", $d[19]);
    $sheet->setCellValue("V{$row}", $d[20]);
    $sheet->setCellValue("W{$row}", $d[21]);
    $sheet->setCellValue("X{$row}", $d[22]);
    $sheet->setCellValue("Y{$row}", $d[23]);
    $sheet->setCellValue("Z{$row}", $d[24]);
    $sheet->setCellValue("AA{$row}", $d[25]);
    $sheet->setCellValue("AB{$row}", $d[26]);
    $sheet->setCellValue("AC{$row}", $d[27]);
    $sheet->setCellValue("AD{$row}", $d[28]);
    $sheet->setCellValue("AE{$row}", $d[29]);

    $row += 2;
}

// 內容資料-小計
$data = [
    ["小計", "11", "597", "4", "524", "818", "71", "0", "15", "10", "23", "52", "2478", "17", "68", "0", "9", "0", "0", "5", "246", "1412", "6360"],
    ["小計", "11", "597", "4", "524", "818", "71", "0", "15", "10", "23", "52", "2478", "17", "68", "0", "9", "0", "0", "5", "246", "1412", "6360"],
    ["小計", "11", "597", "4", "524", "818", "71", "0", "15", "10", "23", "52", "2478", "17", "68", "0", "9", "0", "0", "5", "246", "1412", "6360"],
    ["小計", "11", "597", "4", "524", "818", "71", "0", "15", "10", "23", "52", "2478", "17", "68", "0", "9", "0", "0", "5", "246", "1412", "6360"],
    ["小計", "11", "597", "4", "524", "818", "71", "0", "15", "10", "23", "52", "2478", "17", "68", "0", "9", "0", "0", "5", "246", "1412", "6360"],
    ["小計", "11", "597", "4", "524", "818", "71", "0", "15", "10", "23", "52", "2478", "17", "68", "0", "9", "0", "0", "5", "246", "1412", "6360"],
];

$row = 10;
foreach ($data as $d) {
    $sheet->setCellValue("B{$row}", $d[0]);
    $sheet->setCellValue("C{$row}", $d[1]);
    $sheet->setCellValue("E{$row}", $d[2]);
    $sheet->setCellValue("G{$row}", $d[3]);
    $sheet->setCellValue("J{$row}", $d[4]);
    $sheet->setCellValue("K{$row}", $d[5]);
    $sheet->setCellValue("M{$row}", $d[6]);
    $sheet->setCellValue("N{$row}", $d[7]);
    $sheet->setCellValue("O{$row}", $d[8]);
    $sheet->setCellValue("P{$row}", $d[9]);
    $sheet->setCellValue("Q{$row}", $d[10]);
    $sheet->setCellValue("R{$row}", $d[11]);
    $sheet->setCellValue("T{$row}", $d[12]);
    $sheet->setCellValue("V{$row}", $d[13]);
    $sheet->setCellValue("W{$row}", $d[14]);
    $sheet->setCellValue("X{$row}", $d[15]);
    $sheet->setCellValue("Y{$row}", $d[16]);
    $sheet->setCellValue("Z{$row}", $d[17]);
    $sheet->setCellValue("AA{$row}", $d[18]);
    $sheet->setCellValue("AB{$row}", $d[19]);
    $sheet->setCellValue("AC{$row}", $d[20]);
    $sheet->setCellValue("AD{$row}", $d[21]);
    $sheet->setCellValue("AE{$row}", $d[22]);

    $row += 2;
}
$sheet->mergeCells('C10:D10');
$sheet->mergeCells('E10:F10');
$sheet->mergeCells('G10:I10');
$sheet->mergeCells('K10:L10');
$sheet->mergeCells('R10:S10');
$sheet->mergeCells('T10:U10');
$sheet->mergeCells('C12:D12');
$sheet->mergeCells('E12:F12');
$sheet->mergeCells('G12:I12');
$sheet->mergeCells('K12:L12');
$sheet->mergeCells('R12:S12');
$sheet->mergeCells('T12:U12');
$sheet->mergeCells('E14:F14');
$sheet->mergeCells('G14:I14');
$sheet->mergeCells('K14:L14');
$sheet->mergeCells('R14:S14');
$sheet->mergeCells('T14:U14');
$sheet->mergeCells('E16:F16');
$sheet->mergeCells('G16:I16');
$sheet->mergeCells('K16:L16');
$sheet->mergeCells('R16:S16');
$sheet->mergeCells('T16:U16');
$sheet->mergeCells('E18:F18');
$sheet->mergeCells('G18:I18');
$sheet->mergeCells('K18:L18');
$sheet->mergeCells('R18:S18');
$sheet->mergeCells('T18:U18');

$sheet->getStyle('A9:B20')->getFont()->setSize(9);
$sheet->getStyle('B21:C21')->getFont()->setSize(9);
$sheet->getStyle('B9:AE20')->getFont()->setSize(10);
$sheet->getStyle('A9:A20')->getAlignment()->setWrapText(true);
$sheet->getStyle("B9:AE20")->applyFromArray($thinBorder);
$sheet->getStyle('A9:AE20')->applyFromArray($center);
$sheet->mergeCells('C21:AE21');
$sheet->getStyle("A21:AE21")->applyFromArray($thinBorder);
$sheet->getStyle('C10:AE10')->getFont()->setBold(true)->getColor()->setARGB('0000E3');
$sheet->getStyle('C12:AE12')->getFont()->setBold(true)->getColor()->setARGB('0000E3');
$sheet->getStyle('C14:AE14')->getFont()->setBold(true)->getColor()->setARGB('0000E3');
$sheet->getStyle('C16:AE16')->getFont()->setBold(true)->getColor()->setARGB('0000E3');
$sheet->getStyle('C18:AE18')->getFont()->setBold(true)->getColor()->setARGB('0000E3');
$sheet->getStyle('C20:AE20')->getFont()->setBold(true)->getColor()->setARGB('0000E3');
$sheet->getStyle('A9:A10')->applyFromArray($thinBottomMedium);
$sheet->getStyle('A11:A12')->applyFromArray($thinBottomMedium);
$sheet->getStyle('A13:A14')->applyFromArray($thinBottomMedium);
$sheet->getStyle('A15:A16')->applyFromArray($thinBottomMedium);
$sheet->getStyle('A17:A18')->applyFromArray($thinBottomMedium);
$sheet->getStyle('A19:A20')->applyFromArray($thinBottomMedium);
$sheet->getStyle('B10:AE10')->applyFromArray($thinBottomMedium);
$sheet->getStyle('B12:AE12')->applyFromArray($thinBottomMedium);
$sheet->getStyle('B14:AE14')->applyFromArray($thinBottomMedium);
$sheet->getStyle('B16:AE16')->applyFromArray($thinBottomMedium);
$sheet->getStyle('B18:AE18')->applyFromArray($thinBottomMedium);
$sheet->getStyle('B20:AE20')->applyFromArray($thinBottomMedium);
$sheet->getStyle('A21:AE21')->applyFromArray($thinBottomMedium);

// 黃底
$sheet->getStyle('B10:AE10')->applyFromArray($yellowFill2);
$sheet->getStyle('B12:AE12')->applyFromArray($yellowFill2);
$sheet->getStyle('B14:AE14')->applyFromArray($yellowFill2);
$sheet->getStyle('B16:AE16')->applyFromArray($yellowFill2);
$sheet->getStyle('B18:AE18')->applyFromArray($yellowFill2);
$sheet->getStyle('B20:AE20')->applyFromArray($yellowFill2);

// 粉紅底
$sheet->getStyle('R9')->applyFromArray($pinkFill);
$sheet->getStyle('R11')->applyFromArray($pinkFill);
$sheet->getStyle('R13')->applyFromArray($pinkFill);
$sheet->getStyle('R15')->applyFromArray($pinkFill);
$sheet->getStyle('R17')->applyFromArray($pinkFill);
$sheet->getStyle('R19')->applyFromArray($pinkFill);
$sheet->getStyle('T9')->applyFromArray($pinkFill);
$sheet->getStyle('T11')->applyFromArray($pinkFill);
$sheet->getStyle('T13')->applyFromArray($pinkFill);
$sheet->getStyle('T15')->applyFromArray($pinkFill);
$sheet->getStyle('T17')->applyFromArray($pinkFill);
$sheet->getStyle('T19')->applyFromArray($pinkFill);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(7);
$sheet->getColumnDimension('B')->setWidth(7);
$sheet->getColumnDimension('C')->setWidth(5);
$sheet->getColumnDimension('D')->setWidth(5);
$sheet->getColumnDimension('E')->setWidth(6);
$sheet->getColumnDimension('F')->setWidth(6);
$sheet->getColumnDimension('G')->setWidth(7);
$sheet->getColumnDimension('H')->setWidth(7);
$sheet->getColumnDimension('I')->setWidth(7);
$sheet->getColumnDimension('J')->setWidth(8);
$sheet->getColumnDimension('K')->setWidth(6);
$sheet->getColumnDimension('L')->setWidth(6);
$sheet->getColumnDimension('M')->setWidth(7);
$sheet->getColumnDimension('N')->setWidth(6);
$sheet->getColumnDimension('O')->setWidth(7);
$sheet->getColumnDimension('P')->setWidth(7);
$sheet->getColumnDimension('Q')->setWidth(6);
$sheet->getColumnDimension('R')->setWidth(6);
$sheet->getColumnDimension('S')->setWidth(6);
$sheet->getColumnDimension('T')->setWidth(6);
$sheet->getColumnDimension('U')->setWidth(6);
$sheet->getColumnDimension('V')->setWidth(6);
$sheet->getColumnDimension('W')->setWidth(6);
$sheet->getColumnDimension('X')->setWidth(6);
$sheet->getColumnDimension('Y')->setWidth(6);
$sheet->getColumnDimension('Z')->setWidth(6);
$sheet->getColumnDimension('AA')->setWidth(6);
$sheet->getColumnDimension('AB')->setWidth(6);
$sheet->getColumnDimension('AC')->setWidth(6);
$sheet->getColumnDimension('AD')->setWidth(6);
$sheet->getColumnDimension('AE')->setWidth(7);

// 欄高
$sheet->getRowDimension(9)->setRowHeight(30);
$sheet->getRowDimension(10)->setRowHeight(30);
$sheet->getRowDimension(11)->setRowHeight(30);
$sheet->getRowDimension(12)->setRowHeight(30);
$sheet->getRowDimension(13)->setRowHeight(30);
$sheet->getRowDimension(14)->setRowHeight(30);
$sheet->getRowDimension(15)->setRowHeight(30);
$sheet->getRowDimension(16)->setRowHeight(30);
$sheet->getRowDimension(17)->setRowHeight(30);
$sheet->getRowDimension(18)->setRowHeight(30);
$sheet->getRowDimension(19)->setRowHeight(30);
$sheet->getRowDimension(20)->setRowHeight(30);

// 檔名
$filetype = (isset($_GET['type']) && $_GET['type'] == 'week') ? '週報' : '月報';
$period = $_GET['period'] == 'total' ? '累計' : $filetype;
$filename = "當期{$period}-{$republicStartDate}-{$republicEndDate}-員警舉發-整理交通秩序成果統計比較表.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

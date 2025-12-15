<?php
require '../vendor/autoload.php';
require 'styles.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Borders;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("舉發違反道路交通管理事件統計(移公路主管機關裁罰)");

// 分局變數
$organize = "";
switch ($_GET['organize']) {
    case 'one':
        $organize = "一分局";
        break;
    case 'two':
        $organize = "二分局";
        break;
    case 'three':
        $organize = "三分局";
        break;
    case 'traffic':
        $organize = "交通隊";
        break;
    case 'all':
        $organize = "全局";
        break;
}

// 日期變數
$startDate = "114年10月01日";
$endDate   = "114年10月31日";
$makeDate  = "中華民國 114年11月04日";
$makeDate2  = "中華民國 114 年 11 月";
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(12);

// 開頭
$sheet->setCellValue('A1', '公 開 類');
$sheet->setCellValue('Q1', '編報機關');
$sheet->setCellValue('A2', '月    報');
$sheet->setCellValue('B2', '每月終了後10日內編報');
$sheet->setCellValue('Q2', '表    號');
$sheet->setCellValue('R2', '10956-00-01-2');
$sheet->mergeCells('R1:V1');
$sheet->mergeCells('R2:V2');
$sheet->getStyle('A1:A2')->applyFromArray($thinBorder);
$sheet->getStyle('Q1:Q2')->applyFromArray($thinBorder);
$sheet->getStyle('R1:V1')->applyFromArray($thinBorder);
$sheet->getStyle('R2:V2')->applyFromArray($thinBorder);
$sheet->getStyle('A2:V2')->applyFromArray($thinBottomMedium);

// 標題
$sheet->mergeCells('A3:V3');
$sheet->setCellValue('A3', "{$organize}-舉發違反道路交通管理事件統計-移公路主管機關裁罰");
$sheet->getStyle('A3')->applyFromArray($center);
$sheet->getStyle('A3')->getFont()->setSize(16);
// 副標
$sheet->mergeCells('A4:U4');
$sheet->setCellValue('A4', "{$makeDate2}");
$sheet->getStyle('A4')->applyFromArray($center);
$sheet->getStyle('A4')->getFont()->setSize(12);
$sheet->setCellValue('V4', '單位：件');
$sheet->getStyle('A4:V4')->applyFromArray($thinBottomMedium);

// 表頭
$headers = [
    'A9'  => '車',
    'A10' => '輛',
    'A11' => '與',
    'A12' => '舉',
    'A13' => '發',
    'A14' => '方',
    'A15' => '式',

    'B5'  => '項',
    'B6'  => '目',
    'B7'  => '與',
    'B8'  => '適',
    'B9'  => '用',
    'B10' => '條',
    'B11' => '例',

    'C5'  => '總計',

    'D5' => '無照行駛',
    'D14' => '第12條',
    'E5' => '牌照及標明事項之違規',
    'E14' => '第13條',
    'F5' => '牌照行照違規',
    'F14' => '第14條',
    'G5' => '未依規定換照或繳回',
    'G14' => '第15條',
    'H5' => '異動申報、安全及營業設備違規',
    'H14' => '第16條',
    'I5' => '違反定期檢驗',
    'I14' => '第17條',
    'J5' => '基本設備之變換、修復未檢驗',
    'J14' => '第18條',
    'K5' => '依規定裝設行車紀錄器等',
    'K14' => '第18-1條',
    'L5' => '煞車失靈',
    'L14' => '第19條',
    'M5' => '設備損壞之未修復',
    'M14' => '第20條',
    'N5' => '無照駕駛',
    'N14' => '第21條',
    'O5' => '無職業駕照駕駛',
    'O14' => '第21-1條',
    'P5' => '越級駕駛',
    'P14' => '第22條',
    'Q5' => '駕照借人',
    'Q14' => '第23條',
    'R5' => '不依規定接受道路交通安全講習',
    'R14' => '第24條',
    'S5' => '駕照不變更登記、駕照遺毀',
    'S14' => '第25條',
    'T5' => '職業汽車駕駛人，不依規定期限，參加駕駛執照審驗',
    'T14' => '第26條',
    'U5' => '不依規定繳交通行費',
    'U14' => '第27條',
    'V5' => '違反汽車裝載之規定',
    'V14' => '第29條',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->getStyle('A5:V15')->applyFromArray($center);
$sheet->getStyle('A5:B15')->applyFromArray($thinOutlineBorderMedium);
$sheet->getStyle('C5:C15')->applyFromArray($thinOutlineBorderMedium);
foreach (range('D', 'V') as $col) {
    $sheet->mergeCells("{$col}5:{$col}13");
    $sheet->mergeCells("{$col}14:{$col}15");
    $sheet->getStyle("{$col}5:{$col}13")->applyFromArray($thinOutlineBorder);
    $sheet->getStyle("{$col}14:{$col}15")->applyFromArray($thinOutlineBorder);
}
$sheet->mergeCells('C5:C15');
$sheet->getStyle('D14:V15')->applyFromArray($thinBottomMedium);

$headers = [
    'A16' => '合        計',
    'A17' => '汽車',
    'A20' => '逾250cc大型重型機車',
    'A23' => '250cc以下機車',
    'A26' => '動力機械',
    'A27' => '其他',

    'B17' => '小計',
    'B18' => '逕舉',
    'B19' => '攔停',
    'B20' => '小計',
    'B21' => '逕舉',
    'B22' => '攔停',
    'B23' => '小計',
    'B24' => '逕舉',
    'B25' => '攔停',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->mergeCells('A16:B16');
$sheet->mergeCells('A26:B26');
$sheet->mergeCells('A27:B27');
$sheet->mergeCells('A17:A19');
$sheet->mergeCells('A20:A22');
$sheet->mergeCells('A23:A25');
$sheet->getStyle('A16:B27')->applyFromArray($center);
$sheet->getStyle('A16:B16')->applyFromArray($thinBorder);
$sheet->getStyle('B17:B25')->applyFromArray($thinBorder);
$sheet->getStyle('A26:B26')->applyFromArray($thinBorder);
$sheet->getStyle('A20:A22')->applyFromArray($thinOutlineBorder);
$sheet->getStyle('A23:A25')->applyFromArray($thinOutlineBorder);
$sheet->getStyle('C16:C27')->applyFromArray($thinBorder);

// 內容資料
$data = [
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
];

$row = 16;
foreach ($data as $d) {
    $sheet->setCellValue("C{$row}", $d[0]);
    $sheet->setCellValue("D{$row}", $d[1]);
    $sheet->setCellValue("E{$row}", $d[2]);
    $sheet->setCellValue("F{$row}", $d[3]);
    $sheet->setCellValue("G{$row}", $d[4]);
    $sheet->setCellValue("H{$row}", $d[5]);
    $sheet->setCellValue("I{$row}", $d[6]);
    $sheet->setCellValue("J{$row}", $d[7]);
    $sheet->setCellValue("K{$row}", $d[8]);
    $sheet->setCellValue("L{$row}", $d[9]);
    $sheet->setCellValue("M{$row}", $d[10]);
    $sheet->setCellValue("N{$row}", $d[11]);
    $sheet->setCellValue("O{$row}", $d[12]);
    $sheet->setCellValue("P{$row}", $d[13]);
    $sheet->setCellValue("Q{$row}", $d[14]);
    $sheet->setCellValue("R{$row}", $d[15]);
    $sheet->setCellValue("S{$row}", $d[16]);
    $sheet->setCellValue("T{$row}", $d[17]);
    $sheet->setCellValue("U{$row}", $d[18]);
    $sheet->setCellValue("V{$row}", $d[19]);

    $row++;
}
$sheet->getStyle("C16:V" . ($row - 1))->applyFromArray($thinBorder);
$sheet->getStyle('C27:V27')->applyFromArray($thinBottomMedium);
$sheet->getStyle("C16:V" . ($row - 1))->getFont()->setSize(10);

// 表頭
$headers = [
    'A32'  => '車',
    'A33' => '輛',
    'A34' => '與',
    'A35' => '舉',
    'A36' => '發',
    'A37' => '方',
    'A38' => '式',

    'B28'  => '項',
    'B29'  => '目',
    'B30'  => '與',
    'B31'  => '適',
    'B32'  => '用',
    'B33' => '條',
    'B34' => '例',

    'C28' => '違規使用專用車輛或車廂',
    'C37' => '第29-1條',
    'D28' => '超載',
    'D37' => '第29-2條',
    'E28' => '汽車裝載違反規定',
    'E37' => '第30條',
    'F28' => '車輛機件設備、附著物不穩妥或脫落',
    'F37' => '第30-1條',
    'G28' => '未繫安全帶、未安置幼童安全椅等',
    'G37' => '第31條',
    'H28' => '有礙安全駕駛之行為',
    'H37' => '第31-1條',
    'I28' => '無證行駛動力機械',
    'I37' => '第32條',
    'J28' => '違規動力器具，於道路上行駛或使用',
    'J37' => '第32-1條',
    'K28' => '高、快速公路違規',
    'K37' => '第33條',
    'L28' => '疲勞或患病駕駛',
    'L37' => '第34條',
    'M28' => '酒駕、毒駕',
    'M37' => '第35條',
    'N28' => '不依規定裝設、使用酒精鎖',
    'N37' => '第35-1條',
    'O28' => '計程車駕駛人未辦理執業登記',
    'O37' => '第36條',
    'P28' => '違規纜客、拒載、繞道',
    'P37' => '第38條',
    'Q28' => '不在未劃分標線道路之中央右側部分駕車',
    'Q37' => '第39條',
    'R28' => '違反速限',
    'R37' => '第40條',
    'S28' => '按鳴喇叭不依規定',
    'S37' => '第41條',
    'T28' => '不依規定使用燈光者',
    'T37' => '第42條',
    'U28' => '危險駕駛',
    'U37' => '第43條',
    'V28' => '不依規定減速慢行',
    'V37' => '第44條',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->getStyle('A28:V38')->applyFromArray($center);
$sheet->getStyle('A28:B38')->applyFromArray($thinOutlineBorderMedium);
foreach (range('C', 'V') as $col) {
    $sheet->mergeCells("{$col}28:{$col}36");
    $sheet->mergeCells("{$col}37:{$col}38");
    $sheet->getStyle("{$col}28:{$col}36")->applyFromArray($thinOutlineBorder);
    $sheet->getStyle("{$col}37:{$col}38")->applyFromArray($thinOutlineBorder);
}
$sheet->getStyle('C37:V38')->applyFromArray($thinBottomMedium);

$headers = [
    'A39' => '合        計',
    'A40' => '汽車',
    'A43' => '逾250cc大型重型機車',
    'A46' => '250cc以下機車',
    'A49' => '動力機械',
    'A50' => '其他',

    'B40' => '小計',
    'B41' => '逕舉',
    'B42' => '攔停',
    'B43' => '小計',
    'B44' => '逕舉',
    'B45' => '攔停',
    'B46' => '小計',
    'B47' => '逕舉',
    'B48' => '攔停',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->mergeCells('A39:B39');
$sheet->mergeCells('A49:B49');
$sheet->mergeCells('A50:B50');
$sheet->mergeCells('A40:A42');
$sheet->mergeCells('A43:A45');
$sheet->mergeCells('A46:A48');
$sheet->getStyle('A39:B50')->applyFromArray($center);
$sheet->getStyle('A39:B39')->applyFromArray($thinBorder);
$sheet->getStyle('B40:B48')->applyFromArray($thinBorder);
$sheet->getStyle('A49:B50')->applyFromArray($thinBorder);
$sheet->getStyle('A43:A45')->applyFromArray($thinOutlineBorder);
$sheet->getStyle('A46:A48')->applyFromArray($thinOutlineBorder);
$sheet->getStyle('C39:C50')->applyFromArray($thinBorder);

// 內容資料
$data = [
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
];

$row = 39;
foreach ($data as $d) {
    $sheet->setCellValue("C{$row}", $d[0]);
    $sheet->setCellValue("D{$row}", $d[1]);
    $sheet->setCellValue("E{$row}", $d[2]);
    $sheet->setCellValue("F{$row}", $d[3]);
    $sheet->setCellValue("G{$row}", $d[4]);
    $sheet->setCellValue("H{$row}", $d[5]);
    $sheet->setCellValue("I{$row}", $d[6]);
    $sheet->setCellValue("J{$row}", $d[7]);
    $sheet->setCellValue("K{$row}", $d[8]);
    $sheet->setCellValue("L{$row}", $d[9]);
    $sheet->setCellValue("M{$row}", $d[10]);
    $sheet->setCellValue("N{$row}", $d[11]);
    $sheet->setCellValue("O{$row}", $d[12]);
    $sheet->setCellValue("P{$row}", $d[13]);
    $sheet->setCellValue("Q{$row}", $d[14]);
    $sheet->setCellValue("R{$row}", $d[15]);
    $sheet->setCellValue("S{$row}", $d[16]);
    $sheet->setCellValue("T{$row}", $d[17]);
    $sheet->setCellValue("U{$row}", $d[18]);
    $sheet->setCellValue("V{$row}", $d[19]);

    $row++;
}
$sheet->getStyle("C39:V" . ($row - 1))->applyFromArray($thinBorder);
$sheet->getStyle('C50:V50')->applyFromArray($thinBottomMedium);
$sheet->getStyle("C39:V" . ($row - 1))->getFont()->setSize(10);

// 表頭
$headers = [
    'A55'  => '車',
    'A56' => '輛',
    'A57' => '與',
    'A58' => '舉',
    'A59' => '發',
    'A60' => '方',
    'A61' => '式',

    'B51'  => '項',
    'B52'  => '目',
    'B53'  => '與',
    'B54'  => '適',
    'B55'  => '用',
    'B56' => '條',
    'B57' => '例',

    'C51' => '爭道行駛',
    'C60' => '第45條',
    'D51' => '交會不依規定',
    'D60' => '第46條',
    'E51' => '超車不依規定',
    'E60' => '第47條',
    'F51' => '轉彎或變換車道不依規定',
    'F60' => '第48條',
    'G51' => '迴車不依規定',
    'G60' => '第49條',
    'H51' => '倒車不依規定',
    'H60' => '第50條',
    'I51' => '上下坡不依規定',
    'I60' => '第51條',
    'J51' => '行經渡口不依規定',
    'J60' => '第52條',
    'K51' => '闖紅燈',
    'K60' => '第53條',
    'L51' => '大眾捷運系統車輛共用通行交岔路口闖紅燈',
    'L60' => '第53-1條',
    'M51' => '闖越平交道或在平交道違規',
    'M60' => '第54條',
    'N51' => '違規臨時停車',
    'N60' => '第55條',
    'O51' => '違規停車',
    'O60' => '第56條',
    'P51' => '違規開、關車門',
    'P60' => '第56-1條',
    'Q51' => '汽車買賣業、修理業違規停車',
    'Q60' => '第57條',
    'R51' => '路口淨空',
    'R60' => '第58條',
    'S51' => '駕車時故障未依規定處理',
    'S60' => '第59條',
    'T51' => '汽車駕駛人違規',
    'T60' => '第60條',
    'U51' => '違規肇事',
    'U60' => '第61條',
    'V51' => '肇事後處理不當',
    'V60' => '第62條',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->getStyle('A51:V61')->applyFromArray($center);
$sheet->getStyle('A51:B61')->applyFromArray($thinOutlineBorderMedium);
foreach (range('C', 'V') as $col) {
    $sheet->mergeCells("{$col}51:{$col}59");
    $sheet->mergeCells("{$col}60:{$col}61");
    $sheet->getStyle("{$col}51:{$col}59")->applyFromArray($thinOutlineBorder);
    $sheet->getStyle("{$col}60:{$col}61")->applyFromArray($thinOutlineBorder);
}
$sheet->getStyle('C60:V61')->applyFromArray($thinBottomMedium);

$headers = [
    'A62' => '合        計',
    'A63' => '汽車',
    'A66' => '逾273cc大型重型機車',
    'A69' => '273cc以下機車',
    'A72' => '動力機械',
    'A73' => '其他',

    'B63' => '小計',
    'B64' => '逕舉',
    'B65' => '攔停',
    'B66' => '小計',
    'B67' => '逕舉',
    'B68' => '攔停',
    'B69' => '小計',
    'B70' => '逕舉',
    'B71' => '攔停',
];

foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->mergeCells('A62:B62');
$sheet->mergeCells('A72:B72');
$sheet->mergeCells('A73:B73');
$sheet->mergeCells('A63:A65');
$sheet->mergeCells('A66:A68');
$sheet->mergeCells('A69:A71');
$sheet->getStyle('A62:B73')->applyFromArray($center);
$sheet->getStyle('A62:B62')->applyFromArray($thinBorder);
$sheet->getStyle('B63:B71')->applyFromArray($thinBorder);
$sheet->getStyle('A72:B73')->applyFromArray($thinBorder);
$sheet->getStyle('A66:A68')->applyFromArray($thinOutlineBorder);
$sheet->getStyle('A69:A71')->applyFromArray($thinOutlineBorder);
$sheet->getStyle('C62:C73')->applyFromArray($thinBorder);

// 內容資料
$data = [
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
];

$row = 62;
foreach ($data as $d) {
    $sheet->setCellValue("C{$row}", $d[0]);
    $sheet->setCellValue("D{$row}", $d[1]);
    $sheet->setCellValue("E{$row}", $d[2]);
    $sheet->setCellValue("F{$row}", $d[3]);
    $sheet->setCellValue("G{$row}", $d[4]);
    $sheet->setCellValue("H{$row}", $d[5]);
    $sheet->setCellValue("I{$row}", $d[6]);
    $sheet->setCellValue("J{$row}", $d[7]);
    $sheet->setCellValue("K{$row}", $d[8]);
    $sheet->setCellValue("L{$row}", $d[9]);
    $sheet->setCellValue("M{$row}", $d[10]);
    $sheet->setCellValue("N{$row}", $d[11]);
    $sheet->setCellValue("O{$row}", $d[12]);
    $sheet->setCellValue("P{$row}", $d[13]);
    $sheet->setCellValue("Q{$row}", $d[14]);
    $sheet->setCellValue("R{$row}", $d[15]);
    $sheet->setCellValue("S{$row}", $d[16]);
    $sheet->setCellValue("T{$row}", $d[17]);
    $sheet->setCellValue("U{$row}", $d[18]);
    $sheet->setCellValue("V{$row}", $d[19]);

    $row++;
}
$sheet->getStyle("C62:V" . ($row - 1))->applyFromArray($thinBorder);
$sheet->getStyle('C73:V73')->applyFromArray($thinBottomMedium);
$sheet->getStyle("C62:V" . ($row - 1))->getFont()->setSize(10);

// 灰底
$sheet->getStyle('C16:C27')->applyFromArray($grayFill);
$sheet->getStyle('D16:V16')->applyFromArray($grayFill);
$sheet->getStyle('D17:V17')->applyFromArray($grayFill);
$sheet->getStyle('D20:V20')->applyFromArray($grayFill);
$sheet->getStyle('C38:V38')->applyFromArray($grayFill);
$sheet->getStyle('C39:V39')->applyFromArray($grayFill);
$sheet->getStyle('C40:V40')->applyFromArray($grayFill);
$sheet->getStyle('C43:V43')->applyFromArray($grayFill);
$sheet->getStyle('C49:V49')->applyFromArray($grayFill);
$sheet->getStyle('C62:V62')->applyFromArray($grayFill);
$sheet->getStyle('C63:V63')->applyFromArray($grayFill);
$sheet->getStyle('C66:V66')->applyFromArray($grayFill);
$sheet->getStyle('C69:V69')->applyFromArray($grayFill);

// // 欄寬
$sheet->getColumnDimension('A')->setWidth(12);
$sheet->getColumnDimension('B')->setWidth(12);
$sheet->getColumnDimension('C')->setWidth(13);
$sheet->getColumnDimension('D')->setWidth(10);


require 'achievement_highway_02.php';

// 檔名
$filename = "{$republicStartDate}-{$republicEndDate}-{$organize}-舉發違反道路交通管理事件統計-移公路主管機關裁罰.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

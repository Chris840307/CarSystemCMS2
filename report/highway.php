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
$sheet->setTitle("全局舉發違反高速公路及快速公路管制規定成果統計表");

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
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(12);

// 開頭
$sheet->setCellValue('A1', '公 開 類');
$sheet->setCellValue('V1', '編報機關');
$sheet->setCellValue('A2', '月    報');
$sheet->setCellValue('B2', '每月終了後10日內編報');
$sheet->setCellValue('V2', '表    號');
$sheet->setCellValue('W2', '1736 -01-02-2');
$sheet->mergeCells('W2:X2');
$sheet->getStyle('A1:A2')->applyFromArray($thinBorder);
$sheet->getStyle('V1:V2')->applyFromArray($thinBorder);
$sheet->getStyle('W2:X2')->applyFromArray($thinBorder);
$sheet->getStyle('A2:X2')->applyFromArray($thinBottomMedium);

// 標題
$sheet->mergeCells('A3:X3');
$sheet->setCellValue('A3', '全局舉發違反高速公路及快速公路管制規定成果統計表');
$sheet->getStyle('A3')->applyFromArray($center);
$sheet->getStyle('A3')->getFont()->setSize(16);
// 副標
$sheet->mergeCells('A4:W4');
$sheet->setCellValue('A4', "統計範圍：{$startDate} 到 {$endDate}");
$sheet->getStyle('A4')->applyFromArray($center);
$sheet->getStyle('A4')->getFont()->setSize(12);
$sheet->setCellValue('X4', '單位：件');
$sheet->getStyle('A4:X4')->applyFromArray($thinBottomMedium);

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

    'C6'  => '舉',
    'C8'  => '發',
    'C10' => '總',
    'C12' => '件',
    'C14' => '數',

    'E5' => '移',
    'L5' => '公',
    'W5' => '路',

    'D6' => '車',
    'D7' => '速',
    'D8' => '超',
    'D9' => '過',
    'D10' => '最',
    'D11' => '高',
    'D12' => '速',
    'D13' => '限',
    'D14' => '33條第1',
    'D15' => '項第1款',

    'E6' => '車',
    'E7' => '速',
    'E8' => '低',
    'E9' => '於',
    'E10' => '最',
    'E11' => '低',
    'E12' => '速',
    'E13' => '限',
    'E14' => '33條第1',
    'E15' => '項第1款',

    'F6' => '未與',
    'F7' => '前車',
    'F8' => '保持',
    'F9' => '安全',
    'F10' => '距離',
    'F11' => '',
    'F12' => '',
    'F13' => '',
    'F14' => '33條第1',
    'F15' => '項第2款',

    'G6' => '慢速',
    'G7' => '小型',
    'G8' => '車不',
    'G9' => '依規',
    'G10' => '定車',
    'G11' => '道行',
    'G12' => '駛',
    'G13' => '',
    'G14' => '33條第1',
    'G15' => '項第3款',

    'H6' => '大型',
    'H7' => '車不',
    'H8' => '依規',
    'H9' => '定車',
    'H10' => '道行',
    'H11' => '駛',
    'H12' => '',
    'H13' => '',
    'H14' => '33條第1',
    'H15' => '項第3款',

    'I6' => '未依',
    'I7' => '規定',
    'I8' => '變換',
    'I9' => '車道',
    'I10' => '',
    'I11' => '',
    'I12' => '',
    'I13' => '',
    'I14' => '33條第1',
    'I15' => '項第4款',

    'J6' => '車',
    'J7' => '內',
    'J8' => '有',
    'J9' => '站',
    'J10' => '立',
    'J11' => '乘',
    'J12' => '客',
    'J13' => '',
    'J14' => '33條第1',
    'J15' => '項第5款',

    'K6' => '同向前100',
    'K7' => '公尺內有',
    'K8' => '車輛行駛',
    'K9' => '仍使用遠',
    'K10' => '光燈隧道',
    'K11' => '內未開頭',
    'K12' => '燈夜間未',
    'K13' => '使用燈光',
    'K14' => '33條第1',
    'K15' => '項第6款',

    'L6' => '大型',
    'L7' => '重型',
    'L8' => '機車',
    'L9' => '未開啟',
    'L10' => '頭燈',
    'L11' => '',
    'L12' => '',
    'L13' => '',
    'L14' => '33條第1',
    'L15' => '項第6款',

    'M6' => '違規',
    'M7' => '超車',
    'M8' => '迴車',
    'M9' => '倒車',
    'M10' => '逆向',
    'M11' => '行駛',
    'M12' => '',
    'M13' => '',
    'M14' => '33條第1',
    'M15' => '項第7款',

    'N6' => '違',
    'N7' => '規',
    'N8' => '減',
    'N9' => '速',
    'N10' => '',
    'N11' => '',
    'N12' => '',
    'N13' => '',
    'N14' => '33條第1',
    'N15' => '項第8款',

    'O6' => '於路肩外',
    'O7' => '收費站區',
    'O8' => '中央分隔',
    'O9' => '帶隧道內',
    'O10' => '交流道違',
    'O11' => '規停車或',
    'O12' => '臨時停車',
    'O13' => '',
    'O14' => '33條第1',
    'O15' => '項第8款',

    'P6' => '小型',
    'P7' => '車違',
    'P8' => '規行',
    'P9' => '駛路',
    'P10' => '肩',
    'P11' => '',
    'P12' => '',
    'P13' => '',
    'P14' => '33條第1',
    'P15' => '項第9款',

    'Q6' => '大型',
    'Q7' => '車違',
    'Q8' => '規行',
    'Q9' => '駛路',
    'Q10' => '肩',
    'Q11' => '',
    'Q12' => '',
    'Q13' => '',
    'Q14' => '33條第1',
    'Q15' => '項第9款',

    'R6' => '未依',
    'R7' => '施工',
    'R8' => '之安',
    'R9' => '全設',
    'R10' => '施指',
    'R11' => '示行',
    'R12' => '駛',
    'R13' => '',
    'R14' => '33條第1',
    'R15' => '項第10款',

    'S6' => '裝置',
    'S7' => '貨物',
    'S8' => '未依',
    'S9' => '規定',
    'S10' => '嚴密',
    'S11' => '覆蓋',
    'S12' => '',
    'S13' => '',
    'S14' => '33條第1',
    'S15' => '項第11款',

    'T6' => '裝置',
    'T7' => '貨物',
    'T8' => '未依',
    'T9' => '規定',
    'T10' => '捆紥',
    'T11' => '牢固',
    'T12' => '',
    'T13' => '',
    'T14' => '33條第1',
    'T15' => '項第11款',

    'U6' => '行駛高',
    'U7' => '快速公',
    'U8' => '路未依',
    'U9' => '標誌標',
    'U10' => '線號誌',
    'U11' => '指示行',
    'U12' => '車',
    'U13' => '',
    'U14' => '33條第1',
    'U15' => '項第12款',

    'V6' => '進入',
    'V7' => '或行',
    'V8' => '駛高',
    'V9' => '快速',
    'V10' => '公路',
    'V11' => '禁止',
    'V12' => '通行',
    'V13' => '路段',
    'V14' => '33條第1',
    'V15' => '項第13款',

    'W6' => '以不',
    'W7' => '當方',
    'W8' => '式迫',
    'W9' => '使前',
    'W10' => '車讓',
    'W11' => '道',
    'W12' => '',
    'W13' => '',
    'W14' => '33條第1',
    'W15' => '項第14款',

    'X6' => '向車',
    'X7' => '外丟',
    'X8' => '棄物',
    'X9' => '品或',
    'X10' => '廢棄',
    'X11' => '物',
    'X12' => '',
    'X13' => '',
    'X14' => '33條第1',
    'X15' => '項第15款',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->getStyle('A5:X15')->applyFromArray($center);
$sheet->getStyle('A5:B15')->applyFromArray($thinOutlineBorderMedium);
$sheet->getStyle('C5:C15')->applyFromArray($thinOutlineBorderMedium);
foreach (range('D', 'X') as $col) {
    $sheet->getStyle("{$col}6:{$col}15")->applyFromArray($thinOutlineBorder);
}

$headers = [
    'A16' => '合        計',
    'A18' => '汽車',
    'A20' => '逾250cc',
    'A21' => '大型重',
    'A22' => '型機車',
    'A23' => '250cc',
    'A24' => '以下',
    'A25' => '機車',

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
$sheet->getStyle('A16:B25')->applyFromArray($center);
$sheet->getStyle('A16:B16')->applyFromArray($thinBorder);
$sheet->getStyle('B17:B25')->applyFromArray($thinBorder);
$sheet->getStyle('A20:A22')->applyFromArray($thinOutlineBorder);
$sheet->getStyle('C16:C25')->applyFromArray($thinOutlineBorderMedium);

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
    $sheet->setCellValue("W{$row}", $d[20]);
    $sheet->setCellValue("X{$row}", $d[21]);

    $row++;
}
$sheet->getStyle("C16:X" . ($row - 1))->applyFromArray($thinBorder);
$sheet->getStyle('D15:X15')->applyFromArray($thinBottomMedium);
$sheet->getStyle('D25:X25')->applyFromArray($thinBottomMedium);
$sheet->getStyle("C16:X" . ($row - 1))->getFont()->setSize(10);

// 表頭
$headers = [
    'A30'  => '車',
    'A31' => '輛',
    'A32' => '與',
    'A33' => '舉',
    'A34' => '發',
    'A35' => '方',
    'A36' => '式',

    'B26'  => '項',
    'B27'  => '目',
    'B28'  => '與',
    'B29'  => '適',
    'B30'  => '用',
    'B31' => '條',
    'B32' => '例',

    'C27' => '車輪',
    'C28' => '輪胎',
    'C29' => '膠皮',
    'C30' => '或車',
    'C31' => '輛機',
    'C32' => '件脫',
    'C33' => '落',
    'C34' => '',
    'C35' => '33條',
    'C36' => '第1項第16款',

    'F26' => '監',
    'K26' => '理',
    'P26' => '機',
    'W26' => '關',

    'D27' => '輪胎',
    'D28' => '胎紋',
    'D29' => '深度',
    'D30' => '不符',
    'D31' => '規定',
    'D32' => '',
    'D33' => '',
    'D34' => '',
    'D35' => '33條',
    'D36' => '第1項第17款',

    'E27' => '利用內側',
    'E28' => '車道超車',
    'E29' => '後如有安',
    'E30' => '全距離未',
    'E31' => '駛回原車',
    'E32' => '道或未以',
    'E33' => '最高速限',
    'E34' => '行駛',
    'E35' => '33條',
    'E36' => '第2項',

    'F27' => '擅自開',
    'F28' => '啟或穿',
    'F29' => '越中央',
    'F30' => '分隔帶',
    'F31' => '分項設',
    'F32' => '施',
    'F33' => '',
    'F34' => '',
    'F35' => '33條',
    'F36' => '第3項',

    'G27' => '車輛',
    'G28' => '輪胎',
    'G29' => '粘附',
    'G30' => '泥沙',
    'G31' => '致污',
    'G32' => '染路',
    'G33' => '面',
    'G34' => '',
    'G35' => '33條',
    'G36' => '第3項',

    'H27' => '拖拉',
    'H28' => '故障',
    'H29' => '車輛',
    'H30' => '使用',
    'H31' => '非鋼',
    'H32' => '質連',
    'H33' => '杆',
    'H34' => '',
    'H35' => '33條',
    'H36' => '第3項',

    'I27' => '行駛高',
    'I28' => '快速公',
    'I29' => '路途中',
    'I30' => '缺水缺',
    'I31' => '電缺燃',
    'I32' => '料',
    'I33' => '',
    'I34' => '',
    'I35' => '33條',
    'I36' => '第3項',

    'J27' => '駛進收費',
    'J28' => '站繳費未',
    'J29' => '依標誌標',
    'J30' => '線號誌指',
    'J31' => '示過站繳',
    'J32' => '費',
    'J33' => '',
    'J34' => '',
    'J35' => '33條',
    'J36' => '第3項',

    'K27' => '停放服',
    'K28' => '務區休',
    'K29' => '息站內',
    'K30' => '之車輛',
    'K31' => '逾4小時',
    'K32' => '',
    'K33' => '',
    'K34' => '',
    'K35' => '33條',
    'K36' => '第3項',

    'L27' => '未經登記',
    'L28' => '許可之拖',
    'L29' => '吊車擅自',
    'L30' => '在高快速',
    'L31' => '公路沿線',
    'L32' => '路權範圍',
    'L33' => '內營業',
    'L34' => '',
    'L35' => '33條',
    'L36' => '第3項',

    'M27' => '裝載超',
    'M28' => '長物品',
    'M29' => '後伸部',
    'M30' => '分遮擋',
    'M31' => '車後燈',
    'M32' => '光號牌',
    'M33' => '',
    'M34' => '',
    'M35' => '33條',
    'M36' => '第3項',

    'N27' => '經核准',
    'N28' => '之大型',
    'N29' => '重型機',
    'N30' => '車行駛',
    'N31' => '快速公',
    'N32' => '路同車',
    'N33' => '道併駛',
    'N34' => '',
    'N35' => '33條',
    'N36' => '第3項',

    'O27' => '大型重機',
    'O28' => '駕駛人',
    'O29' => '及乘客',
    'O30' => '未配戴',
    'O31' => '全罩式',
    'O32' => '或露臉式',
    'O33' => '安全帽',
    'O34' => '',
    'O35' => '33條',
    'O36' => '第3項',

    'P27' => '行人部',
    'P28' => '隊行車',
    'P29' => '或演習',
    'P30' => '慢車進',
    'P31' => '入高快',
    'P32' => '速公路',
    'P33' => '',
    'P34' => '',
    'P35' => '33條',
    'P36' => '第4項',

    'Q27' => '禁止進',
    'Q28' => '入之人',
    'Q29' => '員車輛',
    'Q30' => '或動力',
    'Q31' => '機械進',
    'Q32' => '入高快',
    'Q33' => '速公路',
    'Q34' => '',
    'Q35' => '33條',
    'Q36' => '第4項',

    'R27' => '550cc以上',
    'R28' => '大型重型',
    'R29' => '機車行駛',
    'R30' => '高速公路',
    'R31' => '行駛未經',
    'R32' => '公告允許',
    'R33' => '之路段',
    'R34' => '',
    'R35' => '92條第7',
    'R36' => '項第1款',

    'S27' => '550cc以上',
    'S28' => '大型重型',
    'S29' => '機車行駛',
    'S30' => '高速公路',
    'S31' => '未依公告',
    'S32' => '允許時段',
    'S33' => '規定行駛',
    'S34' => '',
    'S35' => '92條第7',
    'S36' => '項第2款',

    'T27' => '550cc以上',
    'T28' => '大型重型',
    'T29' => '機車行駛',
    'T30' => '高速公路',
    'T31' => '領有駕照',
    'T32' => '未符合',
    'T33' => '92條第2項',
    'T34' => '規定',
    'T35' => '92條第7',
    'T36' => '項第3款',

    'U27' => '550cc以上',
    'U28' => '大型重型',
    'U29' => '機車行駛',
    'U30' => '高速公路',
    'U31' => '同車道',
    'U32' => '併駛超車',
    'U33' => '或未依規定',
    'U34' => '使用路肩',
    'U35' => '92條第7',
    'U36' => '項第4款',

    'V27' => '550cc以上',
    'V28' => '大型重型',
    'V29' => '機車行駛',
    'V30' => '高速公路',
    'V31' => '未依規定',
    'V32' => '附載人員',
    'V33' => '或物品',
    'V34' => '',
    'V35' => '92條第7',
    'V36' => '項第5款',

    'W27' => '550cc以上',
    'W28' => '大型重型',
    'W29' => '機車行駛',
    'W30' => '高速公路',
    'W31' => '未依規定',
    'W32' => '戴安全帽',
    'W33' => '',
    'W34' => '',
    'W35' => '92條第7',
    'W36' => '項第6款',

    'X27' => '其他',
    'X28' => '違反33條',
    'X29' => '未列之',
    'X30' => '條款',
    'X31' => '',
    'X32' => '',
    'X33' => '',
    'X34' => '',
    'X35' => '',
    'X36' => '',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->getStyle('A26:X36')->applyFromArray($center);
$sheet->getStyle('A26:B36')->applyFromArray($thinOutlineBorderMedium);
$sheet->getStyle('A37:B46')->applyFromArray($thinOutlineBorderMedium);
foreach (range('C', 'X') as $col) {
    $sheet->getStyle("{$col}27:{$col}36")->applyFromArray($thinOutlineBorder);
}
$sheet->getStyle('C27:C36')->getFont()->getColor()->setARGB(Color::COLOR_RED);
$sheet->getStyle('D27:D36')->getFont()->getColor()->setARGB(Color::COLOR_RED);

$headers = [
    'A37' => '合        計',
    'A39' => '汽車',
    'A41' => '逾250cc',
    'A42' => '大型重',
    'A43' => '型機車',
    'A44' => '250cc',
    'A45' => '以下',
    'A46' => '機車',

    'B38' => '小計',
    'B39' => '逕舉',
    'B40' => '攔停',
    'B41' => '小計',
    'B42' => '逕舉',
    'B43' => '攔停',
    'B44' => '小計',
    'B45' => '逕舉',
    'B46' => '攔停',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->mergeCells('A37:B37');
$sheet->getStyle('A37:B46')->applyFromArray($center);
$sheet->getStyle('A37:B37')->applyFromArray($thinBorder);
$sheet->getStyle('B38:B46')->applyFromArray($thinBorder);
$sheet->getStyle('A41:A43')->applyFromArray($thinOutlineBorder);
$sheet->getStyle('A37:B46')->applyFromArray($thinOutlineBorderMedium);

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
];

$row = 37;
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
    $sheet->setCellValue("W{$row}", $d[20]);
    $sheet->setCellValue("X{$row}", $d[21]);

    $row++;
}
$sheet->getStyle("C37:X" . ($row - 1))->applyFromArray($thinBorder);
$sheet->getStyle('C36:X36')->applyFromArray($thinBottomMedium);
$sheet->getStyle('C46:X46')->applyFromArray($thinBottomMedium);
$sheet->getStyle("C37:X" . ($row - 1))->getFont()->setSize(10);

$sheet->setCellValue('A47', "填  表");
$sheet->setCellValue('E47', "審  核");
$sheet->setCellValue('J47', "業務主管人員");
$sheet->setCellValue('O47', "機關首長");
$sheet->setCellValue('U47', "{$makeDate}  編製");
$sheet->setCellValue('A49', "資料來源：新北市、桃園市、臺中市、臺南市、高雄市、新竹縣、苗栗縣、彰化縣、南投縣、雲林縣、嘉義縣、屏東縣、基隆市、新竹市及國道公路警察局。");
$sheet->setCellValue('A50', "填表說明：(一)本表編製1式2份，先送會計室(統計室)會核，並經機關首長核章後，1份送會計室(統計室)，1份自存外，本表於規定期限內由網際網路線上傳送至內政部警政署警政統計資料庫。");
$sheet->setCellValue('A51', "          (二)新表自統計期104年7月1日起適用。");
$sheet->getStyle("C47:X50")->getFont()->setSize(10);

// 灰底
$sheet->getStyle('C16:C25')->applyFromArray($grayFill);
$sheet->getStyle('D16:X16')->applyFromArray($grayFill);
$sheet->getStyle('D17:X17')->applyFromArray($grayFill);
$sheet->getStyle('D20:X20')->applyFromArray($grayFill);
$sheet->getStyle('D23:X23')->applyFromArray($grayFill);
$sheet->getStyle('D37:X37')->applyFromArray($grayFill);
$sheet->getStyle('D38:X38')->applyFromArray($grayFill);
$sheet->getStyle('D41:X41')->applyFromArray($grayFill);
$sheet->getStyle('D44:X44')->applyFromArray($grayFill);

// // 欄寬
$sheet->getColumnDimension('A')->setWidth(9);
$sheet->getColumnDimension('B')->setWidth(8);
$sheet->getColumnDimension('C')->setWidth(13);
$sheet->getColumnDimension('D')->setWidth(10);

// 檔名
$filename = "{$republicStartDate}-{$republicEndDate}-{$organize}-舉發違反高速公路及快速公路管制規定成果統計表.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

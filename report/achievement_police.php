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
$sheet->setTitle("舉發違反道路交通管理事件統計(警察機關裁罰)");

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
$sheet->setCellValue('R2', '10956-00-02-2');
$sheet->mergeCells('R1:V1');
$sheet->mergeCells('R2:V2');
$sheet->getStyle('A1:A2')->applyFromArray($thinBorder);
$sheet->getStyle('Q1:Q2')->applyFromArray($thinBorder);
$sheet->getStyle('R1:V1')->applyFromArray($thinBorder);
$sheet->getStyle('R2:V2')->applyFromArray($thinBorder);
$sheet->getStyle('A2:V2')->applyFromArray($thinBottomMedium);

// 標題
$sheet->mergeCells('A3:V3');
$sheet->setCellValue('A3', "{$organize}-舉發違反道路交通管理事件統計-警察機關裁罰");
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

    'D5' => '計程車駕駛人之消極資格',
    'D14' => '第37條',
    'E5' => '慢車無照及個人行動器具違規',
    'E14' => '第69條',
    'F5' => '行駛淘汰車',
    'F14' => '第70條',
    'G5' => '電輔車未審驗或未貼標章',
    'G14' => '第71條',
    'H5' => '微電車牌照、審驗違規',
    'H14' => '第71-1條',
    'I5' => '微電車牌照違規',
    'I14' => '第71-2條',
    'J5' => '慢車違規變更裝置',
    'J14' => '第72條',
    'K5' => '慢車超速',
    'K14' => '第72-1條',
    'L5' => '未滿十四歲駕駛慢車及租賃業者之處罰',
    'L14' => '第72-2條',
    'M5' => '慢車違規',
    'M14' => '第73條',
    'N5' => '慢車不依規定行駛、禮讓',
    'N14' => '第74條',
    'O5' => '慢車闖平交道',
    'O14' => '第75條',
    'P5' => '慢車違規載客、貨',
    'P14' => '第76條',
    'Q5' => '行人違規',
    'Q14' => '第78條',
    'R5' => '行人闖越平交道',
    'R14' => '第80條',
    'S5' => '行人攀跳行車',
    'S14' => '第81條',
    'T5' => '違規攬客',
    'T14' => '第80-1條',
    'U5' => '道路障礙',
    'U14' => '第82條',
    'V5' => '阻礙交通',
    'V14' => '第83條',
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
    'A17' => '微型電動二輪車',
    'A20' => '其他慢車',
    'A26' => '其他',

    'B17' => '小計',
    'B18' => '逕舉',
    'B19' => '攔停',
    'B20' => '小計',
    'B21' => '腳踏自行車',
    'B22' => '電動輔助自行車',
    'B23' => '人力行駛車輛',
    'B24' => '獸力行駛車輛',
    'B25' => '個人行動器具',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->mergeCells('A16:B16');
$sheet->mergeCells('A26:B26');
$sheet->mergeCells('A17:A19');
$sheet->mergeCells('A20:A25');
$sheet->getStyle('A16:B26')->applyFromArray($center);
$sheet->getStyle('A16:B16')->applyFromArray($thinBorder);
$sheet->getStyle('B17:B25')->applyFromArray($thinBorder);
$sheet->getStyle('A20:A25')->applyFromArray($thinOutlineBorder);
$sheet->getStyle('C16:C26')->applyFromArray($thinBorder);

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
$sheet->getStyle('C26:V26')->applyFromArray($thinBottomMedium);
$sheet->getStyle("C16:V" . ($row - 1))->getFont()->setSize(10);

// 表頭
$headers = [
    'A31'  => '車',
    'A32' => '輛',
    'A33' => '與',
    'A34' => '舉',
    'A35' => '發',
    'A36' => '方',
    'A37' => '式',

    'B27'  => '項',
    'B28'  => '目',
    'B29'  => '與',
    'B30'  => '適',
    'B31'  => '用',
    'B32' => '條',
    'B33' => '例',

    'C27' => '疏縱或牽繫禽、畜、寵物在道路奔走，妨害交通者',
    'C36' => '第84條',
    'D27' => '其他警察機關處罰未列之項目',
    'D36' => '第37條、第69-84條',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->getStyle('A27:V37')->applyFromArray($center);
$sheet->getStyle('A27:B37')->applyFromArray($thinOutlineBorderMedium);
foreach (range('C', 'V') as $col) {
    $sheet->mergeCells("{$col}27:{$col}35");
    $sheet->mergeCells("{$col}36:{$col}37");
    $sheet->getStyle("{$col}27:{$col}35")->applyFromArray($thinOutlineBorder);
    $sheet->getStyle("{$col}36:{$col}37")->applyFromArray($thinOutlineBorder);
}
$sheet->getStyle('C36:V37')->applyFromArray($thinBottomMedium);

$headers = [
    'A38' => '合        計',
    'A39' => '微型電動二輪車',
    'A42' => '其他慢車',
    'A48' => '其他',

    'B39' => '小計',
    'B40' => '逕舉',
    'B41' => '攔停',
    'B42' => '小計',
    'B43' => '腳踏自行車',
    'B44' => '電動輔助自行車',
    'B45' => '人力行駛車輛',
    'B46' => '獸力行駛車輛',
    'B47' => '個人行動器具',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
    $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet->mergeCells('A38:B38');
$sheet->mergeCells('A48:B48');
$sheet->mergeCells('A39:A41');
$sheet->mergeCells('A42:A47');
$sheet->getStyle('A38:B48')->applyFromArray($center);
$sheet->getStyle('A38:B38')->applyFromArray($thinBorder);
$sheet->getStyle('A48:B48')->applyFromArray($thinBorder);
$sheet->getStyle('B39:B47')->applyFromArray($thinBorder);
$sheet->getStyle('A42:A47')->applyFromArray($thinOutlineBorder);

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
];

$row = 38;
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
$sheet->getStyle("C38:V" . ($row - 1))->applyFromArray($thinBorder);
$sheet->getStyle('C48:V48')->applyFromArray($thinBottomMedium);
$sheet->getStyle("C38:V" . ($row - 1))->getFont()->setSize(10);

$sheet->setCellValue('A49', "填  表");
$sheet->setCellValue('G49', "審  核");
$sheet->setCellValue('K49', "業務主管人員");
$sheet->setCellValue('K50', "主辦統計人員");
$sheet->setCellValue('O49', "機關首長");
$sheet->setCellValue('T49', "{$makeDate}編製");
$sheet->setCellValue('A51', "資料來源：本局轄區內取締違反道路交通管理事件之舉發件數彙編。");
$sheet->setCellValue('A52', "填表說明：本表編製1式3份，先送會計室(統計室)會核，並經機關首長核章後，1份送各直轄市、縣(市)政府主計處，1份送會計室(統計室) ，1份自存外，本表於規定期限內由網際網路線上傳送至內政部警政署警政統計資料庫。");

// 灰底
$sheet->getStyle('C16:C26')->applyFromArray($grayFill);
$sheet->getStyle('D16:V16')->applyFromArray($grayFill);
$sheet->getStyle('D17:V17')->applyFromArray($grayFill);
$sheet->getStyle('D20:V20')->applyFromArray($grayFill);
$sheet->getStyle('C38:V38')->applyFromArray($grayFill);
$sheet->getStyle('C39:V39')->applyFromArray($grayFill);
$sheet->getStyle('C42:V42')->applyFromArray($grayFill);

// // 欄寬
$sheet->getColumnDimension('A')->setWidth(12);
$sheet->getColumnDimension('B')->setWidth(15);
$sheet->getColumnDimension('C')->setWidth(13);
$sheet->getColumnDimension('D')->setWidth(10);

// 檔名
$filename = "{$republicStartDate}-{$republicEndDate}-{$organize}-舉發違反道路交通管理事件統計-警察機關裁罰.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

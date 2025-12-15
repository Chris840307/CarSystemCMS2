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

$sheet2 = $spreadsheet->createSheet();
$sheet2->setTitle("舉發違反道路交通管理事件統計(移公路主管機關裁罰)");

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
$sheet2->setCellValue('A1', '公 開 類');
$sheet2->setCellValue('Q1', '編報機關');
$sheet2->setCellValue('A2', '月    報');
$sheet2->setCellValue('B2', '每月終了後10日內編報');
$sheet2->setCellValue('Q2', '表    號');
$sheet2->setCellValue('R2', '10956-00-01-2');
$sheet2->mergeCells('R1:V1');
$sheet2->mergeCells('R2:V2');
$sheet2->getStyle('A1:A2')->applyFromArray($thinBorder);
$sheet2->getStyle('Q1:Q2')->applyFromArray($thinBorder);
$sheet2->getStyle('R1:V1')->applyFromArray($thinBorder);
$sheet2->getStyle('R2:V2')->applyFromArray($thinBorder);
$sheet2->getStyle('A2:V2')->applyFromArray($thinBottomMedium);

// 標題
$sheet2->mergeCells('A3:V3');
$sheet2->setCellValue('A3', "{$organize}-舉發違反道路交通管理事件統計-移公路主管機關裁罰");
$sheet2->getStyle('A3')->applyFromArray($center);
$sheet2->getStyle('A3')->getFont()->setSize(16);
// 副標
$sheet2->mergeCells('A4:U4');
$sheet2->setCellValue('A4', "{$makeDate2}");
$sheet2->getStyle('A4')->applyFromArray($center);
$sheet2->getStyle('A4')->getFont()->setSize(12);
$sheet2->setCellValue('V4', '單位：件');
$sheet2->getStyle('A4:V4')->applyFromArray($thinBottomMedium);

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

    'C5' => '高速公路大重機違規',
    'C14' => '第92條',
    'D5' => '其他未列之項目',
    'D14' => '12~36條、38~62條、92條',
];
foreach ($headers as $cell => $text) {
    $sheet2->setCellValue($cell, $text);
    $sheet2->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet2->getStyle('A5:V15')->applyFromArray($center);
$sheet2->getStyle('A5:B15')->applyFromArray($thinOutlineBorderMedium);
foreach (range('C', 'V') as $col) {
    $sheet2->mergeCells("{$col}5:{$col}13");
    $sheet2->mergeCells("{$col}14:{$col}15");
    $sheet2->getStyle("{$col}5:{$col}13")->applyFromArray($thinOutlineBorder);
    $sheet2->getStyle("{$col}14:{$col}15")->applyFromArray($thinOutlineBorder);
}
$sheet2->getStyle('C14:D15')->applyFromArray($thinBottomMedium);

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
    $sheet2->setCellValue($cell, $text);
    $sheet2->getStyle($cell)->getAlignment()->setWrapText(true);
}
$sheet2->mergeCells('A16:B16');
$sheet2->mergeCells('A26:B26');
$sheet2->mergeCells('A27:B27');
$sheet2->mergeCells('A17:A19');
$sheet2->mergeCells('A20:A22');
$sheet2->mergeCells('A23:A25');
$sheet2->getStyle('A16:B27')->applyFromArray($center);
$sheet2->getStyle('A16:B16')->applyFromArray($thinBorder);
$sheet2->getStyle('B17:B25')->applyFromArray($thinBorder);
$sheet2->getStyle('A26:B26')->applyFromArray($thinBorder);
$sheet2->getStyle('A20:A22')->applyFromArray($thinOutlineBorder);
$sheet2->getStyle('A23:A25')->applyFromArray($thinOutlineBorder);
$sheet2->getStyle('C16:C27')->applyFromArray($thinBorder);
$sheet2->getStyle('A27:V27')->applyFromArray($thinBottomMedium);

// 內容資料
$data = [
    ["0", "0"],
    ["0", "0"],
    ["0", "0"],
    ["0", "0"],
    ["0", "0"],
    ["0", "0"],
    ["0", "0"],
    ["0", "0"],
    ["0", "0"],
    ["0", "0"],
    ["0", "0"],
    ["0", "0"],
];

$row = 16;
foreach ($data as $d) {
    $sheet2->setCellValue("C{$row}", $d[0]);
    $sheet2->setCellValue("D{$row}", $d[1]);

    $row++;
}
$sheet2->getStyle("C16:V" . ($row - 1))->applyFromArray($thinBorder);
$sheet2->getStyle('C27:V27')->applyFromArray($thinBottomMedium);
$sheet2->getStyle("C16:V" . ($row - 1))->getFont()->setSize(10);

$sheet2->setCellValue('A28', "填  表");
$sheet2->setCellValue('G28', "審  核");
$sheet2->setCellValue('K28', "業務主管人員");
$sheet2->setCellValue('K29', "主辦統計人員");
$sheet2->setCellValue('O28', "機關首長");
$sheet2->setCellValue('T28', "{$makeDate}編製");
$sheet2->setCellValue('A30', "資料來源：本局轄區內取締違反道路交通管理事件之舉發件數彙編。");
$sheet2->setCellValue('A31', "填表說明：本表編製1式3份，先送會計室(統計室)會核，並經機關首長核章後，1份送各直轄市、縣(市)政府主計處，1份送會計室(統計室) ，1份自存外，本表於規定期限內由網際網路線上傳送至內政部警政署警政統計資料庫。");

// 灰底
$sheet2->getStyle('C16:D16')->applyFromArray($grayFill);
$sheet2->getStyle('C17:D17')->applyFromArray($grayFill);
$sheet2->getStyle('C20:D20')->applyFromArray($grayFill);

// // 欄寬
$sheet2->getColumnDimension('A')->setWidth(12);
$sheet2->getColumnDimension('B')->setWidth(12);
$sheet2->getColumnDimension('C')->setWidth(13);
$sheet2->getColumnDimension('D')->setWidth(10);

// 檔名
$filename = "{$republicStartDate}-{$republicEndDate}-{$organize}-舉發違反道路交通管理事件統計-移公路主管機關裁罰.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

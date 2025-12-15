<?php
require '../../vendor/autoload.php';
require '../styles.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("獎勵金點數核對表");

// 分局變數
$organize = "";
switch ($_GET['organize']) {
    case 'one':
        $organize = "第一分局";
        break;
    case 'two':
        $organize = "第二分局";
        break;
    case 'three':
        $organize = "第三分局";
        break;
    case 'traffic':
        $organize = "交通隊";
        break;
    case 'garrison':
        $organize = "警備隊";
        break;
}

// 日期變數
$makeDate  = "11407";
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(12);

// 標題
$sheet->mergeCells('A1:O1');
$sheet->setCellValue('A1', "{$makeDate}新竹市警察局{$organize}員警獎勵金點數核對清冊表");
$sheet->getStyle('A1')->applyFromArray($center);
$sheet->getStyle('A1')->getFont()->setSize(16);

// 表頭
$headers = [
    'A2' => '單位',
    'B2' => '職別',
    'C2' => '姓名',
    'D2' => '取締點數',
    'E2' => '交整時',
    'F2' => '交整點數',
    'G2' => '機車拖吊',
    'H2' => '小型車拖吊',
    'I2' => '大型車拖吊',
    'J2' => '拖吊點數',
    'K2' => 'A1件數',
    'L2' => 'A2件數',
    'M2' => 'A3件數',
    'N2' => '事故處理點',
    'O2' => '合計點數',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}

// 內容資料
$data = [
    ["新竹市警察局第一分局", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
];

$row = 3;
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

    $row++;
}
$sheet->getStyle("A2:O" . ($row - 1))->applyFromArray($thinBorderMedium);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(20);
$sheet->getColumnDimension('N')->setWidth(14);

// 檔名
$filename = "{$makeDate}新竹市警察局{$organize}員警獎勵金點數核對清冊表.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

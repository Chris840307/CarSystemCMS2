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
$sheet->setTitle("印領清冊");

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
    case 'safety':
        $organize = "保安隊";
        break;
}

// 日期變數
$makeDate  = "114年7月";
$republicStartDate = '11407';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(12);

// 標題
$sheet->mergeCells('A1:I1');
$sheet->setCellValue('A1', "新竹市警察局{$organize}");
$sheet->getStyle('A1')->applyFromArray($center);
$sheet->getStyle('A1')->getFont()->setSize(16);

$sheet->mergeCells('A2:I2');
$sheet->setCellValue('A2', "{$makeDate}「直接處理道路交通安全任務人員獎勵金」印領清冊");
$sheet->getStyle('A2')->applyFromArray($center);
$sheet->getStyle('A2')->getFont()->setSize(16);

// 表頭
$headers = [
    'A3' => '編號',
    'B3' => '單位',
    'C3' => '職別',
    'D3' => '姓名',
    'E3' => '配發點數',
    'F3' => '每筆點數金額',
    'G3' => '應領獎勵金額',
    'H3' => '實領獎勵金額',
    'I3' => '備考',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}
$sheet->getStyle('A3:I3')->getAlignment()->setWrapText(true);

// 內容資料
$data = [
    ["保安隊", "", "", "", "", "", "", "", "0"],
];

$seq = 1; //編號
$row = 4;
foreach ($data as $d) {
    $sheet->setCellValue("A{$row}", $seq++);  //編號
    $sheet->setCellValue("B{$row}", $d[0]);
    $sheet->setCellValue("C{$row}", $d[1]);
    $sheet->setCellValue("D{$row}", $d[2]);
    $sheet->setCellValue("E{$row}", $d[3]);
    $sheet->setCellValue("F{$row}", $d[4]);
    $sheet->setCellValue("G{$row}", $d[5]);
    $sheet->setCellValue("H{$row}", $d[6]);
    $sheet->setCellValue("I{$row}", $d[7]);

    $row++;
}

$data = [
    ["小計", "", "", "", "", "", "", ""],
];

$row;
foreach ($data as $d) {
    $sheet->setCellValue("B{$row}", $d[0]);
    $sheet->setCellValue("C{$row}", $d[1]);
    $sheet->setCellValue("D{$row}", $d[2]);
    $sheet->setCellValue("E{$row}", $d[3]);
    $sheet->setCellValue("F{$row}", $d[4]);
    $sheet->setCellValue("G{$row}", $d[5]);
    $sheet->setCellValue("H{$row}", $d[6]);
    $sheet->setCellValue("I{$row}", $d[7]);

    $row++;
}
$sheet->getStyle("A1:I" . ($row - 1))->applyFromArray($center);

// 全表格框線
$sheet->getStyle("A1:I" . ($row - 1))->applyFromArray($thinBorder);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(12);
$sheet->getColumnDimension('B')->setWidth(11);
$sheet->getColumnDimension('C')->setWidth(10);
$sheet->getColumnDimension('D')->setWidth(10);
$sheet->getColumnDimension('E')->setWidth(10);
$sheet->getColumnDimension('F')->setWidth(12);
$sheet->getColumnDimension('G')->setWidth(13);
$sheet->getColumnDimension('H')->setWidth(10);
$sheet->getColumnDimension('I')->setWidth(12);

// 檔名
$filename = "{$republicStartDate}外勤員警獎勵金印領清冊.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

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
$sheet->setTitle("未達清冊");

// 日期變數
$republicStartDate = '1141118';
$republicEndDate   = '1141119';

// 全局字體：標楷體
$spreadsheet->getDefaultStyle()->getFont()->setName('標楷體')->setSize(12);

// 表頭
$headers = [
    'A1' => '序',
    'B1' => '單號',
    'C1' => '掛號號碼',
    'D1' => '付郵日',
    'E1' => '收件人',
    'F1' => '地址',
    'G1' => '舉發單位',
];
foreach ($headers as $cell => $text) {
    $sheet->setCellValue($cell, $text);
}

// 內容資料
$data = [
    ["E44A09461", "564445-30321417", "1140912", "馬可威", "100台北市中正區水源路４７號十三樓之１", "第二分局文華派出所"],
];

$seq = 1; //序
$row = 2;
foreach ($data as $d) {
    $sheet->setCellValue("A{$row}", $seq++);  //序
    $sheet->setCellValue("B{$row}", $d[0]);
    $sheet->setCellValue("C{$row}", $d[1]);
    $sheet->setCellValue("D{$row}", $d[2]);
    $sheet->setCellValue("E{$row}", $d[3]);
    $sheet->setCellValue("F{$row}", $d[4]);
    $sheet->setCellValue("G{$row}", $d[5]);

    $row++;
}

// 全表格框線
$sheet->getStyle("A1:G" . ($row - 1))->applyFromArray($thinBorder);

// 欄寬
$sheet->getColumnDimension('A')->setWidth(6);
$sheet->getColumnDimension('B')->setWidth(14);
$sheet->getColumnDimension('C')->setWidth(18);
$sheet->getColumnDimension('D')->setWidth(9);
$sheet->getColumnDimension('E')->setWidth(34);
$sheet->getColumnDimension('F')->setWidth(80);
$sheet->getColumnDimension('G')->setWidth(24);

// 檔名
$filename = "{$republicStartDate}-{$republicEndDate}-未達清冊.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

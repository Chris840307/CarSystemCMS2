<?php
require_once "../db.php";
require '../../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

$queryDate = $_GET['query_date'] ?? date('Y-m-d');

$sql = "SELECT *
        FROM search_log
        WHERE DATE(cdate) = :query_date
        ORDER BY id ASC";
$stmt = $pdo->prepare($sql);
$stmt->execute([
    ':query_date' => $queryDate
]);
$rows = $stmt->fetchAll(PDO::FETCH_ASSOC);

$data = [];
foreach ($rows as $row) {
    $detail = json_decode($row['detail'], true);
    $search_target = isset($detail['searchr']) ? $detail['searchr'] : '未知';

    $data[] = [
        $row['id'],
        $row['username'],
        $row['ip'],
        $row['useragent'],
        $row['cdate'],
        $search_target,
        $row['searchr'],
    ];
}

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

$sheet->mergeCells('A1:G1');
$sheet->setCellValue('A1', '案件查詢系統');
$sheet->getStyle('A1')->applyFromArray([
    'font' => ['size' => 20],
    'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER],
]);

$sheet->mergeCells('A2:G2');
$sheet->setCellValue('A2', '稽核日期：' . date('Y-m-d'));
$sheet->getStyle('A2')->applyFromArray([
    'font' => ['size' => 10],
    'alignment' => ['horizontal' => Alignment::HORIZONTAL_RIGHT],
]);

$headers = ['序號', '使用者名稱', 'IP', '瀏覽器版本', '時間', '查詢標的', '查詢事由'];
$sheet->fromArray($headers, NULL, 'A3');
$sheet->getStyle('A3:G3')->applyFromArray([
    'alignment' => ['horizontal' => Alignment::HORIZONTAL_LEFT],
    'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
]);

$startRow = 4;
foreach ($data as $i => $row) {
    $rowIndex = $startRow + $i;
    $sheet->fromArray($row, NULL, 'A' . $rowIndex);
    $sheet->getStyle('A' . ($startRow + $i))->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
    $sheet->getStyle('A' . ($startRow + $i) . ':G' . ($startRow + $i))->applyFromArray([
        'alignment' => ['vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true],
    ]);
}

$colWidths = [5, 14, 15, 50, 20, 50, 15];
foreach (range('A', 'G') as $i => $col) {
    $sheet->getColumnDimension($col)->setWidth($colWidths[$i]);
}

$filename = 'search_log_' . date('Ymd_His') . '.xlsx';
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment;filename=\"$filename\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

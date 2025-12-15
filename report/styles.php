<?php

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Borders;
use PhpOffice\PhpSpreadsheet\Style\Fill;

// 字體樣式
$center = [
    'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER],
    'font' => ['name' => '標楷體']
];
$left = [
    'alignment' => ['horizontal' => Alignment::HORIZONTAL_LEFT, 'vertical' => Alignment::VERTICAL_CENTER],
    'font' => ['name' => '標楷體']
];
$right = [
    'alignment' => ['horizontal' => Alignment::HORIZONTAL_RIGHT, 'vertical' => Alignment::VERTICAL_CENTER],
    'font' => ['name' => '標楷體']
];
$top = [
    'alignment' => ['horizontal' => Alignment::HORIZONTAL_LEFT, 'vertical' => Alignment::VERTICAL_TOP],
    'font' => ['name' => '標楷體']
];
$centerTop = [
    'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_TOP],
    'font' => ['name' => '標楷體']
];
$newMingCenter = [
    'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER],
    'font' => ['name' => '新細明體']
];
$newMingRight = [
    'alignment' => ['horizontal' => Alignment::HORIZONTAL_RIGHT, 'vertical' => Alignment::VERTICAL_CENTER],
    'font' => ['name' => '新細明體']
];

// 邊框
$thinBorder = ['borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]]];
$thinBorderMedium = ['borders' => ['allBorders' => ['borderStyle' => Border::BORDER_MEDIUM]]];
$thinOutlineBorder = ['borders' => ['outline' => ['borderStyle' => Border::BORDER_THIN]]];
$thinOutlineBorderMedium = ['borders' => ['outline' => ['borderStyle' => Border::BORDER_MEDIUM]]];

// 底線
$thinBottom = ['borders' => ['bottom' => ['borderStyle' => Border::BORDER_THIN]]];
$thinBottomMedium = ['borders' => ['bottom' => ['borderStyle' => Border::BORDER_MEDIUM]]];

// 斜線
$diagonal = ['borders' => ['diagonal' => ['borderStyle' => Border::BORDER_THIN, 'direction' => Borders::DIAGONAL_DOWN]]];

// 橘底
$orangeFill = ['fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => 'FFCC00']]];
$orangeFill2 = ['fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => 'FFCC99']]];
// 綠底
$greenFill = ['fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => '99CC00']]];
$greenFill2 = ['fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => 'CCFFCC']]];
// 黃底
$yellowFill = ['fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => 'FFFFCC']]];
$yellowFill2 = ['fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => 'FFFF99']]];
$yellowFill3 = ['fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => 'FFFF00']]];
// 籃底
$blueFill = ['fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => 'CCFFFF']]];
// 灰底
$grayFill = ['fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => 'C0C0C0']]];
// 粉紅底
$pinkFill = ['fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => 'FF99CC']]];

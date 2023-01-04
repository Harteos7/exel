<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//load spreadsheet
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("hello world.xlsx");
$sheet = $spreadsheet->getActiveSheet();

//change it
$sheet->setCellValue('A1', 'Noms des applications');
$sheet->setCellValue('A2', 'test');
$sheet->setCellValue('A3', 'rien');
$sheet->setCellValue('A4', 'cest pas A3');

//take it
$b = 'A1';
$a1 = $sheet->getCell($b);
$a2 = $sheet->getCell('A2');
$a3 = $sheet->getCell('A3');
$a4 = $sheet->getCell('A4');
$a5 = $sheet->getCell('A5');

//write it again to Filesystem with the same name (=replace)
$writer = new Xlsx($spreadsheet);
$writer->save('hello world.xlsx');

$arr = array(
    'A1' => strval($a1), 
    'A2' => strval($a2), 
    'A3' => strval($a3), 
    'A4' => strval($a4),
    'A5' => strval($a5)
);
print_r ($arr);
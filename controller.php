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
$a2 = $sheet->getCell('A2');
$a3 = $sheet->getCell('A3');
$a4 = $sheet->getCell('A4');

//write it again to Filesystem with the same name (=replace)
$writer = new Xlsx($spreadsheet);
$writer->save('hello world.xlsx');

echo "A2 is $a2 A3 is $a3 A4 is $a4";



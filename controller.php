<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//load spreadsheet
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("hello world.xlsx");
$sheet = $spreadsheet->getActiveSheet();

//it's the code for writing data in file.xlsx
//change it
$sheet->setCellValue('A1', 'Noms des applications');
$sheet->setCellValue('A2', 'test');
$sheet->setCellValue('A3', 'rien');
$sheet->setCellValue('A4', 'cest pas A3');
$sheet->setCellValue('B1', 'Prix');
$sheet->setCellValue('B2', '19745');
$sheet->setCellValue('B3', '1475621');
$sheet->setCellValue('B4', '155');
$sheet->setCellValue('B5', '834548');
$sheet->setCellValue('C1', 'Lieu');
$sheet->setCellValue('C2', 'Lion');
$sheet->setCellValue('C3', 'Lile');

//write it again to Filesystem with the same name (=replace)
$writer = new Xlsx($spreadsheet);
$writer->save('hello world.xlsx');

// it's the code for read data in file.xlsx
$arr = array();
$c = 'A';
for ($i = 1; ; $i++) { // $i and $c are the coordinates, respectively the number is the letter (A1, A2, B2, C3, ...)

    $b = strval($c).strval($i); // $b is the coordinates of the exel box
    $a = $sheet->getCell($b); // $a is the content of the exel box
    if ($a == '') { // we check that the new cell has data otherwise we change the column
        $c++;
        $i = 1;
        $b = strval($c).strval($i);
        $a = $sheet->getCell($b);

            if ($a == '') {break;} // we check if the first cell of the new column has data if not, we stop the for (break)
            else
            $arr[strval($b)] = strval($a);
        
    } 
    else
    $arr[strval($b)] = strval($a); // we put everything in array (the key is the coordinates and the value of their data)
}
print_r ($arr);

<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


// it's the code for read data in file.xlsx
function read(string $exel)
{

    //load spreadsheet
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load(strval($exel));
    $sheet = $spreadsheet->getActiveSheet();

    $arr = array();
    $letter = 'A';
    for ($number = 1; ; $number++) { // $number and $letter are the coordinates (A1, A2, B2, C3, ...)

        $id = strval($letter) . strval($number); // $id is the coordinates of the exel box
        $cell = $sheet->getCell($id); // $cell is the content of the exel box
        if ($cell == '') { // we check that the new cell has data otherwise we change the column
            $letter++;
            $number = 1;
            $id = strval($letter) . strval($number);
            $cell = $sheet->getCell($id);

            if ($cell == '') {
                break;
            } // we check if the first cell of the new column has data if not, we stop the for (break)
            else
                $arr[strval($id)] = strval($cell);

        } else
            $arr[strval($id)] = strval($cell); // we put everything in array (the key is the coordinates and the value of their data)
    }
    print_r($arr);
    return $arr;
}

//it's the code for writing data in file.xlsx
function writenew(string $exel)
{

    //load spreadsheet
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load(strval($exel));
    $sheet = $spreadsheet->getActiveSheet();

    $sheet->setCellValue('A1', 'Noms des applications');
    $sheet->setCellValue('A2', 'pomme and fraise');
    $sheet->setCellValue('A3', 'macdo drive online');
    $sheet->setCellValue('A4', 'wish');



    //write it again to Filesystem with the same name (=replace)
    $writer = new Xlsx($spreadsheet);
    $writer->save($exel);

}

function sortbyletter(string $letter,array $arr)
{

    foreach ($arr as $cle => $valeur) {
        if ($letter == $cle[0]) {
            echo 'La clé ' . $cle . ' contient la valeur ' . $valeur . "\n";
        }
    }


}

function sortbynumber(int $number,array $arr)
{

    foreach ($arr as $cle => $valeur) {
        if ($number == $cle[1]) {
            echo 'La clé ' . $cle . ' contient la valeur ' . $valeur . "\n";
        }
    }

}

$arr=read('hello world.xlsx');

writenew('hello world.xlsx');

sortbyletter('A',$arr);

sortbynumber('1',$arr);


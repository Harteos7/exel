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

    $letterM='';  //$letterM is the largest letter used in the coordinates of the exel table
    $numberM=0;   //$numberM is the largest number used in the coordinates of the exel array

    foreach ($arr as $cle => $valeur) {
        $letterM = strval($letterM) . strval($cle[0]);
    }
    $letterM = substr(strval($letterM), -1); //gets the largest letter

    foreach ($arr as $cle => $valeur) {
        if($numberM<strval($cle[1])){
            $numberM = strval($cle[1]); //get the largest number of coordinates
        }
    }
    
    $letter = 'A';
    for ($number = 1; ;) { // $number and $letter are the coordinates (A1, A2, B2, C3, ...)

        $id = strval($letter) . strval($number); // $id is the coordinates of the exel box
        $cell = $sheet->getCell($id); // $cell is the content of the exel box
        echo $id;
        if ($cell == '' ) // this script saves the cell if it is filled
            {} else {
                $cell = $sheet->getCell($id);
                $arr[strval($id)] = strval($cell);
            }

        if ($number == $numberM) { // either we change column if we are on the last line or we change line
            $letter++;
            $number = 1;
            $id = strval($letter) . strval($number);
        } else {$number++;}

        if ($letter == $letterM && $number == $numberM) { // if we are on the last cell
            $id = strval($letter) . strval($number); // $id is the coordinates of the exel box
            $cell = $sheet->getCell($id); // $cell is the content of the exel box
            if ($cell == '' ) // this script saves the cell if it is filled
            {break;}
            $cell = $sheet->getCell($id);
            $arr[strval($id)] = strval($cell);
            break; // end

        }
          
    }
    print_r($arr);


    return $arr;
}

//it's the code for writing data in file.xlsx
    function writenew(string $exel,array $arr)
    {

        //load spreadsheet
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load(strval($exel));
        $sheet = $spreadsheet->getActiveSheet();

        $letterM='';  //$letterM is the largest letter used in the coordinates of the exel table
        $numberM=0;   //$numberM is the largest number used in the coordinates of the exel array
    
        foreach ($arr as $cle => $valeur) {
            $letterM = strval($letterM) . strval($cle[0]);
        }
        $letterM = substr(strval($letterM), -1); //gets the largest letter
    
        foreach ($arr as $cle => $valeur) {
            if($numberM<strval($cle[1])){
                $numberM = strval($cle[1]); // get the largest number of coordinates
            }
        }

        $numberM = $numberM+1;

        $nom = readline("Nom: ");

        $prix = readline("Prix : ");

        $lieu = readline("Lieu : ");

        $sheet->setCellValue('A'.strval($numberM), strval($nom));
        $sheet->setCellValue('B'.strval($numberM), strval($prix));
        $sheet->setCellValue('C'.strval($numberM), strval($lieu));

        //write it again to Filesystem with the same name (=replace)
        $writer = new Xlsx($spreadsheet);
        $writer->save($exel);

    }

//it's the code for sort by column
function sortbyletter(string $letter,array $arr)
{

    foreach ($arr as $cle => $valeur) {
        if ($letter == $cle[0]) {
            echo 'La clé ' . $cle . ' contient la valeur ' . $valeur . "\n";
        }
    }

}

//it's the code for sort by line
function sortbynumber(int $number,array $arr)
{

    foreach ($arr as $cle => $valeur) {
        if ($number == $cle[1]) {
            echo 'La clé ' . $cle . ' contient la valeur ' . $valeur . "\n";
        }
    }

}

//it is a function that fills the empty cases as long as there is a data on the line (to be able to display an application well)
//be careful : this function works only if a column of the exel file is completely filled !
function debugexel(array $arr,string $exel)
{
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load(strval($exel));
    $sheet = $spreadsheet->getActiveSheet();
    $letterM='';  //$letterM is the largest letter used in the coordinates of the exel table
    $numberM=0;   //$numberM is the largest number used in the coordinates of the exel array

    foreach ($arr as $cle => $valeur) {
        $letterM = strval($letterM) . strval($cle[0]);
    }
    $letterM = substr(strval($letterM), -1); //gets the largest letter

    foreach ($arr as $cle => $valeur) {
        if($numberM<strval($cle[1])){
            $numberM = strval($cle[1]); // get the largest number of coordinates
        }
    }

    for ($letter='A',$number='1'; ;){
        $cell = strval($letter) . strval($number);
        if (array_key_exists(strval($cell), $arr)) {}//We look if the coordinate exists in $arr so if it has data, if it is the case we do nothing
            else{ // otherwise we create a data for this cell named "nothing"
                $sheet->setCellValue(strval($cell), 'nothing');
                $writer = new Xlsx($spreadsheet);
                $writer->save($exel);
            }
        if ($number<$numberM)
        {
            $number++;
        }else
        {
            $letter++;
            $number=1;
        }
        if ($letter > $letterM) {break;}
    }


}

$arr=read('hello world.xlsx');

writenew('hello world.xlsx',$arr);

// sortbyletter('A',$arr);

// sortbynumber('1',$arr);

// debugexel($arr,'hello world.xlsx');

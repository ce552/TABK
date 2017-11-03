<?php

//PHPExcel is used to convert Excel workbooks into CSV for easier database input.
require_once 'C:/xampp/php/pear/PHPExcel-1.8/Classes/PHPExcel.php';
require_once 'C:/xampp/php/pear/PHPExcel-1.8/Classes/PHPExcel/IOFactory.php';

//Allowed file types.
$allowed_types = ['application/vnd.ms-excel','text/xls','text/xlsx',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];

$timeStamp = date('Ymd');
//Executes once upload form is submitted. File name is sent with the name 'file'

if(isset($_POST)) {
    if(in_array($_FILES['file']['type'], $allowed_types)){

        //Creates a new file to hold upload if doesnt exist. The folder name is upload date in format YYYMMDD.
        if (!file_exists('C:/xampp/htdocs/TABK/uploads/'.$timeStamp)) {
            mkdir('C:/xampp/htdocs/TABK/uploads/' . $timeStamp);
        }

        $uploadDir = 'C:/xampp/htdocs/TABK/uploads/'.$timeStamp.'/';
        $temp = explode(".", $_FILES["file"]["name"]);
        $uploadName = "input" . '.' . end($temp);   //Renames file to 'input'.(extension)

        //Permanently move file to uploads directory corresponding to upload date.
        move_uploaded_file($_FILES["file"]["tmp_name"], $uploadDir. '/' . $uploadName);
        echo "File is valid, and was successfully uploaded." . '<br>';
    }
    else {
        echo "Error: Only excel sheets allowed!" . '<br>';
        die();
    }
}

//Creating PHPExcel Object to access the functions.
$objPHPExcel = new PHPExcel();
$fileName = 'C:/xampp/htdocs/TABK/uploads/'.$timeStamp.'/input.xlsx';
try {
    //Locate file and load it into PHPExcel library for manipulation.
    $inputFileType = PHPExcel_IOFactory::identify($fileName);
    $objReader = PHPExcel_IOFactory::createReader($inputFileType);
    $objPHPExcel = $objReader->load($fileName);
    if (!file_exists('C:/xampp/htdocs/TABK/uploads/CSV/'.$timeStamp)) {
        mkdir('C:/xampp/htdocs/TABK/uploads/CSV/' . $timeStamp);
    }
}
catch (Exception $e) {
    die('Error loading file "'.pathinfo($fileName,PATHINFO_BASENAME).'": '.$e->getMessage());
}

//Convert from excel document to CSV. A folder with todays date is created and converted sheets
//are placed inside them.
$sheetCount = $objPHPExcel->getSheetCount();
echo 'There ',(($sheetCount == 1) ? 'is' : 'are'),' ',$sheetCount,' WorkSheet',(($sheetCount == 1) ? '' : 's'),' in the WorkBook<br /><br />';
$objPHPExcel->setActiveSheetIndex(0);
$csvDir = "C:/xampp/htdocs/TABK/uploads/CSV/".$timeStamp."/";

for ($i = 0; $i < $sheetCount; $i++) {
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'CSV');
    $objWriter->setSheetIndex($i);   // Select which sheet.
    $objWriter->setDelimiter(';');  // Define delimiter
    $objWriter->save( $csvDir . "sheet#" . $i . ".csv");
    echo "Sheet " . $i . " done" . "<br>";
}

?>

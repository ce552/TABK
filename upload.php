<?php

require_once 'C:/xampp/php/pear/PHPExcel-1.8/Classes/PHPExcel.php';
require_once 'C:/xampp/php/pear/PHPExcel-1.8/Classes/PHPExcel/IOFactory.php';

$allowed_types = ['application/vnd.ms-excel','text/xls','text/xlsx',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];

if(isset($_POST)) {
    if(in_array($_FILES['file']['type'], $allowed_types)){
        $uploaddir = 'C:/xampp/htdocs/TABK/uploads/'; //Change to your permanent directory
        $uploadfile = $uploaddir . basename($_FILES['file']['name']);
        $temp = explode(".", $_FILES["file"]["name"]);
        $newfilename = "input" . '.' . end($temp);

        move_uploaded_file($_FILES["file"]["tmp_name"], "C:/xampp/htdocs/TABK/uploads/" . $newfilename);
        echo "File is valid, and was successfully uploaded.\n";
    } else {
        echo "Error: Only excel sheets allowed!";
        die();
    }
}
$FileName = 'C:/xampp/htdocs/TABK/uploads/input.xlsx';
convertXLStoCSV($FileName);

function convertXLStoCSV ($inputFileName)
{
    $objPHPExcel = new PHPExcel();

    try {
        $inputFileType = PHPExcel_IOFactory::identify($inputFileName);
        $objReader = PHPExcel_IOFactory::createReader($inputFileType);
        $objPHPExcel = $objReader->load($inputFileName);
        /**  Read the Date when the workbook was created (as a PHP timestamp value)  **/
        $creationDatestamp = $objPHPExcel->getProperties()->getCreated();
        /**  Format the date and time using the standard PHP date() function  **/
        $createName = "C:/xampp/htdocs/TABK/uploads/CSV/" . date("m.d.y", $creationDatestamp);
        $creationDate = date('l, d<\s\up>S</\s\up> F Y',$creationDatestamp);
        $creationTime = date('g:i A',$creationDatestamp);
        echo '<b>Created On: </b>',$creationDate,' at ',$creationTime,'<br />';
    }
    catch (Exception $e) {
        die('Error loading file "'.pathinfo($inputFileName,PATHINFO_BASENAME).'": '.$e->getMessage());
    }

    $sheetCount = $objPHPExcel->getSheetCount();
    echo 'There ',(($sheetCount == 1) ? 'is' : 'are'),' ',$sheetCount,' WorkSheet',(($sheetCount == 1) ? '' : 's'),' in the WorkBook<br /><br />';
    $objPHPExcel->setActiveSheetIndex(0);
    // Export to CSV file.
    for ($i = 0; $i < $sheetCount; $i++) {
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'CSV');
        $objWriter->setSheetIndex($i);   // Select which sheet.
        $objWriter->setDelimiter(';');  // Define delimiter
        $objWriter->save($createName . $i . ".csv");
        echo "done sheet " . $i . "<br>";
    }
}
?>

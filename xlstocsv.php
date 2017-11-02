<?php


// Error reporting
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

// Path to PHPExcel classes
require_once '/Applications/XAMPP/xamppfiles/pear/PHPExcel-1.8/Classes/PHPExcel.php';
require_once '/Applications/XAMPP/xamppfiles/pear/PHPExcel-1.8/Classes/PHPExcel/IOFactory.php';

// Your input Excel file.
$excelFile = '/Applications/XAMPP/xamppfiles';

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

//  Read your Excel workbook
try
{
    $inputFileType = PHPExcel_IOFactory::identify($excelFile);
    $objReader = PHPExcel_IOFactory::createReader($inputFileType);
    $objPHPExcel = $objReader->load($excelFile);
}
catch(Exception $e)
{
    die('Error loading file "'.pathinfo($excelFile,PATHINFO_BASENAME).'": '.$e->getMessage());
}

// Export to CSV file.
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'CSV');
$objWriter->setSheetIndex(0);   // Select which sheet.
$objWriter->setDelimiter(';');  // Define delimiter
$objWriter->save('testExportFile.csv');

echo "done";
?>

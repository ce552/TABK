<?php

$allowed_types = ['application/vnd.ms-excel','text/xls','text/xlsx',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];

if(isset($_POST)) {
    if(in_array($_FILES['file']['type'], $allowed_types)){
    $uploaddir = 'C:/xampp/htdocs/TABK/uploads/'; //Change to your permanent directory
    $uploadfile = $uploaddir . basename($_FILES['file']['name']);

    move_uploaded_file($_FILES['file']['tmp_name'], $uploadfile);
    echo "File is valid, and was successfully uploaded.\n";
    } else {
        echo "Error: Only excel sheets allowed!";
        die();
    }
}
?>

<?php

require __DIR__ . '/vendor/autoload.php';

use \PhpOffice\PhpSpreadsheet\Reader\Xls;
use \PhpOffice\PhpSpreadsheet\Writer\Csv;

// Check if the form was submitted
if($_SERVER["REQUEST_METHOD"] == "POST"){
    
    if(isset($_FILES["xlsfile"]) && $_FILES["xlsfile"]["error"] == 0){
        $allowed = array("xls" => "application/vnd.ms-excel");
        $filename = $_FILES["xlsfile"]["name"];
        $filetype = $_FILES["xlsfile"]["type"];
        $filesize = $_FILES["xlsfile"]["size"];
        
        // Verify file extension
        $ext = pathinfo($filename, PATHINFO_EXTENSION);
        if(!array_key_exists($ext, $allowed)) die("Error: Please select a valid file format.");
        $filename_new = rand(2, 4).time().rand(2, 4);
        $filename_new_xls = $filename_new.'.'.$ext;
        // Verify file size - 5MB maximum
        $maxsize = 5 * 1024 * 1024;
        if($filesize > $maxsize) die("Error: File size is larger than the allowed limit.");
        
        // Verify MYME type of the file
        if(in_array($filetype, $allowed)){
            // Check whether file exists before uploading it
            
            move_uploaded_file($_FILES["xlsfile"]["tmp_name"], __DIR__ . "/xls-downloads/" . $filename_new_xls);
            
           
        } else{
            echo "Error: There was a problem uploading your file. Please try again.";  die;
        }
    } else{
        echo "Error: " . $_FILES["xlsfile"]["error"]; die;
    }
}

$xls_file = __DIR__ . "/xls-downloads/" . $filename_new_xls;

$reader = new Xls();
$spreadsheet = $reader->load($xls_file);

$loadedSheetNames = $spreadsheet->getSheetNames();

$writer = new Csv($spreadsheet);
$writer->setDelimiter(';');

foreach($loadedSheetNames as $sheetIndex => $loadedSheetName) {
    $writer->setSheetIndex($sheetIndex);
    $writer->save(__DIR__ . "/csv-downloads/".$filename_new.".csv");
}

header("Location: ".$_SERVER['HTTP_HOST']."/csv-downloads/".$filename_new.".csv");


?>

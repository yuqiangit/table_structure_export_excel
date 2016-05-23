<?php
/**
 * Created by PhpStorm.
 * User: shihuipeng
 * Date: 16/5/18
 * Time: 下午6:22
 */


/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

/** Include PHPExcel */
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';


// Create new PHPExcel object
echo date('H:i:s') , " Create new PHPExcel object" , EOL;
$objPHPExcel = new PHPExcel();

// Set document properties
echo date('H:i:s') , " Set document properties" , EOL;
$objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
    ->setLastModifiedBy("Maarten Balliauw")
    ->setTitle("Office 2007 XLSX Test Document")
    ->setSubject("Office 2007 XLSX Test Document")
    ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
    ->setKeywords("office 2007 openxml php")
    ->setCategory("Test result file");

// Set default font
//echo date('H:i:s') , " Set default font" , EOL;
$objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
    ->setSize(10);

// Add some data, resembling some different data types
//echo date('H:i:s') , " Add some data" , EOL;
//$objPHPExcel->getActiveSheet()->setCellValue('A1', 'String')
//    ->setCellValue('B1', 'Simple')
//    ->setCellValue('C1', 'PHPExcel');
//
//$objPHPExcel->getActiveSheet()->setCellValue('A2', 'String')
//    ->setCellValue('B2', 'Symbols')
//    ->setCellValue('C2', '!+&=()~§±æþ');
//
//$objPHPExcel->getActiveSheet()->setCellValue('A3', 'String')
//    ->setCellValue('B3', 'UTF-8')
//    ->setCellValue('C3', 'Создать MS Excel Книги из PHP скриптов');
//
//$objPHPExcel->getActiveSheet()->setCellValue('A4', 'Number')
//    ->setCellValue('B4', 'Integer')
//    ->setCellValue('C4', 12);
//
//$objPHPExcel->getActiveSheet()->setCellValue('A5', 'Number')
//    ->setCellValue('B5', 'Float')
//    ->setCellValue('C5', 34.56);
//
//$objPHPExcel->getActiveSheet()->setCellValue('A6', 'Number')
//    ->setCellValue('B6', 'Negative')
//    ->setCellValue('C6', -7.89);
//
//$objPHPExcel->getActiveSheet()->setCellValue('A7', 'Boolean')
//    ->setCellValue('B7', 'True')
//    ->setCellValue('C7', true);
//
//$objPHPExcel->getActiveSheet()->setCellValue('A8', 'Boolean')
//    ->setCellValue('B8', 'False')
//    ->setCellValue('C8', false);

//$dateTimeNow = time();
//$objPHPExcel->getActiveSheet()->setCellValue('A9', 'Date/Time')
//    ->setCellValue('B9', 'Date')
//    ->setCellValue('C9', PHPExcel_Shared_Date::PHPToExcel( $dateTimeNow ));
//$objPHPExcel->getActiveSheet()->getStyle('C9')->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDD2);
//
//$objPHPExcel->getActiveSheet()->setCellValue('A10', 'Date/Time')
//    ->setCellValue('B10', 'Time')
//    ->setCellValue('C10', PHPExcel_Shared_Date::PHPToExcel( $dateTimeNow ));
//$objPHPExcel->getActiveSheet()->getStyle('C10')->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_TIME4);
//
//$objPHPExcel->getActiveSheet()->setCellValue('A11', 'Date/Time')
//    ->setCellValue('B11', 'Date and Time')
//    ->setCellValue('C11', PHPExcel_Shared_Date::PHPToExcel( $dateTimeNow ));
//$objPHPExcel->getActiveSheet()->getStyle('C11')->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
//
//$objPHPExcel->getActiveSheet()->setCellValue('A12', 'NULL')
//    ->setCellValue('C12', NULL);

$link = mysqli_connect('127.0.0.1', 'root', '123456', 'LikingFit');
if (mysqli_connect_errno()) {
    exit('1');
}

mysqli_query($link, 'set names utf8');
$tables = mysqli_query($link, 'show tables');
$tableName = mysqli_fetch_all($tables);

$i = 1;
foreach($tableName as $v){
    $i ++;
    $sql = sprintf("select * from TABLES where TABLE_SCHEMA='LikingFit' and TABLE_NAME= '%s'",$v[0]);
    mysqli_query($link,'use information_schema');
    $tableInfo = mysqli_query($link,$sql);
    $tableInfo = mysqli_fetch_all($tableInfo);
    $objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue('A'.$i, $v[0].'--'.$tableInfo[0][20]);

    mysqli_query($link,'use LikingFit');
    $tableFileInfo = mysqli_query($link,sprintf('show full fields from `%s`',$v[0]));
    $res = mysqli_fetch_all($tableFileInfo);
    foreach($res as $ve){
        $i ++;
        //0字段名 1类型 0备注
        // Add some data
        $objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('B'.$i, $ve[0])
            ->setCellValue('C'.$i, $ve[1])
            ->setCellValue('D'.$i, $ve[8]);

    }
}


mysqli_close($link);

//$objRichText = new PHPExcel_RichText();
//$objRichText->createText('你好 ');

//$objPayable = $objRichText->createTextRun('你 好 吗？');
//$objPayable->getFont()->setBold(true);
//$objPayable->getFont()->setItalic(true);
//$objPayable->getFont()->setColor( new PHPExcel_Style_Color( PHPExcel_Style_Color::COLOR_DARKGREEN ) );

//$objRichText->createText(', unless specified otherwise on the invoice.');

//$objPHPExcel->getActiveSheet()->setCellValue('A13', 'Rich Text')
//    ->setCellValue('C13', $objRichText);


$objRichText2 = new PHPExcel_RichText();
$objRichText2->createText("black text\n");

$objRed = $objRichText2->createTextRun("red text");
$objRed->getFont()->setColor( new PHPExcel_Style_Color(PHPExcel_Style_Color::COLOR_RED  ) );

$objPHPExcel->getActiveSheet()->getCell("C14")->setValue($objRichText2);
$objPHPExcel->getActiveSheet()->getStyle("C14")->getAlignment()->setWrapText(true);


$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);

// Rename worksheet
echo date('H:i:s') , " Rename worksheet" , EOL;
$objPHPExcel->getActiveSheet()->setTitle('Datatypes');


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);


// Save Excel 2007 file
echo date('H:i:s') , " Write to Excel2007 format" , EOL;
$callStartTime = microtime(true);

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
$callEndTime = microtime(true);
$callTime = $callEndTime - $callStartTime;

echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo(__FILE__, PATHINFO_BASENAME)) , EOL;
echo 'Call time to write Workbook was ' , sprintf('%.4f',$callTime) , " seconds" , EOL;
// Echo memory usage
echo date('H:i:s') , ' Current memory usage: ' , (memory_get_usage(true) / 1024 / 1024) , " MB" , EOL;


echo date('H:i:s') , " Reload workbook from saved file" , EOL;
$callStartTime = microtime(true);

$objPHPExcel = PHPExcel_IOFactory::load(str_replace('.php', '.xlsx', __FILE__));

$callEndTime = microtime(true);
$callTime = $callEndTime - $callStartTime;
echo 'Call time to reload Workbook was ' , sprintf('%.4f',$callTime) , " seconds" , EOL;
// Echo memory usage
echo date('H:i:s') , ' Current memory usage: ' , (memory_get_usage(true) / 1024 / 1024) , " MB" , EOL;


var_dump($objPHPExcel->getActiveSheet()->toArray());


// Echo memory peak usage
echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , EOL;

// Echo done
echo date('H:i:s') , " Done testing file" , EOL;
echo 'File has been created in ' , getcwd() , EOL;

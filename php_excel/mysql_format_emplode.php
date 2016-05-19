<?php
/**
 * Created by PhpStorm.
 * User: shihuipeng
 * Date: 16/5/18
 * Time: 下午6:22
 */

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
//date_default_timezone_set('Europe/London');

if (PHP_SAPI == 'cli')
    die('This example should only be run from a Web Browser');

/** Include PHPExcel */
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';


// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

// Set document properties
$objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
    ->setLastModifiedBy("Maarten Balliauw")
    ->setTitle("Office 2007 XLSX Test Document")
    ->setSubject("Office 2007 XLSX Test Document")
    ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
    ->setKeywords("office 2007 openxml php")
    ->setCategory("Test result file");


// Miscellaneous glyphs, UTF-8
//$objPHPExcel->setActiveSheetIndex(0)
//    ->setCellValue('A4', 'Miscellaneous glyphs')
//    ->setCellValue('A5', 'éàèùâêîôûëïüÿäöüç');



$objPHPExcel->setActiveSheetIndex(0)
    ->setCellValue('B1', 'filed')
    ->setCellValue('C1', 'type')
    ->setCellValue('D1', 'remark');

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
// Rename worksheet
$objPHPExcel->getActiveSheet()->setTitle('Simple');


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);


// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="01simple.xls"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
exit;

<?php
require dirname(__FILE__) . '/src/excel/excel_xml.php';
use Jingwu\Excel\Excel_XML;

$data = array(
        0 => array('Nr.', 'Name', 'E-Mail'),
        array(1, 'Oliver Schwarz', '100000@qq.com'),
        array(2, 'Hasematzel', '100001@qq.com'));

$xls = new Excel_XML;
$xls->addWorksheet('Names', $data);
$xls->sendWorkbook('test.xml');
//$xls->writeWorkbook('test.xml');


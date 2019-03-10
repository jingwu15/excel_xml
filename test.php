<?php
//require dirname(__FILE__) . '/src/excel/excel_xml.php';
require dirname(__FILE__) . '/vendor/autoload.php';
use Jingwu\Excel\Excel_XML;

$data = array(
    0 => array('ID', '用户名', '邮箱'),
    array(1, '张三', '100000@qq.com'),
    array(2, '李四', '100001@qq.com')
);

$xls = new Excel_XML;
$xls->addWorksheet('Names', $data);
$xls->sendWorkbook('test.xml');
//$xls->writeWorkbook('test.xml');


<?php
require '../vendor/autoload.php';
include '../src/QuickExcel.php';
//表头
$excelHead = ['id', '名称', '时间'];
$data = [];
for ($i = 0; $i <= 5; $i++) {
    $data[$i] = [
        $i + 1,
        '名称' . ($i + 1),
        date('Y-m-d H:i:s', time())
    ];
}
$fileName = '测试';
$excel = new QuickExcel();
$savePath = './'; //保存在当前文件夹
$fileName = $excel::excelOut($excelHead, $data, $fileName, './'); //返回fileName

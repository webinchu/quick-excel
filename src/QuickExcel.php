<?php

namespace webing;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Created by phpStorm.
 * User: webin
 * Date: 2021/11/12
 * Time: 16:03
 */
class QuickExcel
{
    /*
     * excelOut 表格输出(直接下载)
     * 注: 每个表头顺序($head) = 表格数据顺序($data)
     * @param array $head 表头
     * @param array|object $data 表数据(二维 数组/对象)
     * @param string $filename 文件名
     * @param string $filePath 文件保存地址 空:直接下载
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @return string 文件名
    */
    public static function excelOut(array $head, $data, string $filename, $filePath = '')
    {
        $fixed = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

        //实例化表格
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        //写入表头
        $i = 0;
        foreach ($head as $value) {
            $sheet->setCellValue($fixed[$i] . '1', $head[$i]);
            $i++;
        }

        //重构$data
        $t = 0;
        foreach ($data as $key => $value) {
            if (is_object($value) && $value->getData()) {
                foreach ($value->getData() as $v) {
                    //这里是因为model数据会出现null的数据而且保证model字段的值不能为null,如果字段值有null的可以考虑将对象($data)转换为数组
                    if ($v !== null) {
                        $newData[$key][$t] = $v;
                        $t++;
                    }
                }
            } else {
                foreach ($value as $v) {
                    $newData[$key][$t] = $v;
                    $t++;
                }
            }
            $t = 0;
        }
        //写入表数据
        $number = 1;
        foreach ($newData as $key => $value) {
            $number++;
            foreach ($value as $k => $v) {
                $sheet->setCellValue($fixed[$k] . $number, $v . "\t");
            }
        }
        return self::downloadExcel($spreadsheet, $filename, 'Xlsx', $filePath);
    }

    private static function downloadExcel($newExcel, $filename, $format, $filePath = '')
    {
        $filename = $filename . '.' . strtolower($format);
        $objWriter = IOFactory::createWriter($newExcel, $format);
        if ($filePath) {
            $objWriter->save($filePath . $filename);
            return $filename;
        } else {
            // $format只能为 Xlsx 或 Xls
            if ($format == 'Xlsx') {
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            } elseif ($format == 'Xls') {
                header('Content-Type: application/vnd.ms-excel');
            }
            header("Content-Disposition: attachment;filename="
                . $filename);
            header('Cache-Control: max-age=0');
            $objWriter->save('php://output');
        }
        exit;
    }
}

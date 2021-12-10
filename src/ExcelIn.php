<?php

namespace Phpfunction\Excel;

use PhpOffice\PhpSpreadsheet\IOFactory;

/**
 * Excel导出
 * Class ExcelOut
 * @package Phpfunction\Phpexcel
 */
class ExcelIn
{
    public $file = '';          // 文件名


    public function __construct()
    {
        // 设置php超时时间及内存
        set_time_limit(0);
        ini_set('memory_limit', '1024M');
    }


    /**
     * 数据列表
     */
    public function getData()
    {
        $file_name = $this->file;
        $type      = pathinfo($file_name);
        $type      = strtolower($type["extension"]);

        if ("xlsx" == $type) {
            $objReader = IOFactory::createReader("Xlsx");
        } elseif ("xls" == $type) {
            $objReader = IOFactory::createReader("Xls");
        } else {
            return "文件类型不正确";
        }

        $objPHPExcel   = $objReader->load($file_name);
        $sheet         = $objPHPExcel->getSheet(0);
        $highestRow    = $sheet->getHighestRow();    // 取得总行数
        $highestColumn = $sheet->getHighestColumn(); // 取得总列数
        $data          = [];

        for ($i = 1; $i <= $highestRow; $i++) {
            $tmp      = [];
            $tmp["A"] = $sheet
                ->getCell("A" . $i)
                ->getCalculatedValue();
            $tmp["B"] = $sheet
                ->getCell("B" . $i)
                ->getCalculatedValue();
            $data[]   = $tmp;
        }

        return $data;
    }


}

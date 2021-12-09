<?php

namespace Phpfunction\Excel;

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
     * 样式——对齐
     * @return array[]
     */
    private function exe()
    {
    }


}

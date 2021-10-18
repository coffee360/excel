<?php

namespace Phpfunction\Excel;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;

/**
 * Excel导出
 * Class ExcelOut
 * @package Phpfunction\Phpexcel
 */
class ExcelOut
{
    public $file  = '';          // 文件名
    public $title = '';          // 表名
    public $head  = [];          // 表头
    public $list  = [];          // 数据


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
    private function getStyleAlignment($horizontal = 2)
    {
        $horizontal_name = Alignment::HORIZONTAL_CENTER;
        if (1 == $horizontal) {
            $horizontal_name = Alignment::HORIZONTAL_LEFT;
        } elseif (3 == $horizontal) {
            $horizontal_name = Alignment::HORIZONTAL_RIGHT;
        }

        return [
            'alignment' => [
                //水平居中
                'horizontal' => $horizontal_name,

                //垂直居中
                'vertical'   => Alignment::VERTICAL_CENTER,
            ],
        ];
    }


    /**
     * 保存
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function save()
    {
        if (empty($this->file)) {
            return [
                'errcode' => 1,
                'errmsg'  => 'file不能为空',
            ];
        }

        if (empty($this->title)) {
            return [
                'errcode' => 1,
                'errmsg'  => 'title不能为空',
            ];
        }

        $newExcel = new Spreadsheet();            //创建一个新的excel文档
        $objSheet = $newExcel->getActiveSheet();  //获取当前操作sheet的对象
        $objSheet->setTitle($this->title);        //设置当前sheet的标题

        if (!empty($this->head)) {
            $list_new   = [];
            $list_new[] = $this->head;

            foreach ($this->list as $k => $v) {
                $tmp = [];
                foreach ($this->head as $k2 => $v2) {
                    $tmp[$k2] = $v[$k2];
                }
                $list_new[] = $tmp;
            }
        } else {
            $list_new = $this->list;
        }

        foreach ($list_new as $k => $v) {
            $v = array_values($v);
            foreach ($v as $k2 => $v2) {
                if ($k2 < 26) {
                    $col = chr(65 + $k2);
                } elseif ($k2 >= 26) {
                    $col = 'A' . chr(65 + $k2 - 26);
                }

                $objSheet->getColumnDimension($col)
                    ->setWidth(30);

                if (!empty($this->head) && empty($k)) {
                    $objSheet->getStyle($col . ($k + 1))
                        ->getFont()
                        ->setBold(true); //字体加粗

                    $objSheet->getStyle($col . ($k + 1))
                        ->getFill()
                        ->setFillType(Fill::FILL_SOLID)
                        ->getStartColor()
                        ->setARGB('FF808080');
                }

                $objSheet->setCellValue($col . ($k + 1), " " . (new StringApp())->removeEmoji($v2));

                // 数字右对齐
                $style_alignment = $this->getStyleAlignment();
                if (is_numeric($v2)) {
                    $style_alignment = $this->getStyleAlignment(3);
                }
                $newExcel->getActiveSheet()
                    ->getStyle($col . ($k + 1))
                    ->applyFromArray($style_alignment);
            }
        }

        /*--------------下面是设置其他信息------------------*/
        header('Content-Type: application/vnd.ms-excel');
        header("Content-Disposition: attachment;filename=" . urlencode($this->file) . ".xls");
        header('Cache-Control: max-age=0');
        $objWriter = IOFactory::createWriter($newExcel, 'Xls');
        $objWriter->save('php://output');
    }

}

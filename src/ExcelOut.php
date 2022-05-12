<?php

namespace Phpfunction\Excel;

use Phpfunction\App\StringApp;
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
    public $width = [];          // 表头
    public $list  = [];          // 数据

    public $file_save = "";     // 保存的文件


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

                $objSheet->setCellValue($col . ($k + 1), (new StringApp())->removeEmoji($v2));

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


    /**
     * 保存
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function saveWithStyle()
    {
        if (empty($this->file) && empty($this->file_save)) {
            return [
                'errcode' => 1,
                'errmsg'  => 'file和file_save不能同时为空',
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

                // 1.样式传值
                $width = $this->width[$k2] ?? 30;

                $text       = $v2;
                $num_col    = 1;
                $num_row    = 1;
                $height     = 0;
                $size       = 0;
                $bold       = 0;
                $text_align = 0;
                if (is_array($v2)) {
                    $text    = $v2['text'] ?? "";
                    $num_col = $v2['num_col'] ?? 1;
                    $num_row = $v2['num_row'] ?? 1;
                    $height  = $v2['height'] ?? 0;
                    $size    = $v2['size'] ?? 0;
                    $bold    = $v2['bold'] ?? 0;
                    // 1=左，2=中，3=右
                    $text_align = $v2['text_align'] ?? 0;
                }

                // 2.写入数据
                // 2.1.合并单元格
                if ($num_col > 1) {
                    if ($k2 < 26) {
                        $col_to = chr(65 + $k2 + $num_col - 1);
                    } elseif ($k2 >= 26) {
                        $col_to = 'A' . chr(65 + $k2 + $num_col - 1 - 26);
                    }
                    $objSheet->mergeCells($col . ($k + 1) . ":" . $col_to . ($k + 1));
                }
                if ($num_row > 1) {
                    $objSheet->mergeCells($col . ($k + 1) . ":" . $col . ($k + $num_row));
                }
                $objSheet->getColumnDimension($col)
                    ->setWidth($width);

                // 2.2.单元格字体大小
                if ($size > 0) {
                    $objSheet->getStyle($col . ($k + 1))
                        ->getFont()
                        ->setSize(20);
                }

                // 2.2.单元格是否加粗
                if ($bold > 0) {
                    $objSheet->getStyle($col . ($k + 1))
                        ->getFont()
                        ->setBold(true);
                }

                // 2.2.单元格行高
                if ($height > 0) {
                    $objSheet->getRowDimension(($k + 1))
                        ->setRowHeight($height);
                }

                // 2.2.单元格内容
                $objSheet->setCellValue($col . ($k + 1), (new StringApp())->removeEmoji($text));

                // 2.3.单元格对对齐，数字右对齐
                $style_alignment = $this->getStyleAlignment();
                if (is_numeric($text)) {
                    $style_alignment = $this->getStyleAlignment(3);
                }
                if (!empty($text_align)) {
                    $style_alignment = $this->getStyleAlignment($text_align) ?? $style_alignment;
                }
                $newExcel->getActiveSheet()
                    ->getStyle($col . ($k + 1))
                    ->applyFromArray($style_alignment);
            }
        }

        $objWriter = IOFactory::createWriter($newExcel, 'Xls');
        if (!empty($this->file_save)) {
            $objWriter->save($this->file_save);
        } else {
            /*--------------下面是设置其他信息------------------*/
            header('Content-Type: application/vnd.ms-excel');
            header("Content-Disposition: attachment;filename=" . urlencode($this->file) . ".xls");
            header('Cache-Control: max-age=0');

            $objWriter->save('php://output');
        }
    }

}

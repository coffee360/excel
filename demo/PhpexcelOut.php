<?php

require "../vendor/autoload.php";

$a        = new \Phpfunction\Excel\ExcelOut();
$a->head  = [
    'name' => '姓名',
    'sex'  => '性别（男女）',
    'tel'  => '电话'
];
$a->list  = [
    [
        'name' => '小名',
        'sex'  => '男'
    ],
    [
        'name' => '小名1',
        'sex'  => '男1'
    ],
    [
        'sex'  => '男2',
        'name' => '小名2',

    ],
    [
        'name' => '小名3',
        'tel'  => '13512345678',

    ],
    [
        'sex' => '男4',

    ]
];
$a->file  = '通寻路';
$a->title = '通寻路';
print_r($a->save());

<?php

use Biin2013\Spreadsheet\Export;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PHPUnit\Framework\Attributes\Test;
use PHPUnit\Framework\TestCase;

class ExportTest extends TestCase
{
    private function data(): array
    {
        return [
            [
                'config' => [
                    'sheet_name' => 'sheet',
                    'data_children_key' => 'children',
                    'data_children_text_key' => 'name',
                    'data_children_cell_merge' => true,
                    'data_value_key' => 'name',
                    'start_row' => 1,
                    'start_column' => 1,
                    'alignment_horizontal' => Alignment::HORIZONTAL_CENTER,
                    'alignment_vertical' => Alignment::VERTICAL_CENTER,
                    'header_font_size' => 16,
                    'header_font_bold' => false,
                    'default_column_width' => 'auto',
                    'default_row_height' => 25
                ],
                'header' => [
                    [
                        'field' => 'address',
                        'name' => '地址',
                        'children' => [
                            [
                                'field' => 'province',
                                'name' => '省'
                            ],
                            [
                                'field' => 'city',
                                'name' => '市'
                            ],
                            [
                                'field' => 'district',
                                'name' => '区'
                            ]
                        ]
                    ],
                    [
                        'field' => 'userinfo',
                        'name' => '用户信息',
                        'children' => [
                            [
                                'field' => 'name',
                                'name' => '姓名'
                            ],
                            [
                                'field' => 'sex',
                                'name' => '性别',
                                'config' => [
                                    'format' => fn($val) => $val == 1 ? '男' : '女',
                                    'custom' => function (Worksheet $worksheet, $val, $data, $row, $col) {
                                        $val == 1
                                            ? $worksheet->getStyle([$col, $row])->getFont()->setColor(new Color(Color::COLOR_BLUE))
                                            : $worksheet->getStyle([$col, $row])->getFont()->setColor(new Color(Color::COLOR_RED));
                                    }
                                ]
                            ],
                            [
                                'field' => 'birthday',
                                'name' => '生日'
                            ]
                        ]
                    ],
                    [
                        'field' => 'created_at',
                        'name' => '创建时间',
                        'config' => [
                            'width' => 40
                        ]
                    ]
                ],
                'data' => [
                    [
                        'province' => '北京',
                        'city' => '北京',
                        'district' => '东城区',
                        'name' => '张三',
                        'sex' => 1,
                        'birthday' => '2022-01-01',
                        'created_at' => '2022-01-01 00:00:00'
                    ],
                    [
                        'name' => '湖北',
                        'children' => [
                            [
                                'name' => '武汉市',
                                'children' => [
                                    [
                                        'name' => '武昌区',
                                        'children' => [
                                            [
                                                'name' => '李四',
                                                'sex' => 1,
                                                'birthday' => '1990-01-01',
                                                'created_at' => '2022-01-01 00:00:00'
                                            ],
                                            [
                                                'name' => '王五',
                                                'sex' => 2,
                                                'birthday' => '1991-11-01',
                                                'created_at' => '2022-01-01 00:00:00'
                                            ],
                                            [
                                                'name' => '赵六',
                                                'sex' => 1,
                                                'birthday' => '1992-12-31',
                                                'created_at' => '2022-01-01 00:00:00'
                                            ]
                                        ]
                                    ],
                                    [
                                        'name' => '汉口区',
                                        'children' => [
                                            [
                                                'name' => '周七',
                                                'sex' => 1,
                                                'birthday' => '1993-01-01',
                                                'created_at' => '2022-01-01 00:00:00'
                                            ],
                                            [
                                                'name' => '王八',
                                                'sex' => 1,
                                                'birthday' => '1993-01-01',
                                                'created_at' => '2022-01-01 00:00:00'
                                            ]
                                        ]
                                    ],
                                    [
                                        'name' => '江夏区',
                                        'children' => [
                                            [
                                                'name' => '王九',
                                                'sex' => 1,
                                                'birthday' => '1994-01-01',
                                                'created_at' => '2022-01-01 00:00:00'
                                            ]
                                        ]
                                    ]
                                ]
                            ],
                            [
                                'name' => '咸宁市',
                                'children' => [
                                    [
                                        'name' => '咸安区',
                                        'children' => [
                                            [
                                                'name' => '张六',
                                                'sex' => 1,
                                                'birthday' => '1989-3-01',
                                                'created_at' => '2022-01-01 00:00:00'
                                            ],
                                            [
                                                'name' => '王七',
                                                'sex' => 1,
                                                'birthday' => '1991-6-21',
                                                'created_at' => '2022-01-01 00:00:00'
                                            ]
                                        ]
                                    ]
                                ]
                            ]
                        ]
                    ],
                    [
                        'province' => '广州',
                        'city' => '深圳',
                        'district' => '宝安区',
                        'name' => '小丽',
                        'sex' => 2,
                        'birthday' => '2008-08-08',
                        'created_at' => '2021-01-01 00:00:00'
                    ]
                ]
            ]
        ];
    }

    #[Test]
    public function export()
    {
        //print_r((new Export())->build($this->data())->export('./Export'));
        print_r(Export::make()->build($this->data())->export('./export'));
    }
}
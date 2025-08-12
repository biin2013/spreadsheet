<?php

namespace Biin2013\Spreadsheet;

use Closure;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;

class Export
{
    protected string $fileType = 'Xlsx';
    protected array $data = [];
    protected Spreadsheet $spreadsheet;
    public const ROW_KEY = '_row';
    public const MERGE_ROW_KEY = '_merge_row';
    public const MERGE_COL_KEY = '_merge_col';
    public const DATA_CHILDREN_KEY = 'children';
    private int $currentSheetIndex = 0;
    private array $merges = [];
    private array $currentConfig = [];
    private array $defaultConfig = [
        'sheet_name' => 'Sheet',
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
    ];
    private array $flattenHeader = [];
    private array $flattenHeaderData = [];
    private array $flattenData = [];

    public function __construct(array $data = [], Spreadsheet $spreadsheet = null)
    {
        $this->spreadsheet = $spreadsheet ?? new Spreadsheet();
        $this->data = $data;
    }

    public function data(array $data = null): array|static
    {
        if (is_null($data)) {
            return $this->data;
        }

        $this->data = $data;

        return $this;
    }

    public function xlsx(): static
    {
        $this->fileType = 'Xlsx';

        return $this;
    }

    public function xls(): static
    {
        $this->fileType = 'Xls';

        return $this;
    }

    public function csv(): static
    {
        $this->fileType = 'Csv';

        return $this;
    }

    public function writer(): IWriter
    {
        return IOFactory::createWriter($this->spreadsheet, $this->fileType);
    }

    public static function make(array $data = []): static
    {
        return new static($data);
    }

    public function build(array $data = null): static
    {
        if ($data) {
            $this->data = $data;
        }
        foreach ($this->data as $index => $val) {
            $this->currentSheetIndex = $index;
            $this->currentConfig = array_merge($this->defaultConfig, $val['config'] ?? []);

            $this->resolveHeaderMerge($val['header']);
            $this->resolveDataMerge($val['data']);
            $this->merges[$index] = [];

            $headerRow = max(array_column($val['header'], self::MERGE_ROW_KEY));
            $this->flattenHeader[$index] = $this->flattenHeader($val['header'], $headerRow);
            $this->flattenHeaderData[$index] = $this->flattenHeaderData($val['header']);
            $this->flattenData[$index] = $this->flattenData(
                $val['data'],
                $this->currentConfig['start_row'] + $headerRow,
                $this->currentConfig['start_column']
            );

            $this->spreadsheet
                ->getDefaultStyle()
                ->getAlignment()
                ->setHorizontal($this->currentConfig['alignment_horizontal'])
                ->setVertical($this->currentConfig['alignment_vertical']);

            $this->buildSheet($index, $headerRow);
        }

        return $this;
    }

    private function resolveHeaderMerge(array &$header): void
    {
        foreach ($header as &$value) {
            if ($value['children'] ?? null) {
                $this->resolveHeaderMerge($value['children']);
                $value[self::MERGE_COL_KEY] = array_sum(array_column($value['children'], self::MERGE_COL_KEY));
                $value[self::MERGE_ROW_KEY] = 1;
                $value[self::ROW_KEY] = max(array_column($value['children'], self::ROW_KEY)) + 1;
            } else {
                $value[self::MERGE_COL_KEY] = 1;
                $value[self::MERGE_ROW_KEY] = 1;
                $value[self::ROW_KEY] = 1;
            }
        }

        foreach ($header as &$val) {
            if (empty($val['children'])) {
                $val[self::MERGE_ROW_KEY] = max(array_column($header, self::ROW_KEY));
            }
        }
    }

    private function resolveDataMerge(array &$data): void
    {
        foreach ($data as &$value) {
            if ($value[self::DATA_CHILDREN_KEY] ?? null) {
                $this->resolveDataMerge($value[self::DATA_CHILDREN_KEY]);
                $value[self::MERGE_ROW_KEY] = array_sum(array_column($value[self::DATA_CHILDREN_KEY], self::MERGE_ROW_KEY));
            } else {
                $value[self::MERGE_COL_KEY] = 1;
                $value[self::MERGE_ROW_KEY] = 1;
            }
        }
    }

    private function flattenHeader(array $header, int $startRow = 0, int $startColumn = 0): array
    {
        $flattenHeader = [];
        foreach ($header as $key => $value) {
            $startColumn++;
            if (!empty($value['children'])) {
                $startRow++;
                //$this->merges[$this->currentSheetIndex][] = [$startColumn, $startRow, $startColumn + $value[self::MERGE_COL_KEY], $startRow + $value[self::MERGE_ROW_KEY]];
                $result = $this->flattenHeader($value['children'], $startRow, $startColumn);
                $flattenHeader = array_merge($flattenHeader, $result);
            } else {
                $flattenHeader[] = $this->getHeaderFields($value);
            }
        }

        return $flattenHeader;
    }

    private function flattenHeaderData(
        array $header,
        int   $currentRow = 0,
    ): array
    {
        $data = [];
        foreach ($header as $value) {
            $data[$currentRow] = $data[$currentRow] ?? [];
            $data[$currentRow][] = $this->getHeaderFields($value);
            if (!empty($value['children'])) {
                $nextRow = $currentRow + 1;
                $data[$nextRow] = array_merge(
                    $data[$nextRow] ?? [],
                    ...$this->flattenHeaderData(
                    $value['children'],
                    $nextRow
                ));
            }
        }
        return $data;
    }

    private function getHeaderFields(array $value): array
    {
        return [
            'field' => $value['field'],
            'name' => $value['name'],
            'config' => $value['config'] ?? [],
            self::MERGE_COL_KEY => $value[self::MERGE_COL_KEY],
            self::MERGE_ROW_KEY => $value[self::MERGE_ROW_KEY],
            self::ROW_KEY => $value[self::ROW_KEY]
        ];
    }

    private function flattenData(array $data, int $startRow = 1, int $startColumn = 1, array $item = []): array
    {
        $flattenData = [];
        foreach ($data as $value) {
            if (empty($value[self::DATA_CHILDREN_KEY])) {
                $flattenData[] = array_merge($item, $value);
            } else {
                $this->merges[$this->currentSheetIndex][] = [
                    $startColumn,
                    $startRow,
                    $startColumn,
                    $startRow + $value[self::MERGE_ROW_KEY] - 1
                ];
                $key = $this->flattenHeader[$this->currentSheetIndex][$startColumn - 1]['field'];
                $item[$key] = $value[$this->currentConfig['data_children_text_key']];
                $result = $this->flattenData(
                    $value[self::DATA_CHILDREN_KEY],
                    $startRow,
                    $startColumn + 1,
                    $item
                );
                $startRow += $value[self::MERGE_ROW_KEY] - 1;
                $flattenData = array_merge($flattenData, $result);
            }
            $startRow++;
        }

        return $flattenData;
    }

    private function buildSheet(int $sheetIndex, int $headerRow): void
    {
        $header = $this->flattenHeader[$sheetIndex];
        $headerData = $this->flattenHeaderData[$sheetIndex];
        $data = $this->flattenData[$sheetIndex];

        $this->spreadsheet
            ->createSheet($sheetIndex)
            ->setTitle($this->currentConfig['sheet_name']);

        $this->buildSheetHeader($headerData);
        $this->buildSheetData($header, $data, $headerRow);

        if ($this->currentConfig['data_children_cell_merge']) {
            array_map(fn($item) => $this->spreadsheet->getActiveSheet()->mergeCells($item), $this->merges[$sheetIndex]);
        }
    }

    private function buildSheetHeader(array $header): void
    {
        $spreadsheet = $this->spreadsheet->getActiveSheet();
        $spreadsheet->getDefaultRowDimension()->setRowHeight($this->currentConfig['default_row_height']);

        foreach ($header as $row => $value) {
            $col = 1;
            foreach ($value as $val) {
                $spreadsheet->setCellValue([$col, $row + 1], $val['name']);

                $spreadsheet->getStyle([$col, $row + 1])
                    ->getFont()
                    ->setSize($this->currentConfig['header_font_size'])
                    ->setBold($this->currentConfig['header_font_bold']);

                $width = $val['config']['width'] ?? $this->currentConfig['default_column_width'];
                if ($width == 'auto') {
                    $spreadsheet->getColumnDimensionByColumn($col)->setAutoSize(true);
                } else {
                    $spreadsheet->getColumnDimensionByColumn($col)->setWidth($width);
                }

                $nextCol = $col + $val[self::MERGE_COL_KEY];
                if ($val[self::MERGE_COL_KEY] > 1 || $val[self::MERGE_ROW_KEY] > 1) {
                    $spreadsheet->mergeCells([$col, $row + 1, $nextCol - 1, $row + $val[self::MERGE_ROW_KEY]]);
                }
                $col = $nextCol;
            }
        }
    }

    private function buildSheetData(array $header, array $data, int $startRow): void
    {
        $worksheet = $this->spreadsheet->getActiveSheet();
        foreach ($data as $row) {
            $startRow++;
            $col = 1;
            foreach ($header as $val) {
                $format = $val['config']['format'] ?? null;
                $text = is_callable($format)
                    ? $format($row[$val['field']], $row, $worksheet)
                    : $row[$val['field']];
                $worksheet->setCellValue([$col, $startRow], $text);

                $custom = $val['config']['custom'] ?? null;
                if (is_callable($custom)) {
                    $custom($worksheet, $row[$val['field']], $row, $startRow, $col);
                }
                $col++;
            }
        }
    }

    public function export(string $rootPath = '.', bool|Closure|string $datePath = true, string|Closure|null $fileName = null): array
    {
        $file = $this->resolveFileName($rootPath, $datePath, $fileName);
        $this->writer()->save($file['full']);

        return $file;
    }

    /**
     * @param string $rootPath
     * @param bool|Closure|string $datePath
     * @param string|Closure|null $fileName
     * @return array
     */
    private function resolveFileName(string $rootPath = '.', bool|Closure|string $datePath = true, Closure|string|null $fileName = null): array
    {
        $date = is_bool($datePath)
            ? ($datePath ? date('Ymd') : '')
            : (is_callable($datePath) ? $datePath() : date($datePath));
        $fullPath = rtrim($rootPath, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . rtrim($date, DIRECTORY_SEPARATOR);

        if (!is_dir($fullPath)) {
            mkdir($fullPath, 0777, true);
        }

        $fileName = is_string($fileName)
            ? $fileName
            : (is_callable($fileName) ? $fileName() : str_replace('.', '', microtime(true)));

        $type = strtolower($this->fileType);
        $full = $fullPath . DIRECTORY_SEPARATOR . trim($fileName, DIRECTORY_SEPARATOR) . '.' . $type;

        return [
            'root' => $rootPath,
            'date' => $date,
            'file' => $fileName,
            'type' => $type,
            'full' => $full
        ];
    }
}
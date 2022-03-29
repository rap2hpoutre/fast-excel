<?php

namespace Rap2hpoutre\FastExcel\Contracts;

interface ExcelInterface
{
    public function data($data);

    public function sheet($sheet_number);

    public function withoutHeaders();

    public function withSheetsNames();

    public function startRow(int $row);

    public function transpose();

    public function configureCsv($delimiter = ',', $enclosure = '"', $encoding = 'UTF-8', $bom = false);

    public function configureReaderUsing(?callable $callback = null);

    public function configureWriterUsing(?callable $callback = null);
}

<?php

namespace Rap2hpoutre\FastExcel\Contracts;

use Box\Spout\Common\Entity\Style\Style;

interface ExportInterface
{
    public function export($path, callable $callback = null);

    public function download($path, callable $callback = null);

    public function headerStyle(Style $style);

    public function rowsStyle(Style $style);
}

<?php

namespace Rap2hpoutre\FastExcel\Exceptions;

use Exception;

class SheetNameMissing extends Exception
{
    public function __construct(string $sheetName)
    {
        parent::__construct("Sheet name [$sheetName] is missing.");
    }
}

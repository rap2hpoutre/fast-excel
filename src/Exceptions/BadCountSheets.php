<?php

namespace Rap2hpoutre\FastExcel\Exceptions;

use Exception;

class BadCountSheets extends Exception
{
    public function __construct()
    {
        parent::__construct("You file does not contains more than one sheet.");
    }
}

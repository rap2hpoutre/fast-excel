<?php

namespace Rap2hpoutre\FastExcel\Contracts;

interface ImportInterface
{
    public function import($path, callable $callback = null);

    public function importSheets($path, callable $callback = null);
}

<?php

namespace Rap2hpoutre\FastExcel\Facades;

use Illuminate\Support\Facades\Facade;

class FastExcel extends Facade
{
    /**
     * Get the registered name of the component.
     *
     * @return string
     */
    protected static function getFacadeAccessor()
    {
        return 'fastexcel';
    }
}

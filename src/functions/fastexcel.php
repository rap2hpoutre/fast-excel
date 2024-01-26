<?php

use Rap2hpoutre\FastExcel\SheetCollection;

if (!function_exists('fastexcel')) {
    /**
     * Return app instance of FastExcel.
     *
     * @return Rap2hpoutre\FastExcel\FastExcel
     */
    function fastexcel($data = null)
    {
        if ($data instanceof SheetCollection) {
            return app()->make('fastexcel')->data($data);
        }

        if (is_object($data) && method_exists($data, 'toArray')) {
            $data = $data->toArray();
        }

        return $data === null ? app()->make('fastexcel') : app()->makeWith('fastexcel', $data);
    }
}

if (!function_exists('obtainTempDirectory')) {
    /**
     * Get a temporary directory.
     *
     * @return TemporaryDirectory
     */
    function obtainTempDirectory(): TemporaryDirectory
    {
        return (new TemporaryDirectory())
            ->name('fast-excel-'.microtime())
            ->force()
            ->deleteWhenDestroyed()
            ->create();
    }
}

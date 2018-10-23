<?php

if (!function_exists('fastexcel')) {
    /**
     * Return app instance of FastExcel.
     *
     * @return Rap2hpoutre\FastExcel\FastExcel
     */
    function fastexcel($data = null)
    {
        if (is_object($data) && method_exists($data, 'toArray')) {
            $data = $data->toArray();
        }

        return blank($data) ? app()->make('fastexcel') : app()->makeWith('fastexcel', $data);
    }
}

<?php

namespace Rap2hpoutre\FastExcel\Tests;

use Orchestra\Testbench\TestCase as OrchestraTestCase;
use Rap2hpoutre\FastExcel\Providers\FastExcelServiceProvider;

/**
 * Class TestCase.
 */
class TestCase extends OrchestraTestCase
{
    /**
     * @return \Illuminate\Support\Collection
     */
    protected function collection()
    {
        return collect([
            ['col1' => 'row1 col1', 'col2' => 'row1 col2'],
            ['col1' => 'row2 col1', 'col2' => ''],
            ['col1' => 'row3 col1', 'col2' => 'row3 col2'],
        ]);
    }

     /**
     * @param  \Illuminate\Foundation\Application  $app
     * @return array
     */
    protected function getPackageProviders($app)
    {
        return [
            FastExcelServiceProvider::class,
        ];
    }
}

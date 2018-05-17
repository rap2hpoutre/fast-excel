<?php

namespace Rap2hpoutre\FastExcel\Tests;

use PHPUnit\Framework\TestCase as BaseTestCase;

/**
 * Class TestCase.
 */
class TestCase extends BaseTestCase
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
}

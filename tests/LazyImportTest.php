<?php

namespace Rap2hpoutre\FastExcel\Tests;

use Illuminate\Support\LazyCollection;
use Rap2hpoutre\FastExcel\FastExcel;

class LazyImportTest extends TestCase
{
    /**
     * Ensure importLazy returns a LazyCollection and yields same data as import.
     */
    public function testImportLazyXlsx()
    {
        $fe = new FastExcel();
        $lazy = $fe->importLazy(__DIR__.'/test1.xlsx');
        $this->assertInstanceOf(LazyCollection::class, $lazy);
        // Materialize to compare with existing helper collection()
        $this->assertEquals($this->collection(), $lazy->collect());
    }

    /**
     * Ensure importLazy supports callback mapping similar to import.
     */
    public function testImportLazyWithCallback()
    {
        $fe = new FastExcel();
        $lazy = $fe->importLazy(__DIR__.'/test1.xlsx', function ($row) {
            return [
                'col1' => $row['col1'],
                'col2' => $row['col2'],
            ];
        });

        $expected = (new FastExcel())->import(__DIR__.'/test1.xlsx', function ($row) {
            return [
                'col1' => $row['col1'],
                'col2' => $row['col2'],
            ];
        });

        $this->assertEquals($expected, $lazy->collect());
    }
}

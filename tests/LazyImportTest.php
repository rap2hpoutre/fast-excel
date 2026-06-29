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

    /**
     * The lazy path shares normalizeRow() with the eager import, so it must also
     * de-duplicate headers (#312/#259): duplicates get a numeric suffix and empty
     * headers a positional name, instead of colliding in array_combine().
     */
    public function testImportLazyDeduplicatesHeaders()
    {
        $file = __DIR__.'/lazy_dup_headers.csv';
        file_put_contents($file, "Name,Name,,Age\nJoe,Smith,x,30\nJane,Doe,y,25\n");

        $rows = (new FastExcel())->importLazy($file)->collect();

        $first = $rows->first();
        $this->assertSame(['Name', 'Name_2', 'column_3', 'Age'], array_keys($first));
        $this->assertSame('Joe', $first['Name']);
        $this->assertSame('Smith', $first['Name_2']);   // would have been lost before
        $this->assertSame('x', $first['column_3']);
        $this->assertSame('30', $first['Age']);
        $this->assertCount(2, $rows);

        unlink($file);
    }
}

<?php

namespace Rap2hpoutre\FastExcel\Tests;

use Rap2hpoutre\FastExcel\FastExcel;

/**
 * Class ChunkTest.
 */
class LimitTest extends TestCase
{
    public function collectionGenerator($n)
    {
        for ($i = 1; $i <= $n; $i++) {
            yield collect(['a' => 'b', 'c' => 'd']);
        }
    }

    public function arrayGenerator($n)
    {
        for ($i = 1; $i <= $n; $i++) {
            yield ['a' => 'b', 'c' => 'd'];
        }
    }

    public function testWithGenerator()
    {
        (new FastExcel($this->collectionGenerator(100)))->export(__DIR__.'/test-generator.xlsx');
        $result = (new FastExcel())->limitRows(10)->import(__DIR__.'/test-generator.xlsx');
        $this->assertEquals(
            collect($this->arrayGenerator(10)),
            $result
        );
        unlink(__DIR__.'/test-generator.xlsx');
    }
}

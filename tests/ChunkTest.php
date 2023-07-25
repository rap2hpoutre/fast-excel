<?php

namespace Rap2hpoutre\FastExcel\Tests;

use Rap2hpoutre\FastExcel\FastExcel;

/**
 * Class ChunkTest.
 */
class ChunkTest extends TestCase
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
        $n = 10;
        (new FastExcel($this->collectionGenerator($n)))->export(__DIR__.'/test-generator.xlsx');
        $result = (new FastExcel())->import(__DIR__.'/test-generator.xlsx');
        $this->assertEquals(
            collect($this->arrayGenerator($n)),
            $result
        );
        unlink(__DIR__.'/test-generator.xlsx');
    }
}

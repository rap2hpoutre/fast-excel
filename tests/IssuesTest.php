<?php

namespace Rap2hpoutre\FastExcel\Tests;

use Box\Spout\Common\Type;
use Box\Spout\Reader\ReaderFactory;
use Illuminate\Support\Collection;
use Rap2hpoutre\FastExcel\FastExcel;
use Rap2hpoutre\FastExcel\SheetCollection;

/**
 * Class IssuesTest.
 */
class IssuesTest extends TestCase
{
    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function testIssue11()
    {
        $original_collection = $this->collection()->map(function ($v) {
            return array_merge($v, ['test' => ['hello', 'hi']]);
        });
        (new FastExcel(clone $original_collection))->export(__DIR__.'/test2.xlsx');
        $this->assertNotEquals($original_collection, (new FastExcel())->import(__DIR__.'/test2.xlsx'));
        $this->assertEquals($this->collection(), (new FastExcel())->import(__DIR__.'/test2.xlsx'));
        unlink(__DIR__.'/test2.xlsx');
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function testIssue18()
    {
        $collection = (new FastExcel())->import(__DIR__.'/test18.csv');
        $this->assertInstanceOf(Collection::class, $collection);
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function testIssue20()
    {
        chdir(__DIR__);
        $path = (new FastExcel($this->collection()))->export('test2.xlsx');
        $this->assertEquals(__DIR__.'/test2.xlsx', $path);
        unlink($path);
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function testIssue19()
    {
        chdir(__DIR__);
        $path = (new FastExcel(collect([['a' => 1, 'b' => 'n', 'c' => 1.32, 'd' => []]])))->export('test2.xlsx');
        $this->assertEquals(collect([['a' => '1', 'b' => 'n', 'c' => '1.32']]), (new FastExcel())->import(__DIR__.'/test2.xlsx'));
        unlink($path);
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function testIssue26()
    {
        chdir(__DIR__);
        foreach ([[[]], null, [null]] as $value) {
            $path = (new FastExcel($value))->export('test2.xlsx');
            $this->assertEquals(collect([]), (new FastExcel())->import(__DIR__.'/test2.xlsx'));
            unlink($path);
        }
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function testIssue32()
    {
        $original_collection = collect([
            [
                'duration_in_months' => 1,
                'expires_at'         => '2018-08-06',
            ],
            [
                'duration_in_months' => null,
                'expires_at'         => '1970-01-01',
            ],
        ]);
        (new FastExcel(clone $original_collection))->export(__DIR__.'/test2.xlsx');
        $res = (new FastExcel())->import(__DIR__.'/test2.xlsx');
        $this->assertEquals($original_collection[1], $res[1]);
        unlink(__DIR__.'/test2.xlsx');
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function testIssue40()
    {
        $col = new SheetCollection(['1st Sheet' => $this->collection(), '2nd Sheet' => $this->collection()]);
        (new FastExcel($col))->export(__DIR__.'/test2.xlsx');
        $reader = ReaderFactory::create(Type::XLSX);
        $reader->open(__DIR__.'/test2.xlsx');
        foreach ($reader->getSheetIterator() as $key => $sheet) {
            $this->assertEquals($sheet->getName(), $key === 2 ? '2nd Sheet' : '1st Sheet');
        }
        $reader->close();
        unlink(__DIR__.'/test2.xlsx');
    }
}

<?php
namespace Rap2hpoutre\FastExcel\Tests;

use PHPUnit\Framework\TestCase;
use Rap2hpoutre\FastExcel\FastExcel;

/**
 * Class FastExcelTest
 * @package Rap2hpoutre\FastExcel\Tests
 */
class FastExcelTest extends TestCase
{

    /**
     * @return \Illuminate\Support\Collection
     */
    private function collection()
    {
        return collect([
            ['col1' => 'row1 col1', 'col2' => 'row1 col2'],
            ['col1' => 'row2 col1', 'col2' => ''],
            ['col1' => 'row3 col1', 'col2' => 'row3 col2'],
        ]);
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function testImportXlsx()
    {
        $collection = (new FastExcel)->import(__DIR__ . '/test1.xlsx');
        $this->assertEquals($this->collection(), $collection);
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function testImportCsv()
    {
        $original_collection = $this->collection();

        $collection = (new FastExcel)->import(__DIR__ . '/test2.csv');
        $this->assertEquals($original_collection, $collection);

        $collection = (new FastExcel)->configureCsv(';')->import(__DIR__ . '/test1.csv');
        $this->assertEquals($original_collection, $collection);
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    private function export($file)
    {
        $original_collection = $this->collection();

        (new FastExcel($original_collection))->export($file);
        $this->assertEquals($original_collection, (new FastExcel)->import($file));
        unlink($file);
    }


    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function testExportXlsx()
    {
        $this->export(__DIR__ . '/test2.xlsx');
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function testExportCsv()
    {
        $this->export(__DIR__ . '/test3.csv');
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function testExcelImportWithCallback()
    {
        $collection = (new FastExcel)->import(__DIR__ . '/test1.xlsx', function ($value) {
            return [
                'test' => $value['col1']
            ];
        });
        $this->assertEquals(
            collect([['test' => 'row1 col1'], ['test' => 'row2 col1'], ['test' => 'row3 col1']]),
            $collection
        );

        $collection = (new FastExcel)->import(__DIR__ . '/test1.xlsx', function ($value) {
            return new Dumb($value['col1']);
        });
        $this->assertEquals(
            collect([new Dumb('row1 col1'), new Dumb('row2 col1'), new Dumb('row3 col1')]),
            $collection
        );
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function testIssue11()
    {
        $original_collection = $this->collection()->map(function($v) {
            return array_merge($v, ['test' => ['hello', 'hi']]);
        });
        (new FastExcel($original_collection))->export(__DIR__ . '/test2.xlsx');
        $this->assertNotEquals($original_collection, (new FastExcel)->import(__DIR__ . '/test2.xlsx'));
        $this->assertEquals($this->collection(), (new FastExcel)->import(__DIR__ . '/test2.xlsx'));
        unlink(__DIR__ . '/test2.xlsx');
    }


    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function testExcelExportWithCallback()
    {
        (new FastExcel($this->collection()))->export(__DIR__ . '/test2.xlsx', function ($value) {
            return [
                'test' => $value['col1']
            ];
        });
        $this->assertEquals(
            collect([['test' => 'row1 col1'], ['test' => 'row2 col1'], ['test' => 'row3 col1']]),
            (new FastExcel)->import(__DIR__ . '/test2.xlsx')
        );
    }
}

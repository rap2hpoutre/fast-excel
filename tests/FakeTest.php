<?php

namespace Rap2hpoutre\FastExcel\Tests;

use Rap2hpoutre\FastExcel\Facades\FastExcel as ExcelFacades;
use Rap2hpoutre\FastExcel\Fake\FastExcelFake;
use Rap2hpoutre\FastExcel\FastExcel;

class FakeTest extends TestCase
{
    protected function setUp(): void
    {
        parent::setUp();
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function testCanFakeAnExport()
    {
        $this->assertInstanceOf(FastExcel::class, $this->app->make('fastexcel'));

        ExcelFacades::fake();

        $this->assertInstanceOf(FastExcelFake::class, $this->app->make('fastexcel'));
    }

    public function testFakeDownload()
    {
        ExcelFacades::fake();

        $this->app->fastexcel->download('bar.xlsx');

        $this->app->fastexcel->assertDownloaded('bar.xlsx');
    }

    public function testFakeDownloadWithData()
    {
        ExcelFacades::fake();

        $original_collection = $this->collection();

        $this->app->fastexcel->data($original_collection)->download('bar.xlsx');

        $this->app->fastexcel->assertDownloaded('bar.xlsx', function ($data) use($original_collection) {
            if($data->toJson()== $original_collection->toJson()) {
                return true;
            }
            return false;
        });
    }

    public function testFakeDownloadCallback()
    {
        ExcelFacades::fake();

        $original_collection = $this->collection();

        $call = function(){
            return 'foo';
        };

        $this->app->fastexcel->data($original_collection)->download('bar.xlsx', $call);

        $this->app->fastexcel->assertDownloadedCallback('bar.xlsx', $call);
    }
}

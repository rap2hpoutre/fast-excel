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
}

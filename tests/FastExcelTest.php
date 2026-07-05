<?php

namespace Rap2hpoutre\FastExcel\Tests;

use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Style\Color;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Common\Exception\IOException;
use OpenSpout\Reader\XLSX\Options;
use Rap2hpoutre\FastExcel\FastExcel;
use Rap2hpoutre\FastExcel\SheetCollection;

/**
 * Class FastExcelTest.
 */
class FastExcelTest extends TestCase
{
    /**
     * @throws IOException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     */
    public function testExportXlsxWithDates()
    {
        $collection = collect([
            ['col1' => new \DateTimeImmutable('1980-09-18 00:00:00.000000')],
            ['col1' => new \DateTimeImmutable('2018-07-02 00:00:00.000000')],
        ]);

        $file = __DIR__.'/test-dates-export.xlsx';
        (new FastExcel(clone $collection))
            ->setColumnStyles([
                0 => (new Style())->setFormat('mm/dd/yyyy'),
            ])
            ->export($file);

        $this->assertEquals($collection, (new FastExcel())->import($file));
        unlink($file);
    }

    /**
     * A row value may be an OpenSpout Cell instance; it must be written
     * through as-is instead of being dropped or stringified.
     *
     * @throws IOException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     */
    public function testExportWithCellInstances()
    {
        $collection = collect([
            ['col1' => Cell::fromValue('hello'), 'col2' => 'world'],
            ['col1' => 'foo', 'col2' => Cell::fromValue('bar')],
        ]);

        $file = __DIR__.'/test-cell-instances.xlsx';
        (new FastExcel(clone $collection))->export($file);

        $this->assertEquals([
            ['col1' => 'hello', 'col2' => 'world'],
            ['col1' => 'foo', 'col2' => 'bar'],
        ], (new FastExcel())->import($file)->toArray());
        unlink($file);
    }

    /**
     * Cell instances must also survive the styled export path (setColumnStyles).
     *
     * @throws IOException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     */
    public function testExportWithCellInstancesAndColumnStyles()
    {
        $collection = collect([
            ['col1' => Cell::fromValue('hello'), 'col2' => 'world'],
        ]);

        $file = __DIR__.'/test-cell-instances-styled.xlsx';
        (new FastExcel(clone $collection))
            ->setColumnStyles([
                1 => (new Style())->setFontBold(),
            ])
            ->export($file);

        $this->assertEquals([
            ['col1' => 'hello', 'col2' => 'world'],
        ], (new FastExcel())->import($file)->toArray());
        unlink($file);
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     */
    public function testTransposeImport()
    {
        $collection = (new FastExcel())->transpose()->import(__DIR__.'/test1.xlsx');

        $this->assertEquals([
            'col1' => ['row1 col1', 'row2 col1', 'row3 col1'],
            'col2' => ['row1 col2', '', 'row3 col2'],
        ], $collection->toArray());
    }

    /**
     * @throws IOException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     */
    public function testTransposeExport()
    {
        $file = __DIR__.'/test-transpose-export.xlsx';
        (new FastExcel($this->collection()))->transpose()->export($file);

        // Each original column becomes a row; the header holds the original row indexes.
        $this->assertEquals([
            [0 => 'row1 col1', 1 => 'row2 col1', 2 => 'row3 col1'],
            [0 => 'row1 col2', 1 => '', 2 => 'row3 col2'],
        ], (new FastExcel())->import($file)->toArray());
        unlink($file);
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     */
    public function testImportXlsx()
    {
        $collection = (new FastExcel())->import(__DIR__.'/test1.xlsx');
        $this->assertEquals($this->collection(), $collection);
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     */
    public function testImportCsv()
    {
        $original_collection = $this->collection();

        $collection = (new FastExcel())->import(__DIR__.'/test2.csv');
        $this->assertEquals($original_collection, $collection);

        $collection = (new FastExcel())->configureCsv(';')->import(__DIR__.'/test1.csv');
        $this->assertEquals($original_collection, $collection);
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    private function export($file)
    {
        $original_collection = $this->collection();

        (new FastExcel(clone $original_collection))->export($file);
        $this->assertEquals($original_collection, (new FastExcel())->import($file));
        unlink($file);
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     */
    public function testExportXlsx()
    {
        $this->export(__DIR__.'/test2.xlsx');
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     */
    public function testExportCsv()
    {
        $this->export(__DIR__.'/test3.csv');
    }

    /**
     * Export data given as any Traversable (e.g. an Iterator), not only a
     * Collection, Generator or array.
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     */
    public function testExportFromTraversable()
    {
        // ArrayIterator is a Traversable that is neither a Generator nor a Collection.
        $data = new \ArrayIterator($this->collection()->toArray());

        $file = (new FastExcel($data))->export(__DIR__.'/test2.xlsx');
        $this->assertEquals($this->collection(), (new FastExcel())->import(__DIR__.'/test2.xlsx'));

        unlink($file);
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     */
    public function testExcelImportWithCallback()
    {
        $collection = (new FastExcel())->import(__DIR__.'/test1.xlsx', function ($value) {
            return [
                'test' => $value['col1'],
            ];
        });
        $this->assertEquals(
            collect([['test' => 'row1 col1'], ['test' => 'row2 col1'], ['test' => 'row3 col1']]),
            $collection
        );

        $collection = (new FastExcel())->import(__DIR__.'/test1.xlsx', function ($value) {
            return new Dumb($value['col1']);
        });
        $this->assertEquals(
            collect([new Dumb('row1 col1'), new Dumb('row2 col1'), new Dumb('row3 col1')]),
            $collection
        );
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testExcelExportWithCallback()
    {
        (new FastExcel(clone $this->collection()))->export(__DIR__.'/test2.xlsx', function ($value) {
            return [
                'test' => $value['col1'],
            ];
        });
        $this->assertEquals(
            collect([['test' => 'row1 col1'], ['test' => 'row2 col1'], ['test' => 'row3 col1']]),
            (new FastExcel())->import(__DIR__.'/test2.xlsx')
        );
        unlink(__DIR__.'/test2.xlsx');
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testExportMultiSheetXLSX()
    {
        $file = __DIR__.'/test_multi_sheets.xlsx';
        $sheets = new SheetCollection([clone $this->collection(), clone $this->collection()]);
        (new FastExcel($sheets))->export($file);
        $this->assertEquals($this->collection(), (new FastExcel())->import($file));
        unlink($file);
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testImportMultiSheetXLSX()
    {
        $collections = [
            collect([['test' => 'row1 col1'], ['test' => 'row2 col1'], ['test' => 'row3 col1']]),
            $this->collection(),
        ];
        $file = __DIR__.'/test_multi_sheets.xlsx';
        $sheets = new SheetCollection($collections);
        (new FastExcel($sheets))->export($file);

        $sheets = (new FastExcel())->importSheets($file);
        $this->assertInstanceOf(SheetCollection::class, $sheets);

        $this->assertEquals($collections[0], collect($sheets->first()));
        $this->assertEquals($collections[1], collect($sheets->all()[1]));

        unlink($file);
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testImportMultiSheetWithSheetNamesXLSX()
    {
        $collections = [
            'Sheet with name A' => collect([['test' => 'row1 col1'], ['test' => 'row2 col1'], ['test' => 'row3 col1']]),
            'Sheet with name B' => $this->collection(),
        ];
        $file = __DIR__.'/test_multi_sheets_with_sheets_names.xlsx';
        $sheets = new SheetCollection($collections);
        (new FastExcel($sheets))->export($file);

        $sheets = (new FastExcel())->withSheetsNames()->importSheets($file);
        $this->assertInstanceOf(SheetCollection::class, $sheets);

        $this->assertEquals($collections['Sheet with name A'], collect($sheets->get('Sheet with name A')));
        $this->assertEquals($collections['Sheet with name B'], collect($sheets->get('Sheet with name B')));

        unlink($file);
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testExportWithHeaderStyle()
    {
        $original_collection = $this->collection();

        $style = new Style();
        $style->setFontBold();
        $style->setFontSize(15);
        $style->setFontColor(Color::BLUE);
        $style->setShouldWrapText();
        $style->setBackgroundColor(Color::YELLOW);
        $file = __DIR__.'/test-header-style.xlsx';
        (new FastExcel(clone $original_collection))
            ->headerStyle($style)
            ->export($file);
        $this->assertEquals($original_collection, (new FastExcel())->import($file));

        unlink($file);
    }

    /**
     * Issue #367: each header column can be styled individually via
     * setHeaderColumnStyles(), the header-row counterpart of setColumnStyles().
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testExportWithHeaderColumnStyles()
    {
        $original_collection = $this->collection();

        $file = __DIR__.'/test-header-column-styles.xlsx';
        (new FastExcel(clone $original_collection))
            ->setHeaderColumnStyles([
                0 => (new Style())->setBackgroundColor(Color::YELLOW),
                1 => (new Style())->setFontColor(Color::BLUE),
            ])
            ->export($file);

        // The reader does not expose cell styles, so assert the styled export
        // still round-trips the data (same approach as the other style tests).
        $this->assertEquals($original_collection, (new FastExcel())->import($file));

        unlink($file);
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testConfigureWriterUsingWithCustomCsvWriter()
    {
        $usedCustomWriter = false;
        $original_collection = $this->collection();
        $file = __DIR__.'/test-custom-writer.csv';

        (new FastExcel(clone $original_collection))
            ->configureWriterUsing(function ($options, $extension) use (&$usedCustomWriter) {
                $this->assertEquals('csv', $extension);
                $this->assertInstanceOf(\OpenSpout\Writer\CSV\Options::class, $options);
                $usedCustomWriter = true;

                return new \OpenSpout\Writer\CSV\Writer($options);
            })
            ->export($file);

        $this->assertTrue($usedCustomWriter);
        $this->assertEquals($original_collection, (new FastExcel())->import($file));
        unlink($file);
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testConfigureWriterUsingFallbackToDefaultWriter()
    {
        $original_collection = $this->collection();
        $file = __DIR__.'/test-writer-fallback.csv';

        (new FastExcel(clone $original_collection))
            ->configureWriterUsing(function ($options, $extension) {
                return null;
            })
            ->export($file);

        $this->assertEquals($original_collection, (new FastExcel())->import($file));
        unlink($file);
    }

    public function testImportXlsxWithCustomDateOption()
    {
        // Default options, dates will end parsed
        $collection = (new FastExcel())->import(__DIR__.'/test-dates.xlsx');

        $this->assertEquals(collect([
            ['col1' => new \DateTimeImmutable('2022-01-02 00:00:00.000000')],
            ['col1' => new \DateTimeImmutable('2022-01-03 00:00:00.000000')],
        ]), $collection);

        $collection = (new FastExcel())
            ->configureOptionsUsing(function ($options) {
                if ($options instanceof Options) {
                    $options->SHOULD_FORMAT_DATES = true;
                }
            })
            ->import(__DIR__.'/test-dates.xlsx');

        $this->assertEquals(collect([
            ['col1' => '1/2/2022'],
            ['col1' => '1/3/2022'],
        ]), $collection);
    }
}

<?php

namespace Rap2hpoutre\FastExcel\Tests;

use Illuminate\Support\Collection;
use OpenSpout\Common\Entity\Style\Style;
use Rap2hpoutre\FastExcel\FastExcel;
use Rap2hpoutre\FastExcel\SheetCollection;
use ZipArchive;

/**
 * Class IssuesTest.
 */
class IssuesTest extends TestCase
{
    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
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
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     */
    public function testIssue18()
    {
        $collection = (new FastExcel())->import(__DIR__.'/test18.csv');
        $this->assertInstanceOf(Collection::class, $collection);
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testIssue20()
    {
        chdir(__DIR__);
        $path = (new FastExcel($this->collection()))->export('test2.xlsx');
        $this->assertEquals(__DIR__.DIRECTORY_SEPARATOR.'test2.xlsx', $path);
        unlink($path);
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     */
    public function testIssue19()
    {
        chdir(__DIR__);
        $path = (new FastExcel(collect([['a' => 1, 'b' => 'n', 'c' => 1.32, 'd' => []]])))->export('test2.xlsx');
        $this->assertEquals(collect([['a' => '1', 'b' => 'n', 'c' => '1.32']]), (new FastExcel())->import(__DIR__.'/test2.xlsx'));
        unlink($path);
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testIssue26()
    {
        chdir(__DIR__);
        foreach ([[[]], [null]] as $value) {
            $path = (new FastExcel($value))->export('test2.xlsx');
            $this->assertEquals(collect([]), (new FastExcel())->import(__DIR__.'/test2.xlsx'));
            unlink($path);
        }
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
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
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testIssue40()
    {
        $col = new SheetCollection(['1st Sheet' => $this->collection(), '2nd Sheet' => $this->collection()]);
        (new FastExcel($col))->export(__DIR__.'/test2.xlsx');

        $options = new \OpenSpout\Reader\XLSX\Options();
        $reader = new \OpenSpout\Reader\XLSX\Reader($options);
        $reader->open(__DIR__.'/test2.xlsx');
        foreach ($reader->getSheetIterator() as $key => $sheet) {
            $this->assertEquals($sheet->getName(), $key === 2 ? '2nd Sheet' : '1st Sheet');
        }
        $reader->close();
        unlink(__DIR__.'/test2.xlsx');
    }

    public function testIssue72()
    {
        $collection = (new FastExcel())->import(__DIR__.'/test72.xlsx');
        $this->assertInstanceOf(Collection::class, $collection);
    }

    public function testIssue93()
    {
        (new FastExcel($this->collection()))->export(__DIR__.'/猫.xlsx');
        $this->assertTrue(file_exists(__DIR__.'/猫.xlsx'));
        unlink(__DIR__.'/猫.xlsx');
    }

    public function testIssue86()
    {
        $users = (new FastExcel())->withoutHeaders()->import(__DIR__.'/test1.xlsx', function ($line) {
            return $line;
        });
        $this->assertCount(4, $users);
        $this->assertEquals($users[0], ['col1', 'col2']);
    }

    public function testIssue104()
    {
        $users = (new FastExcel())->import(__DIR__.'/test104.xlsx', function ($line) {
            return $line;
        });
        $this->assertCount(3, $users);
        $this->assertEquals($users[0], [
            'Name'     => 'joe',
            'Email'    => 'joe@gmail.com',
            'Password' => 'asdadasdasdasdasd',
        ]);
    }

    public function testIssue310()
    {
        $original_collection = $this->collection();
        $delimiter = ';';
        $file = 'issue_310.csv';

        (new FastExcel(clone $original_collection))
            ->configureCsv($delimiter)
            ->export($file);

        $this->assertEquals(
            $original_collection,
            (new FastExcel())
                ->configureCsv($delimiter)
                ->import($file)
        );

        unlink($file);
    }

    /**
     * Issue #185: importing an uploaded CSV/ODS fails because the format is
     * guessed from the path extension, and uploaded files live under an
     * extension-less temporary path (e.g. /tmp/phpXXXX).
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     */
    public function testIssue185()
    {
        // A CSV stored under an extension-less path must be detected by content.
        $csvNoExt = tempnam(sys_get_temp_dir(), 'php');
        copy(__DIR__.'/test2.csv', $csvNoExt);
        $this->assertCount(3, (new FastExcel())->import($csvNoExt));

        // An XLSX stored under an extension-less path must still work.
        $xlsxNoExt = tempnam(sys_get_temp_dir(), 'php');
        copy(__DIR__.'/test104.xlsx', $xlsxNoExt);
        $this->assertCount(3, (new FastExcel())->import($xlsxNoExt));

        // An ODS stored under an extension-less path is detected by sniffing the
        // zip mimetype entry (ODS and XLSX are both zip archives).
        $ods = __DIR__.'/issue_185.ods';
        (new FastExcel($this->collection()))->export($ods);
        $odsNoExt = tempnam(sys_get_temp_dir(), 'php');
        copy($ods, $odsNoExt);
        $this->assertEquals($this->collection(), (new FastExcel())->import($odsNoExt));

        unlink($csvNoExt);
        unlink($xlsxNoExt);
        unlink($odsNoExt);
        unlink($ods);
    }

    /**
     * Issue #185: importing directly from an UploadedFile-like object should
     * resolve the format from its original extension / client mime type.
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     */
    public function testIssue185UploadedFile()
    {
        $makeUploadedFile = function (string $source, string $originalName, string $mime) {
            $tmp = tempnam(sys_get_temp_dir(), 'php');
            copy($source, $tmp);

            // Duck-typed stand-in for Illuminate/Symfony UploadedFile so the
            // test does not require illuminate/http to be installed.
            return new class($tmp, $originalName, $mime) {
                public function __construct(private string $tmp, private string $name, private string $mime)
                {
                }

                public function getClientOriginalExtension(): string
                {
                    return pathinfo($this->name, PATHINFO_EXTENSION);
                }

                public function getClientMimeType(): string
                {
                    return $this->mime;
                }

                public function getPathname(): string
                {
                    return $this->tmp;
                }
            };
        };

        // Resolved from the original extension.
        $csvUpload = $makeUploadedFile(__DIR__.'/test2.csv', 'data.csv', 'text/csv');
        $this->assertCount(3, (new FastExcel())->import($csvUpload));

        // Resolved from the client mime type when the original name has no extension.
        $csvUploadNoExt = $makeUploadedFile(__DIR__.'/test2.csv', 'data', 'text/csv');
        $this->assertCount(3, (new FastExcel())->import($csvUploadNoExt));

        // ODS uploaded file, resolved from its original extension.
        $ods = __DIR__.'/issue_185_upload.ods';
        (new FastExcel($this->collection()))->export($ods);
        $odsUpload = $makeUploadedFile($ods, 'book.ods', 'application/vnd.oasis.opendocument.spreadsheet');
        $this->assertEquals($this->collection(), (new FastExcel())->import($odsUpload));

        unlink($csvUpload->getPathname());
        unlink($csvUploadNoExt->getPathname());
        unlink($odsUpload->getPathname());
        unlink($ods);
    }

    /**
     * Issue #244: exporting multiple sheets to a single-sheet format (CSV) must
     * fail with a clear message instead of a cryptic
     * "Call to undefined method ...\CSV\Writer::getCurrentSheet()" fatal.
     */
    public function testIssue244()
    {
        $sheets = new SheetCollection([
            'articles' => collect([['col1' => 'a1', 'col2' => 'a2']]),
            'blogs'    => collect([['col1' => 'b1', 'col2' => 'b2']]),
        ]);

        $this->expectException(\InvalidArgumentException::class);
        $this->expectExceptionMessage('does not support multiple sheets');

        (new FastExcel($sheets))->export(__DIR__.'/issue_244.csv');
    }

    /**
     * A single-sheet SheetCollection must still export fine to CSV (the sheet
     * name is simply ignored, since CSV has no sheets).
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     */
    public function testIssue244SingleSheetCsv()
    {
        $sheets = new SheetCollection(['only' => $this->collection()]);

        $file = __DIR__.'/issue_244_single.csv';
        (new FastExcel($sheets))->export($file);

        $this->assertEquals($this->collection(), (new FastExcel())->import($file));

        unlink($file);
    }

    /**
     * Issue #162: importing a large file with a callback that returns null must
     * process every row without accumulating them, so memory stays flat. A
     * callback that returns a value is collected instead, which grows memory.
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     */
    public function testIssue162()
    {
        $file = __DIR__.'/issue_162.xlsx';
        (new FastExcel($this->collection()))->export($file);

        $seen = 0;
        $result = (new FastExcel())->import($file, function ($line) use (&$seen) {
            $this->assertArrayHasKey('col1', $line);
            $seen++;

            return null; // stream: do not accumulate the row
        });

        $this->assertSame(3, $seen);    // every row was processed
        $this->assertCount(0, $result); // nothing was kept in memory

        unlink($file);
    }

    /**
     * Issues #312 and #259: a header row with duplicate or empty names must not
     * lose data. Duplicates keep the first occurrence and get a numeric suffix,
     * empty headers get a positional name, so no column collides in the result.
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     */
    public function testIssue312()
    {
        // Header row with a duplicate ("Name") and an empty column.
        $file = __DIR__.'/issue_312.csv';
        file_put_contents($file, "Name,Name,,Age\nJoe,Smith,x,30\nJane,Doe,y,25\n");

        $rows = (new FastExcel())->import($file);

        $first = $rows->first();
        $this->assertSame(['Name', 'Name_2', 'column_3', 'Age'], array_keys($first));
        $this->assertSame('Joe', $first['Name']);
        $this->assertSame('Smith', $first['Name_2']);   // would have been lost before
        $this->assertSame('x', $first['column_3']);
        $this->assertSame('30', $first['Age']);
        $this->assertCount(2, $rows);

        unlink($file);
    }

    /**
     * Issue #193: the string-vs-number cell type on export must be configurable.
     * By default numbers stay numeric; stringValues() forces every scalar to a
     * text cell (preserving leading zeros / long IDs); setColumnFormat() lets
     * each column opt in or out and takes precedence over stringValues().
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testIssue193()
    {
        $file = __DIR__.'/issue_193.xlsx';
        $row = ['id' => 7, 'phone' => '0660123', 'price' => 12.5];

        // Default: numbers stay numeric, strings stay strings (leading zero kept).
        (new FastExcel(collect([$row])))->export($file);
        $default = (new FastExcel())->import($file)->first();
        $this->assertSame(7, $default['id']);
        $this->assertSame('0660123', $default['phone']);
        $this->assertSame(12.5, $default['price']);

        // stringValues(): every scalar becomes a text cell.
        (new FastExcel(collect([$row])))->stringValues()->export($file);
        $strings = (new FastExcel())->import($file)->first();
        $this->assertSame('7', $strings['id']);
        $this->assertSame('0660123', $strings['phone']);
        $this->assertSame('12.5', $strings['price']);

        // setColumnFormat() overrides per column (and wins over stringValues()).
        (new FastExcel(collect([$row])))
            ->stringValues()
            ->setColumnFormat(['phone' => 'number', 'price' => 'number'])
            ->export($file);
        $mixed = (new FastExcel())->import($file)->first();
        $this->assertSame('7', $mixed['id']);          // still string (global flag)
        $this->assertSame(660123, $mixed['phone']);    // forced numeric
        $this->assertSame(12.5, $mixed['price']);      // forced numeric

        unlink($file);
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     *
     * @see https://github.com/rap2hpoutre/fast-excel/issues/372
     */
    public function testIssue372()
    {
        $file = __DIR__.'/issue372.xlsx';

        (new FastExcel(new SheetCollection([
            'A' => collect([['a' => 'b']]),
            'B' => collect(),
        ])))->export($file);

        $reader = new \OpenSpout\Reader\XLSX\Reader(new \OpenSpout\Reader\XLSX\Options());
        $reader->open($file);

        $sheets = [];
        foreach ($reader->getSheetIterator() as $sheet) {
            $rowCount = 0;
            foreach ($sheet->getRowIterator() as $row) {
                $rowCount++;
            }
            $sheets[$sheet->getName()] = $rowCount;
        }

        $reader->close();

        $this->assertArrayHasKey('A', $sheets);
        $this->assertArrayHasKey('B', $sheets);
        $this->assertSame(2, $sheets['A']);
        $this->assertSame(1, $sheets['B']);

        unlink($file);
    }

    /**
     * Issue #252: exporting multiple sheets where each sheet is a Generator (or
     * any other Traversable, not only a Collection). Each sheet value goes
     * through the same Traversable dispatch as a single-sheet export, so headers
     * and rows are written per sheet.
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     *
     * @see https://github.com/rap2hpoutre/fast-excel/issues/252
     */
    public function testIssue252()
    {
        $file = __DIR__.'/issue252.xlsx';

        $generator = function () {
            yield ['name' => 'Alice', 'age' => 30];
            yield ['name' => 'Bob', 'age' => 25];
        };

        (new FastExcel(new SheetCollection([
            'People'  => $generator(),                                   // Generator sheet
            'Numbers' => new \ArrayIterator([['n' => 1], ['n' => 2]]),   // non-Generator Traversable
        ])))->export($file);

        $reader = new \OpenSpout\Reader\XLSX\Reader(new \OpenSpout\Reader\XLSX\Options());
        $reader->open($file);

        $sheets = [];
        foreach ($reader->getSheetIterator() as $sheet) {
            $rows = [];
            foreach ($sheet->getRowIterator() as $row) {
                $rows[] = array_map(fn ($cell) => $cell->getValue(), $row->getCells());
            }
            $sheets[$sheet->getName()] = $rows;
        }

        $reader->close();

        // Both Traversable sheets are written with a header row + their data.
        $this->assertSame(['name', 'age'], $sheets['People'][0]);
        $this->assertSame('Alice', $sheets['People'][1][0]);
        $this->assertSame('Bob', $sheets['People'][2][0]);
        $this->assertCount(3, $sheets['People']);

        $this->assertSame(['n'], $sheets['Numbers'][0]);
        $this->assertCount(3, $sheets['Numbers']);

        unlink($file);
    }

    /**
     * Issue #241: when the row callback throws, import() (and importSheets())
     * must still close the reader. Previously close() ran after the loop, so an
     * exception in the callback leaked the open reader (and XLSX temp files).
     * Wrapping the loop in try/finally releases the reader; a follow-up import
     * of the same file must then succeed.
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     */
    public function testIssue241ReaderClosedWhenCallbackThrows()
    {
        $file = __DIR__.'/test1.xlsx';

        // The callback throws mid-import; the exception must propagate.
        try {
            (new FastExcel())->import($file, function () {
                throw new \RuntimeException('boom');
            });
            $this->fail('Expected RuntimeException was not thrown.');
        } catch (\RuntimeException $e) {
            $this->assertSame('boom', $e->getMessage());
        }

        // Regression guard: the reader was released, so re-importing the same
        // file works and returns the expected rows.
        $rows = (new FastExcel())->import($file);
        $this->assertInstanceOf(Collection::class, $rows);
        $this->assertEquals($this->collection(), $rows);
    }

    /**
     * A string value starting with "=" must be written as literal text (an
     * inline string) when escapeFormulas() is enabled, instead of a live
     * formula cell. This protects against CSV/formula injection and the
     * "corrupt file" errors reported in issue #356.
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testIssue356()
    {
        $file = __DIR__.'/issue356.xlsx';
        $rows = new Collection([['formula' => '=1+2']]);

        // With escapeFormulas(): the value is stored as literal inline-string
        // text, so no live formula element ("<f>...</f>") is written.
        (new FastExcel(clone $rows))->escapeFormulas()->export($file);
        $xml = $this->readXlsxSheetXml($file);
        $this->assertStringContainsString('<t>=1+2</t>', $xml);
        $this->assertStringNotContainsString('<f>', $xml);
        unlink($file);

        // Without escapeFormulas(): the same value is emitted as a live
        // formula cell (the pre-existing, unsafe behavior).
        (new FastExcel(clone $rows))->export($file);
        $xml = $this->readXlsxSheetXml($file);
        $this->assertStringContainsString('<f>1+2</f>', $xml);
        unlink($file);
    }

    /**
     * Issue #420: exporting a sheet with right-to-left content (e.g. Arabic)
     * had no way to set the sheet's reading direction. openspout's
     * SheetView already supports this via setRightToLeft(); rightToLeft()
     * exposes it, applied once the writer is opened (the sheet does not exist
     * yet inside configureWriterUsing()).
     *
     * @see https://github.com/rap2hpoutre/fast-excel/issues/420
     */
    public function testIssue420()
    {
        $file = __DIR__.'/issue420.xlsx';

        (new FastExcel(collect([['name' => 'محمد', 'value' => 100]])))
            ->rightToLeft()
            ->export($file);

        $zip = new ZipArchive();
        $zip->open($file);
        $sheetXml = $zip->getFromName('xl/worksheets/sheet1.xml');
        $zip->close();

        $this->assertStringContainsString('rightToLeft="true"', $sheetXml);

        unlink($file);
    }

    /**
     * rightToLeft() must apply to every sheet in a SheetCollection export, not
     * only the first — SheetView has to be re-applied after each
     * addNewSheetAndMakeItCurrent() call.
     */
    public function testIssue420MultiSheet()
    {
        $file = __DIR__.'/issue420_multisheet.xlsx';

        (new FastExcel(new SheetCollection([
            'Sheet A' => collect([['name' => 'محمد']]),
            'Sheet B' => collect([['name' => 'محمود']]),
            'Sheet C' => collect([['name' => 'احمد']]),
            'Sheet D' => collect([['name' => 'عماد']]),
        ])))->rightToLeft()->export($file);

        $zip = new ZipArchive();
        $zip->open($file);
        $sheet1 = $zip->getFromName('xl/worksheets/sheet1.xml');
        $sheet2 = $zip->getFromName('xl/worksheets/sheet2.xml');
        $sheet3 = $zip->getFromName('xl/worksheets/sheet3.xml');
        $sheet4 = $zip->getFromName('xl/worksheets/sheet4.xml');
        $zip->close();

        $this->assertStringContainsString('rightToLeft="true"', $sheet1);
        $this->assertStringContainsString('rightToLeft="true"', $sheet2);
        $this->assertStringContainsString('rightToLeft="true"', $sheet3);
        $this->assertStringContainsString('rightToLeft="true"', $sheet4);

        unlink($file);
    }

    /**
     * Read the first worksheet XML out of an XLSX file (a zip archive).
     */
    private function readXlsxSheetXml(string $file): string
    {
        $zip = new ZipArchive();
        $zip->open($file);
        $xml = $zip->getFromName('xl/worksheets/sheet1.xml');
        $zip->close();

        return $xml;
    }

    /**
     * Issue #349: startRow() alone also takes the headers from the start row, so
     * reading a file from row N loses the real headers (row N's data is used as
     * the header names). headerRow() decouples the two.
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testIssue349HeaderRowSeparateFromStartRow()
    {
        // Row 1 holds the headers; rows 2..6 hold data 1..5.
        $file = __DIR__.'/issue349.xlsx';
        (new FastExcel($this->issue349Rows()))->export($file);

        // Headers from row 1, data from row 4 (i.e. data 3, 4 and 5).
        $rows = (new FastExcel())->headerRow(1)->startRow(4)->import($file);
        $this->assertSame(['name', 'qty'], array_keys($rows->first()));
        $this->assertSame(
            [['name' => 'c', 'qty' => '3'], ['name' => 'd', 'qty' => '4'], ['name' => 'e', 'qty' => '5']],
            $rows->toArray()
        );

        unlink($file);
    }

    /**
     * headerRow() is what makes chunked imports work: each run reads the same
     * headers but a different slice of data rows, with no overlap.
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testIssue349ChunkedImportWithLimitRows()
    {
        $file = __DIR__.'/issue349_chunks.xlsx';
        (new FastExcel($this->issue349Rows()))->export($file);

        $chunks = [];
        // Data starts on row 2, so chunk k starts at row 2 + (k * size).
        foreach ([2, 4, 6] as $start) {
            $chunks[] = (new FastExcel())->headerRow(1)->startRow($start)->limitRows(2)->import($file)->toArray();
        }

        $this->assertSame([['name' => 'a', 'qty' => '1'], ['name' => 'b', 'qty' => '2']], $chunks[0]);
        $this->assertSame([['name' => 'c', 'qty' => '3'], ['name' => 'd', 'qty' => '4']], $chunks[1]);
        $this->assertSame([['name' => 'e', 'qty' => '5']], $chunks[2]);

        // Every data row appears exactly once across the chunks.
        $this->assertCount(5, array_merge(...$chunks));

        // The lazy path honours it identically.
        $lazy = (new FastExcel())->headerRow(1)->startRow(4)->limitRows(2)->importLazy($file)->all();
        $this->assertSame($chunks[1], $lazy);

        unlink($file);
    }

    /**
     * Without headerRow(), startRow() keeps its historical meaning: the start
     * row is also the header row. This must not change.
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testIssue349StartRowBackwardCompatibility()
    {
        $file = __DIR__.'/issue349_bc.xlsx';
        (new FastExcel($this->issue349Rows()))->export($file);

        // Row 3 is data ('b', '2') and is still consumed as the header row.
        $rows = (new FastExcel())->startRow(3)->import($file);
        // '2' becomes int 2: PHP casts numeric-string array keys.
        $this->assertSame(['b', 2], array_keys($rows->first()));
        $this->assertSame([['b' => 'c', '2' => '3'], ['b' => 'd', '2' => '4'], ['b' => 'e', '2' => '5']], $rows->toArray());

        // headerRow(null) explicitly restores that behaviour too.
        $rows = (new FastExcel())->headerRow(null)->startRow(3)->import($file);
        // '2' becomes int 2: PHP casts numeric-string array keys.
        $this->assertSame(['b', 2], array_keys($rows->first()));

        unlink($file);
    }

    /**
     * Five data rows under a single header row, used by the #349 tests.
     */
    private function issue349Rows(): Collection
    {
        return collect([
            ['name' => 'a', 'qty' => '1'],
            ['name' => 'b', 'qty' => '2'],
            ['name' => 'c', 'qty' => '3'],
            ['name' => 'd', 'qty' => '4'],
            ['name' => 'e', 'qty' => '5'],
        ]);
    }

    /**
     * Issue #213: sizing columns to their content had to be done by hand with
     * setColumnWidth(). autoSizeColumns() measures each column while the rows
     * stream past and writes the widths just before the file is finalized.
     *
     * @see https://github.com/rap2hpoutre/fast-excel/issues/213
     */
    public function testIssue213()
    {
        $file = __DIR__.'/issue213.xlsx';

        (new FastExcel(collect([
            ['id' => 1, 'name' => 'Bob', 'description' => 'a fairly long description value'],
            ['id' => 12345, 'name' => 'Alexandra', 'description' => 'short'],
        ])))->autoSizeColumns()->export($file);

        $cols = $this->columnsFragment($file);

        // Widest value plus two characters of padding: "12345" -> 7,
        // "Alexandra" -> 11, "a fairly long description value" -> 33.
        $this->assertStringContainsString('<col min="1" max="1" width="7"', $cols);
        $this->assertStringContainsString('<col min="2" max="2" width="11"', $cols);
        $this->assertStringContainsString('<col min="3" max="3" width="33"', $cols);

        unlink($file);
    }

    /**
     * Auto-sizing is opt-in: an export that does not ask for it must be byte
     * for byte what it was before, with no <cols> fragment at all.
     */
    public function testIssue213OffByDefault()
    {
        $file = __DIR__.'/issue213_default.xlsx';

        (new FastExcel(collect([['name' => 'a value long enough to notice']])))->export($file);

        $this->assertSame('', $this->columnsFragment($file));

        unlink($file);
    }

    /**
     * A single very long cell must not produce an unusable column, so widths
     * are clamped; the bound is configurable and Excel's own limit is 255.
     */
    public function testIssue213ClampsWidthToBounds()
    {
        $file = __DIR__.'/issue213_clamp.xlsx';
        $rows = collect([['big' => str_repeat('x', 500), 'tiny' => 'a']]);

        (new FastExcel(clone $rows))->autoSizeColumns()->export($file);
        $this->assertStringContainsString('<col min="1" max="1" width="60"', $this->columnsFragment($file));

        (new FastExcel(clone $rows))->autoSizeColumns(true, 25.0, 15.0)->export($file);
        $cols = $this->columnsFragment($file);
        $this->assertStringContainsString('<col min="1" max="1" width="25"', $cols);
        $this->assertStringContainsString('<col min="2" max="2" width="15"', $cols);

        unlink($file);
    }

    public function testIssue213RejectsInvalidBounds()
    {
        $this->expectException(\InvalidArgumentException::class);
        (new FastExcel(collect([['a' => 'b']])))->autoSizeColumns(true, 256.0);
    }

    public function testIssue213RejectsMinWiderThanMax()
    {
        $this->expectException(\InvalidArgumentException::class);
        (new FastExcel(collect([['a' => 'b']])))->autoSizeColumns(true, 20.0, 30.0);
    }

    /**
     * Widths are held on the sheet, not the workbook options, so each sheet of
     * a SheetCollection is sized from its own content.
     */
    public function testIssue213MultiSheetSizesEachSheetSeparately()
    {
        $file = __DIR__.'/issue213_multisheet.xlsx';

        (new FastExcel(new SheetCollection([
            'Narrow' => collect([['col' => 'ab']]),
            'Wide'   => collect([['col' => 'a considerably wider value here']]),
        ])))->autoSizeColumns()->export($file);

        $this->assertStringContainsString(
            '<col min="1" max="1" width="5"',
            $this->columnsFragment($file, 'xl/worksheets/sheet1.xml')
        );
        $this->assertStringContainsString(
            '<col min="1" max="1" width="33"',
            $this->columnsFragment($file, 'xl/worksheets/sheet2.xml')
        );

        unlink($file);
    }

    /**
     * The whole point is that this survives streaming: a generator export must
     * be sized without the rows ever being collected.
     */
    public function testIssue213WorksWhileStreaming()
    {
        $file = __DIR__.'/issue213_generator.xlsx';

        $generator = (function () {
            foreach ([['n' => 'a streamed value that is long'], ['n' => 'b']] as $index => $row) {
                yield $index => $row;
            }
        })();

        (new FastExcel($generator))->autoSizeColumns()->export($file);

        $this->assertStringContainsString('<col min="1" max="1" width="31"', $this->columnsFragment($file));

        unlink($file);
    }

    /**
     * A wrapped column keeps the width it was given: Excel's AutoFit grows the
     * row height there rather than the column, and widening it would undo the
     * wrap the caller asked for.
     */
    public function testIssue213LeavesWrappedColumnsAlone()
    {
        $file = __DIR__.'/issue213_wrap.xlsx';

        (new FastExcel(collect([['c' => str_repeat('y', 80)]])))
            ->withoutHeaders()
            ->rowsStyle((new Style())->setShouldWrapText())
            ->autoSizeColumns()
            ->export($file);

        $this->assertSame('', $this->columnsFragment($file));

        unlink($file);
    }

    /**
     * csv has no notion of column width and ODS stores widths per workbook in
     * points rather than per sheet in characters, so both are left untouched
     * instead of being given a meaningless width.
     */
    public function testIssue213IgnoredByFormatsWithoutSheetWidths()
    {
        foreach (['csv', 'ods'] as $extension) {
            $file = __DIR__.'/issue213_ignored.'.$extension;

            $rows = collect([['name' => 'a value long enough to notice']]);
            (new FastExcel(clone $rows))->autoSizeColumns()->export($file);

            $this->assertEquals($rows, (new FastExcel())->import($file));

            unlink($file);
        }
    }

    /**
     * Read the <cols> fragment of a worksheet, or '' when the sheet has none.
     */
    private function columnsFragment(string $file, string $sheet = 'xl/worksheets/sheet1.xml'): string
    {
        $zip = new ZipArchive();
        $zip->open($file);
        $xml = $zip->getFromName($sheet);
        $zip->close();

        return preg_match('#<cols>.*?</cols>#s', $xml, $matches) === 1 ? $matches[0] : '';
    }
}

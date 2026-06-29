<?php

namespace Rap2hpoutre\FastExcel\Tests;

use Illuminate\Support\Collection;
use Rap2hpoutre\FastExcel\FastExcel;
use Rap2hpoutre\FastExcel\SheetCollection;

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
}

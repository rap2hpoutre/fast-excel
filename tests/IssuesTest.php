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

        unlink($csvNoExt);
        unlink($xlsxNoExt);
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

        unlink($csvUpload->getPathname());
        unlink($csvUploadNoExt->getPathname());
    }
}

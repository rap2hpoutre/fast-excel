<?php

namespace Rap2hpoutre\FastExcel;

use Box\Spout\Reader\ReaderFactory;
use Illuminate\Support\Collection;

/**
 * Trait Importable.
 *
 * @property bool $with_header
 */
trait Importable
{
    /**
     * @var int
     */
    private $sheet_number = 1;

    /**
     * @param string $path
     *
     * @return mixed
     */
    abstract protected function getType($path);

    /**
     * @param \Box\Spout\Reader\ReaderInterface|\Box\Spout\Writer\WriterInterface $reader_or_writer
     *
     * @return mixed
     */
    abstract protected function setOptions(&$reader_or_writer);

    /**
     * @param string $path
     * @param callable|null $rowCallback
     * @param callable|null $filterSheetsCallback: boolean
     * @param callable|null $sheetsCallback
     *
     * @return Collection of Collections (= Sheets) of their respective Row objects
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function import($path, callable $rowCallback = null, callable $filterSheetsCallback = null, callable $sheetsCallback = null)
    {
        /**
         * TODO: consider dropping the hardcoded $sheet_number variable.
         */
        $sheets = $this->importSheets($path, $sheetsCallback)
            ->filter(function($sheet, $key) use ($filterSheetsCallback)
            {
                if(!is_null($filterSheetsCallback))
                {
                    return $filterSheetsCallback($sheet, $key);
                }
                return $this->sheet_number == $key;
            })
            ->map(function(&$sheet, $key) use ($rowCallback)
            {
                $headers = [];
                $collection = [];
                $count_header = 0;

                if($this->with_header)
                {
                    foreach($sheet->getRowIterator() as $k => $row)
                    {
                        if($k == 1)
                        {
                            $headers = $row;
                            $count_header = count($headers);
                            continue;
                        }
                        if($count_header > $count_row = count($row))
                        {
                            $row = array_merge($row, array_fill(0, $count_header - $count_row, null));
                        }
                        $collection[] = $rowCallback ? $rowCallback(array_combine($headers, $row)) : array_combine($headers, $row);
                    }
                }
                else
                {
                    foreach($sheet->getRowIterator() as $row)
                    {
                        $collection[] = $row;
                    }
                }

                return collect($collection);
            });

        return $sheets;
    }

    /**
     * @param string        $path
     * @param callable|null $callback
     *
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     *
     * @return Collection of \Box\Spout\Reader\SheetInterface instances
     */
    public function importSheets($path, callable $callback = null)
    {
        $sheets = collect();

        $reader = $this->openReader($path);

        foreach ($reader->getSheetIterator() as $key => $sheet) {
            if($callback)
            {
                $sheets->put($key, $callback($key, $sheet));
                continue;
            }
            $sheets->put($key, $sheet);
        }

        /**
         * Attempting to close the reader within this extracted method will later cause the application to crash
         * with an ErrorException once the sheets' rows are supposed to be iterated.
         *
         *      Undefined property: Box\Spout\Reader\XLSX\Helper\SharedStringsCaching\InMemoryStrategy::$inMemoryCache
         *      at /var/www/vendor/box/spout/src/Spout/Reader/XLSX/Helper/SharedStringsCaching/InMemoryStrategy.php:67
         *
         * This happens because when we close the reader, the static ReaderFactory class unsets that variable after
         * closing it. Unfortunately, there is no way to somehow globally retain this Reader instance unless it's being
         * refactored into FastExcel.class (which it should, IMO). Hence, we're producing a potential memory leak here
         * but we should not be copy-pasting the openReader() code everywhere.
         *
         * TODO: find a suitable place for this:
         * $reader->close();
         */

        return $sheets;
    }


    /**
     * @param $path
     * @return \Box\Spout\Reader\ReaderInterface $reader
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     */
    private function openReader($path)
    {
        $reader = ReaderFactory::create($this->getType($path));
        $this->setOptions($reader);

        $reader->open($path);

        return $reader;
    }
}

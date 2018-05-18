<?php

namespace Rap2hpoutre\FastExcel;

use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Reader\ReaderInterface;
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
        $reader = $this->openReader($path);
        /**
         * TODO: consider dropping the hardcoded $sheet_number variable.
         */
        $sheets = $this->importSheets($path, $sheetsCallback, $reader)
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

        $reader->close();

        return $sheets;
    }

    /**
     * @param string $path
     * @param callable|null $callback
     * @param ReaderInterface $reader
     *
     * @return Collection of \Box\Spout\Reader\SheetInterface instances
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function importSheets($path, callable $callback = null, ReaderInterface $reader = null)
    {
        $sheets = collect();

        $tempReader = $reader ?: $this->openReader($path);

        foreach ($tempReader->getSheetIterator() as $key => $sheet) {
            if($callback)
            {
                $sheets->put($key, $callback($key, $sheet));
                continue;
            }
            $sheets->put($key, $sheet);
        }

        /**
         * If the following condition is met, the reader was not passed but instantiated locally and needs to be closed.
         */
        if(is_null($reader))
        {
            $tempReader->close();
        }

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

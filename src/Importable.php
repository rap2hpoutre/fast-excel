<?php

namespace Rap2hpoutre\FastExcel;

use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Reader\SheetInterface;
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
     * @param string        $path
     * @param callable|null $callback
     *
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     *
     * @return Collection
     */
    public function import($path, callable $callback = null)
    {
        $reader = ReaderFactory::create($this->getType($path));
        $this->setOptions($reader);
        /* @var \Box\Spout\Reader\ReaderInterface $reader */
        $reader->open($path);

        foreach ($reader->getSheetIterator() as $key => $sheet) {
            if ($this->sheet_number != $key) {
                continue;
            }
            $collection = $this->importSheet($sheet, $callback);
        }
        $reader->close();

        return collect($collection ?? []);
    }

    /**
     * @param SheetInterface $sheet
     * @param callable|null  $callback
     *
     * @return array
     */
    private function importSheet(SheetInterface $sheet, callable $callback = null)
    {
        $headers = [];
        $collection = [];
        $count_header = 0;

        if ($this->with_header) {
            foreach ($sheet->getRowIterator() as $k => $row) {
                if ($k == 1) {
                    $headers = $row;
                    $count_header = count($headers);
                    continue;
                }
                if ($count_header > $count_row = count($row)) {
                    $row = array_merge($row, array_fill(0, $count_header - $count_row, null));
                }
                $collection[] = $callback ? $callback(array_combine($headers, $row)) : array_combine($headers, $row);
            }
        } else {
            foreach ($sheet->getRowIterator() as $row) {
                $collection[] = $row;
            }
        }

        return $collection;
    }
}

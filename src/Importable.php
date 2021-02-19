<?php

namespace Rap2hpoutre\FastExcel;

use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Reader\SheetInterface;
use Illuminate\Support\Collection;

/**
 * Trait Importable.
 *
 * @property int  $start_row
 * @property bool $transpose
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
     * @return string
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
        $reader = $this->reader($path);

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
     * @param string        $path
     * @param callable|null $callback
     *
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     *
     * @return Collection
     */
    public function importSheets($path, callable $callback = null)
    {
        $reader = $this->reader($path);

        $collections = [];
        foreach ($reader->getSheetIterator() as $key => $sheet) {
            $collections[] = $this->importSheet($sheet, $callback);
        }
        $reader->close();

        return new SheetCollection($collections);
    }

    /**
     * @param $path
     *
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     *
     * @return \Box\Spout\Reader\ReaderInterface
     */
    private function reader($path)
    {
        $reader = ReaderFactory::create($this->getType($path));
        $this->setOptions($reader);
        /* @var \Box\Spout\Reader\ReaderInterface $reader */
        $reader->open($path);

        return $reader;
    }

    /**
     * @param array $array
     *
     * @return array
     */
    private function transposeCollection(array $array)
    {
        $collection = [];

        foreach ($array as $row => $columns) {
            foreach ($columns as $column => $value) {
                data_set($collection, implode('.', [
                    $column,
                    $row,
                ]), $value);
            }
        }

        return $collection;
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

        foreach ($sheet->getRowIterator() as $k => $row) {
            if ($k >= $this->start_row) {
                if ($this->with_header) {
                    if ($k == $this->start_row) {
                        $headers = $this->toStrings($row);
                        $count_header = count($headers);
                        continue;
                    }
                    if ($count_header > $count_row = count($row)) {
                        $row = array_merge($row, array_fill(0, $count_header - $count_row, null));
                    } elseif ($count_header < $count_row = count($row)) {
                        $row = array_slice($row, 0, $count_header);
                    }
                }
                if ($callback) {
                    if ($result = $callback(empty($headers) ? $row : array_combine($headers, $row))) {
                        $collection[] = $result;
                    }
                } else {
                    $collection[] = empty($headers) ? $row : array_combine($headers, $row);
                }
            }
        }

        if ($this->transpose) {
            return $this->transposeCollection($collection);
        }

        return $collection;
    }

    /**
     * @param array $values
     *
     * @return array
     */
    private function toStrings($values)
    {
        foreach ($values as &$value) {
            if ($value instanceof \Datetime) {
                $value = $value->format('Y-m-d H:i:s');
            } elseif ($value) {
                $value = (string) $value;
            }
        }

        return $values;
    }
}

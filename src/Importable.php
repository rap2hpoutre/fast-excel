<?php
namespace Rap2hpoutre\FastExcel;


use Box\Spout\Reader\ReaderFactory;
use Illuminate\Support\Collection;

/**
 * Trait Importable
 * @package Rap2hpoutre\FastExcel
 */
trait Importable
{

    /**
     *
     * @param string $path
     * @param callable|null $callback
     * @return Collection
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function import($path, callable $callback = null)
    {
        $headers = [];
        $collection = [];

        $reader = ReaderFactory::create($this->getType($path));
        $this->setOptions($reader);
        $reader->open($path);

        foreach ($reader->getSheetIterator() as $key => $sheet) {
            if ($this->sheet_number != $key) {
                continue;
            }
            if ($this->with_header) {
                foreach ($sheet->getRowIterator() as $k => $row) {
                    if ($k == 1) {
                        $headers = $row;
                        $count_header = count($headers);
                        continue;
                    }
                    if (($count_header ?? 0) > $count_row = count($row)) {
                        $row = array_merge($row, array_fill(0, $count_header - $count_row, null));
                    }
                    $collection[] = $callback ? $callback(array_combine($headers, $row)) : array_combine($headers, $row);
                }
            } else {
                foreach ($sheet->getRowIterator() as $row) {
                    $collection[] = $row;
                }
            }
        }
        $reader->close();

        return collect($collection);
    }

}
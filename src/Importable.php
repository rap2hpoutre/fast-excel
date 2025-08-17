<?php

namespace Rap2hpoutre\FastExcel;

use Illuminate\Support\Collection;
use Illuminate\Support\LazyCollection;
use Illuminate\Support\Str;
use OpenSpout\Common\Entity\Cell;
use OpenSpout\Reader\SheetInterface;
use OpenSpout\Writer\Common\AbstractOptions;

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
     * @param AbstractOptions $options
     *
     * @return mixed
     */
    abstract protected function setOptions(&$options);

    /**
     * @param string        $path
     * @param callable|null $callback
     *
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Common\Exception\IOException
     *
     * @return Collection
     */
    public function import($path, ?callable $callback = null)
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
     * Import file lazily using LazyCollection for memory efficiency.
     *
     * @param string        $path
     * @param callable|null $callback
     *
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Common\Exception\IOException
     *
     * @return LazyCollection
     */
    public function importLazy($path, ?callable $callback = null)
    {
        return new LazyCollection(function () use ($path, $callback) {
            $reader = $this->reader($path);

            try {
                foreach ($reader->getSheetIterator() as $key => $sheet) {
                    if ($this->sheet_number != $key) {
                        continue;
                    }
                    if ($this->transpose) {
                        // Fallback to non-lazy processing when transposing
                        throw new \Exception('Transposing is not supported with lazy import.');
                    }

                    yield from $this->importSheetGenerator($sheet, $callback);
                    break;
                }
            } finally {
                $reader->close();
            }
        });
    }

    /**
     * @param string        $path
     * @param callable|null $callback
     *
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Common\Exception\IOException
     *
     * @return Collection
     */
    public function importSheets($path, ?callable $callback = null)
    {
        $reader = $this->reader($path);

        $collections = [];
        foreach ($reader->getSheetIterator() as $key => $sheet) {
            if ($this->with_sheets_names) {
                $collections[$sheet->getName()] = $this->importSheet($sheet, $callback);
            } else {
                $collections[] = $this->importSheet($sheet, $callback);
            }
        }
        $reader->close();

        return new SheetCollection($collections);
    }

    /**
     * @param $path
     *
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Common\Exception\IOException
     *
     * @return \OpenSpout\Reader\ReaderInterface
     */
    private function reader($path)
    {
        if (Str::endsWith($path, 'csv')) {
            $options = new \OpenSpout\Reader\CSV\Options();
            $this->setOptions($options);
            $reader = new \OpenSpout\Reader\CSV\Reader($options);
        } elseif (Str::endsWith($path, 'ods')) {
            $options = new \OpenSpout\Reader\ODS\Options();
            $this->setOptions($options);
            $reader = new \OpenSpout\Reader\ODS\Reader($options);
        } else {
            $options = new \OpenSpout\Reader\XLSX\Options();
            $this->setOptions($options);
            $reader = new \OpenSpout\Reader\XLSX\Reader($options);
        }

        /* @var \OpenSpout\Reader\ReaderInterface $reader */
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
                data_set(
                    $collection,
                    implode('.', [
                        $column,
                        $row,
                    ]),
                    $value
                );
            }
        }

        return $collection;
    }

    /**
     * Normalize a row according to start_row and headers.
     * - Updates $headers and $count_header when encountering header row.
     * - Pads/truncates rows to header size when headers exist.
     * - Returns combined associative row when headers exist, or the raw row when not.
     * - Returns null to skip processing (before start_row or header row itself).
     *
     * @param int   $key
     * @param array $row
     * @param array $headers
     * @param int   $count_header
     *
     * @return array|null
     */
    private function normalizeRow(int $key, array $row, array &$headers, int &$count_header): ?array
    {
        if ($key < $this->start_row) {
            return null;
        }

        if ($this->with_header) {
            if ($k == $this->start_row) {
                $headers = $this->toStrings($row);
                $count_header = count($headers);

                return null; // skip header row
            }

            if ($count_header > $count_row = count($row)) {
                $row = array_merge($row, array_fill(0, $count_header - $count_row, null));
            } elseif ($count_header < $count_row = count($row)) {
                $row = array_slice($row, 0, $count_header);
            }
        }

        return empty($headers) ? $row : array_combine($headers, $row);
    }

    /**
     * @param SheetInterface $sheet
     * @param callable|null  $callback
     *
     * @return array
     */
    private function importSheet(SheetInterface $sheet, ?callable $callback = null)
    {
        $headers = [];
        $collection = [];
        $count_header = 0;

        foreach ($sheet->getRowIterator() as $key => $rowAsObject) {
            $row = array_map(function (Cell $cell) {
                return match (true) {
                    $cell instanceof Cell\FormulaCell => $cell->getComputedValue(),
                    default                           => $cell->getValue(),
                };
            }, $rowAsObject->getCells());

            $current = $this->normalizeRow($key, $row, $headers, $count_header);
            if ($current === null) {
                continue;
            }

            if ($callback) {
                if ($result = $callback($current)) {
                    $collection[] = $result;
                }
            } else {
                $collection[] = $current;
            }
        }

        if ($this->transpose) {
            return $this->transposeCollection($collection);
        }

        return $collection;
    }

    /**
     * Create a generator that lazily yields imported rows from a sheet.
     *
     * @param SheetInterface $sheet
     * @param callable|null  $callback
     *
     * @return \Generator
     */
    private function importSheetGenerator(SheetInterface $sheet, ?callable $callback = null): \Generator
    {
        $headers = [];
        $count_header = 0;

        foreach ($sheet->getRowIterator() as $k => $rowAsObject) {
            $row = array_map(function (Cell $cell) {
                return match (true) {
                    $cell instanceof Cell\FormulaCell => $cell->getComputedValue(),
                    default                           => $cell->getValue(),
                };
            }, $rowAsObject->getCells());

            $current = $this->normalizeRow($k, $row, $headers, $count_header);
            if ($current === null) {
                continue;
            }

            if ($callback) {
                $result = $callback($current);
                if ($result) {
                    yield $result;
                }
            } else {
                yield $current;
            }
        }
    }

    /**
     * @param array $values
     *
     * @return array
     */
    private function toStrings($values)
    {
        foreach ($values as &$value) {
            if ($value instanceof \DateTime) {
                $value = $value->format('Y-m-d H:i:s');
            } elseif ($value instanceof \DateTimeImmutable) {
                $value = $value->format('Y-m-d H:i:s');
            } elseif ($value) {
                $value = (string) $value;
            }
        }

        return $values;
    }
}

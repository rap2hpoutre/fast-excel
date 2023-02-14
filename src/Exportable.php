<?php

namespace Rap2hpoutre\FastExcel;

use Generator;
use Illuminate\Support\Collection;
use Illuminate\Support\Str;
use InvalidArgumentException;
use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\Style;

/**
 * Trait Exportable.
 *
 * @property bool                           $transpose
 * @property bool                           $with_header
 * @property \Illuminate\Support\Collection $data
 */
trait Exportable
{
    private ?Style $header_style = null;
    private ?Style $rows_style = null;

    /**
     * @param string        $path
     * @param callable|null $callback
     *
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     *
     * @return string
     */
    public function export(string $path, callable $callback = null)
    {
        self::exportOrDownload($path, 'openToFile', $callback);

        return realpath($path) ?: $path;
    }

    /**
     * @param $path
     * @param callable|null $callback
     *
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     *
     * @return \Symfony\Component\HttpFoundation\StreamedResponse|string
     */
    public function download($path, callable $callback = null)
    {
        if (method_exists(response(), 'streamDownload')) {
            return response()->streamDownload(function () use ($path, $callback) {
                self::exportOrDownload($path, 'openToBrowser', $callback);
            });
        }
        self::exportOrDownload($path, 'openToBrowser', $callback);

        return '';
    }

    /**
     * @param Style $style
     *
     * @return Exportable
     */
    public function headerStyle(Style $style)
    {
        $this->header_style = $style;

        return $this;
    }

    /**
     * @param Style $style
     *
     * @return Exportable
     */
    public function rowsStyle(Style $style)
    {
        $this->rows_style = $style;

        return $this;
    }

    /**
     * @param \OpenSpout\Reader\ReaderInterface|\OpenSpout\Writer\WriterInterface $reader_or_writer
     *
     * @return mixed
     */
    abstract protected function setOptions(&$reader_or_writer);

    /**
     * @param $path
     * @param string        $function
     * @param callable|null $callback
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     * @throws \OpenSpout\Common\Exception\SpoutException
     */
    private function exportOrDownload($path, string $function, callable $callback = null)
    {
        /* @var \OpenSpout\Writer\WriterInterface $writer */
        $writer = $this->prepareWriter($path);
        $options = $writer->getOptions();
        $this->setOptions($options);
        $writer->$function($path);

        // It can export one sheet (Collection) or N sheets (SheetCollection)
        $data = $this->prepareDataForExport();

        foreach ($data as $key => $collection) {
            if ($collection instanceof Collection) {
                $this->writeRowsFromCollection($writer, $collection, $callback);
            } elseif ($collection instanceof Generator) {
                $this->writeRowsFromGenerator($writer, $collection, $callback);
            } elseif (is_array($collection)) {
                $this->writeRowsFromArray($writer, $collection, $callback);
            } else {
                throw new InvalidArgumentException('Unsupported type for $data');
            }
            if (is_string($key)) {
                $writer->getCurrentSheet()->setName($key);
            }
            if ($this->hasSheets($writer) && $data->keys()->last() !== $key) {
                $writer->addNewSheetAndMakeItCurrent();
            }
        }
        $writer->close();
    }

    private function prepareWriter($path): \OpenSpout\Writer\WriterInterface
    {
        if (Str::endsWith($path, 'csv')) {
            $writer = new \OpenSpout\Writer\CSV\Writer(new \OpenSpout\Writer\CSV\Options());
        } elseif (Str::endsWith($path, 'ods')) {
            $writer = new \OpenSpout\Writer\ODS\Writer(new \OpenSpout\Writer\ODS\Options());
        } else {
            $writer = new \OpenSpout\Writer\XLSX\Writer(new \OpenSpout\Writer\XLSX\Options());
        }

        return $writer;
    }

    private function hasSheets(\OpenSpout\Writer\WriterInterface $writer): bool
    {
        return $writer instanceof \OpenSpout\Writer\XLSX\Writer || $writer instanceof \OpenSpout\Writer\ODS\Writer;
    }

    private function prepareDataForExport(): SheetCollection|array|Generator|Collection|null
    {
        return $this->transpose
            ? $this->transposeData()
            : (
                $this->data instanceof SheetCollection
                ? $this->data
                : collect([$this->data])
            );
    }

    /**
     * Transpose data from rows to columns.
     *
     * @return SheetCollection
     */
    private function transposeData(): SheetCollection
    {
        $data = $this->data instanceof SheetCollection ? $this->data : collect([$this->data]);
        $transposedData = [];

        foreach ($data as $key => $collection) {
            foreach ($collection as $row => $columns) {
                foreach ($columns as $column => $value) {
                    data_set(
                        $transposedData,
                        implode('.', [
                            $key,
                            $column,
                            $row,
                        ]),
                        $value
                    );
                }
            }
        }

        return new SheetCollection($transposedData);
    }

    private function writeRowsFromCollection($writer, Collection $collection, ?callable $callback = null)
    {
        // Apply callback
        if ($callback) {
            $collection->transform(function ($value) use ($callback) {
                return $callback($value);
            });
        }

        // Prepare collection (i.e remove non-string)
        // and transform into collection of row cells collections
        //$this->transformCollection($collection);

        // Add header row.
        if ($this->with_header) {
            $this->writeHeader($writer, $this->transformRow($collection->first()));
        }

        // Add rows
        $writer->addRows(
            $collection->map(function ($row) {
                return $this->createRow($this->transformRow($row), $this->rows_style);
            })->toArray()
        );
    }

    private function writeRowsFromGenerator($writer, Generator $generator, ?callable $callback = null)
    {
        foreach ($generator as $key => $item) {
            // Apply callback
            if ($callback) {
                $item = $callback($item);
            }

            // Prepare row (i.e remove non-string)
            // and transform to collection of Cells
            $row_cells = $this->transformRow($item);

            // Add header row.
            if ($this->with_header && $key === 0) {
                $this->writeHeader($writer, $row_cells);
            }
            // Write rows (one by one).
            $writer->addRow(
                $this->createRow($row_cells, $this->rows_style)
            );
        }
    }

    private function writeRowsFromArray($writer, array $array, ?callable $callback = null)
    {
        $collection = collect($array);

        if (is_object($collection->first()) || is_array($collection->first())) {
            // provided $array was valid and could be converted to a collection
            $this->writeRowsFromCollection($writer, $collection, $callback);
        }
    }

    private function writeHeader($writer, array $first_row)
    {
        $writer->addRow(
            $this->createRow($this->transformRow(array_keys($first_row)), $this->header_style)
        );
    }

    /**
     * Transform one row (i.e remove non-string).
     * into array of Cells.
     */
    private function transformRow($data): array
    {
        return collect($data)
            ->filter(function ($value) {
                return $this->isValidValue($value);
            })->map(function ($value) {
                return $this->createCell($value, null);
            })->toArray();
    }

    private function isValidValue($value): bool
    {
        return is_string($value) || is_int($value) || is_float($value)
            || is_null($value) || $value instanceof Cell;
    }

    private function createCell($value, ?Style $style): Cell
    {
        return  ($value instanceof Cell)
            ? $value
            : Cell::fromValue($value, $style);
    }

    private function createRow(array $row_cells, ?Style $row_style): Row
    {
        return new Row($row_cells, $row_style);
    }
}

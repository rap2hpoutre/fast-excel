<?php

namespace Rap2hpoutre\FastExcel;

use Generator;
use Illuminate\Support\Collection;
use Illuminate\Support\Str;
use InvalidArgumentException;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Common\Type;
use OpenSpout\Writer\Common\Creator\WriterEntityFactory;

/**
 * Trait Exportable.
 *
 * @property bool                           $transpose
 * @property bool                           $with_header
 * @property \Illuminate\Support\Collection $data
 */
trait Exportable
{
    /**
     * @var Style
     */
    private $header_style;
    private $rows_style;

    /**
     * @param \OpenSpout\Reader\ReaderInterface|\OpenSpout\Writer\WriterInterface $reader_or_writer
     *
     * @return mixed
     */
    abstract protected function setOptions(&$reader_or_writer);

    /**
     * @param string        $path
     * @param callable|null $callback
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     *
     * @return string
     */
    public function export($path, callable $callback = null)
    {
        self::exportOrDownload($path, 'openToFile', $callback);

        return realpath($path) ?: $path;
    }

    /**
     * @param $path
     * @param callable|null $callback
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
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
    private function exportOrDownload($path, $function, callable $callback = null)
    {
        if (Str::endsWith($path, Type::CSV)) {
            $writer = WriterEntityFactory::createCSVWriter();
        } elseif (Str::endsWith($path, Type::ODS)) {
            $writer = WriterEntityFactory::createODSWriter();
        } else {
            $writer = WriterEntityFactory::createXLSXWriter();
        }

        $this->setOptions($writer);
        /* @var \OpenSpout\Writer\WriterInterface $writer */
        $writer->$function($path);

        $has_sheets = ($writer instanceof \OpenSpout\Writer\XLSX\Writer || $writer instanceof \OpenSpout\Writer\ODS\Writer);

        // It can export one sheet (Collection) or N sheets (SheetCollection)
        $data = $this->transpose ? $this->transposeData() : ($this->data instanceof SheetCollection ? $this->data : collect([$this->data]));

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
            if ($has_sheets && $data->keys()->last() !== $key) {
                $writer->addNewSheetAndMakeItCurrent();
            }
        }
        $writer->close();
    }

    /**
     * Transpose data from rows to columns.
     *
     * @return SheetCollection
     */
    private function transposeData()
    {
        $data = $this->data instanceof SheetCollection ? $this->data : collect([$this->data]);
        $transposedData = [];

        foreach ($data as $key => $collection) {
            foreach ($collection as $row => $columns) {
                foreach ($columns as $column => $value) {
                    data_set($transposedData, implode('.', [
                        $key,
                        $column,
                        $row,
                    ]), $value);
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
        $this->prepareCollection($collection);
        // Add header row.
        if ($this->with_header) {
            $this->writeHeader($writer, $collection->first());
        }

        // createRowFromArray works only with arrays
        if (!is_array($collection->first())) {
            $collection = $collection->map(function ($value) {
                return $value->toArray();
            });
        }

        // is_array($first_row) ? $first_row : $first_row->toArray())
        $all_rows = $collection->map(function ($value) {
            return WriterEntityFactory::createRowFromArray($value);
        })->toArray();
        if ($this->rows_style) {
            $this->addRowsWithStyle($writer, $all_rows, $this->rows_style);
        } else {
            $writer->addRows($all_rows);
        }
    }

    private function addRowsWithStyle($writer, $all_rows, $rows_style)
    {
        $styled_rows = [];
        // Style rows one by one
        foreach ($all_rows as $row) {
            $row = WriterEntityFactory::createRowFromArray($row->toArray(), $rows_style);
            array_push($styled_rows, $row);
        }
        $writer->addRows($styled_rows);
    }

    private function writeRowsFromGenerator($writer, Generator $generator, ?callable $callback = null)
    {
        foreach ($generator as $key => $item) {
            // Apply callback
            if ($callback) {
                $item = $callback($item);
            }

            // Prepare row (i.e remove non-string)
            $item = $this->transformRow($item);

            // Add header row.
            if ($this->with_header && $key === 0) {
                $this->writeHeader($writer, $item);
            }
            // Write rows (one by one).
            $writer->addRow(WriterEntityFactory::createRowFromArray($item->toArray(), $this->rows_style));
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

    private function writeHeader($writer, $first_row)
    {
        if ($first_row === null) {
            return;
        }

        $keys = array_keys(is_array($first_row) ? $first_row : $first_row->toArray());

        $writer->addRow(WriterEntityFactory::createRowFromArray($keys, $this->header_style));
    }

    /**
     * Prepare collection by removing non string if required.
     */
    protected function prepareCollection(Collection $collection)
    {
        $need_conversion = false;
        $first_row = $collection->first();

        if (!$first_row) {
            return;
        }

        foreach ($first_row as $item) {
            if (!is_string($item)) {
                $need_conversion = true;
            }
        }
        if ($need_conversion) {
            $this->transform($collection);
        }
    }

    /**
     * Transform the collection.
     */
    private function transform(Collection $collection)
    {
        $collection->transform(function ($data) {
            return $this->transformRow($data);
        });
    }

    /**
     * Transform one row (i.e remove non-string).
     */
    private function transformRow($data)
    {
        return collect($data)->map(function ($value) {
            return is_null($value) ? (string) $value : $value;
        })->filter(function ($value) {
            return is_string($value) || is_int($value) || is_float($value);
        });
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
}

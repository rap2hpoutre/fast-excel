<?php

namespace Rap2hpoutre\FastExcel;

use Box\Spout\Writer\Style\Style;
use Box\Spout\Writer\WriterFactory;
use Generator;
use Illuminate\Support\Collection;
use InvalidArgumentException;

/**
 * Trait Exportable.
 *
 * @property bool $with_header
 * @property \Illuminate\Support\Collection $data
 */
trait Exportable
{
    /**
     * @var Style
     */
    private $header_style;

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
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
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
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
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
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     * @throws \Box\Spout\Common\Exception\SpoutException
     */
    private function exportOrDownload($path, $function, callable $callback = null)
    {
        $writer = WriterFactory::create($this->getType($path));
        $this->setOptions($writer);
        /* @var \Box\Spout\Writer\WriterInterface $writer */
        $writer->$function($path);

        $has_sheets = ($writer instanceof \Box\Spout\Writer\XLSX\Writer || $writer instanceof \Box\Spout\Writer\ODS\Writer);

        // It can export one sheet (Collection) or N sheets (SheetCollection)
        $data = $this->data instanceof SheetCollection ? $this->data : collect([$this->data]);

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
        // Write all rows
        $writer->addRows($collection->toArray());
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
            $writer->addRow($item->toArray());
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
        if ($this->header_style) {
            $writer->addRowWithStyle($keys, $this->header_style);
        } else {
            $writer->addRow($keys);
        }
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
            return is_int($value) || is_float($value) || is_null($value) ? (string) $value : $value;
        })->filter(function ($value) {
            return is_string($value);
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
}

<?php

namespace Smart145\FastExcel;

use Box\Spout\Writer\Style\Style;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Writer\XLSX\Writer;
use Illuminate\Support\Collection;

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
     * @var array
     */
    private $columns_width;

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
     * @param string $path
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

    public function setColumnsWidth($widths)
    {
        $this->columns_width = $widths;

        return $this;
    }

    /**
     * @param $path
     * @param string $function
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
        $this->customizeColumnsWidth($writer);
        /* @var \Box\Spout\Writer\WriterInterface $writer */
        $writer->$function($path);
        $writer->getCurrentSheet()->setName($this->sheet);


        $has_sheets = ($writer instanceof \Box\Spout\Writer\XLSX\Writer || $writer instanceof \Box\Spout\Writer\ODS\Writer);

        // It can export one sheet (Collection) or N sheets (SheetCollection)
        $data = $this->data instanceof SheetCollection ? $this->data : collect([$this->data]);

        foreach ($data as $key => $collection) {
            if ($collection instanceof Collection) {
                // Apply callback
                if ($callback) {
                    $collection->transform(function ($value) use ($callback) {
                        return $callback($value);
                    });
                }
                // Prepare collection (i.e remove non-string)
                $this->prepareCollection();
                // Add header row.
                if ($this->with_header) {
                    $first_row = $collection->first();
                    $keys = $this->hasColumnsHeader() ? array_keys($this->columns_width) : array_keys(is_array($first_row) ? $first_row : $first_row->toArray());
                    if ($this->header_style) {
                        $writer->addRowWithStyle($keys, $this->header_style);
                    } else {
                        $writer->addRow($keys);
                    }
                }
                $writer->addRows($collection->toArray());
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

    private function hasColumnsHeader()
    {
        return $this->columns_width && \count($this->columns_width) && is_string(array_keys($this->columns_width)[0]);
    }

    private function customizeColumnsWidth(Writer $writer)
    {
        $data = $this->data->first();
        $keys = get_object_vars($data);
        $columnsWithCount = count($this->columns_width);
        if ($columnsWithCount === 0) {
            return;
        }
        if ($columnsWithCount > 0 && $columnsWithCount < \count($keys)) {
            throw new \Exception('The columns width elements count must match with header count.');
        }
        $i = 1;
        foreach ($this->columns_width as $width) {
            $writer->setColumnsWidth($width, $i, $i++);
        }
    }

    /**
     * Prepare collection by removing non string if required.
     */
    protected function prepareCollection()
    {
        $need_conversion = false;
        $first_row = $this->data->first();

        if (!$first_row) {
            return;
        }

        foreach ($first_row as $item) {
            if (!is_string($item)) {
                $need_conversion = true;
            }
        }
        //if ($need_conversion) {
            $this->transform();
        //}
    }

    /**
     * Transform the collection.
     */
    private function transform()
    {
        $this->data->transform(function ($data) {
            return collect($data)->map(function ($value) {
                return is_int($value) || is_float($value) || is_null($value) ? (string) $value : $value;
            })->filter(function ($value) {
                return is_string($value);
            });
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


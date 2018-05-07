<?php

namespace Rap2hpoutre\FastExcel;

use Box\Spout\Writer\WriterFactory;
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
     * @param callable|null $callback
     * @return string
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
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
     * @return string
     */
    public function download($path, callable $callback = null)
    {
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
     */
    private function exportOrDownload($path, $function, callable $callback = null)
    {
        $writer = WriterFactory::create($this->getType($path));
        $this->setOptions($writer);
        /* @var \Box\Spout\Writer\WriterInterface $writer */
        $writer->$function($path);
        if ($this->data instanceof Collection) {
            // Prepare collection (i.e remove non-string) only if there is no callback
            if (!$callback) {
                $this->prepareCollection();
            } else {
                $this->data->transform(function ($value) use ($callback) {
                    return $callback($value);
                });
            }
            // Add header row.
            if ($this->with_header) {
                $first_row = $this->data->first();
                $keys = array_keys(is_array($first_row) ? $first_row : $first_row->toArray());
                $writer->addRow($keys);
            }
            $writer->addRows($this->data->toArray());
        }
        $writer->close();
    }

    /**
     * Prepare collection by removing non string if required.
     */
    protected function prepareCollection()
    {
        $need_conversion = false;
        $first_row = $this->data->first();
        foreach ($first_row as $item) {
            if (!is_string($item)) {
                $need_conversion = true;
            }
        }
        if ($need_conversion) {
            $this->data->transform(function ($data) {
                return collect($data)->filter(function ($value) {
                    return is_string($value);
                });
            });
        }
    }
}

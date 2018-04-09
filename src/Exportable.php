<?php
namespace Rap2hpoutre\FastExcel;


use Box\Spout\Writer\WriterFactory;
use Illuminate\Support\Collection;

/**
 * Trait Exportable
 * @package Rap2hpoutre\FastExcel
 */
trait Exportable
{
    /**
     * @var Collection
     */
    protected $data;

    /**
     * @var bool
     */
    private $with_header = true;

    /**
     * @param string $path
     * @return mixed
     */
    abstract protected function getType($path);

    /**
     * @param \Box\Spout\Reader\ReaderInterface|\Box\Spout\Writer\WriterInterface $reader_or_writer
     * @return mixed
     */
    abstract protected function setOptions(&$reader_or_writer);

    /**
     * @param string $path
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function export($path)
    {
        self::exportOrDownload($path, 'openToFile');
    }

    /**
     * @param $path
     * @return string
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function download($path)
    {
        self::exportOrDownload($path, 'openToBrowser');
        return '';
    }

    /**
     * @param $path
     * @param string $function
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    private function exportOrDownload($path, $function)
    {
        $writer = WriterFactory::create($this->getType($path));
        $this->setOptions($writer);
        $writer->$function($path);
        if ($this->data instanceof Collection) {
            $this->prepareCollection();
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
        foreach($first_row as $item) {
            if (!is_string($item)) {
                $need_conversion = true;
            }
        }
        if ($need_conversion) {
            $this->data->transform(function($data) {
                return collect($data)->filter(function ($value) {
                   return is_string($value);
                });
            });
        }
    }

}
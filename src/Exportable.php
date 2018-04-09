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
     * @var
     */
    protected $data;

    /**
     * @param $path
     * @return mixed
     */
    abstract protected function getType($path);

    /**
     * @param $reader_or_writer
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
            $first_row = $this->data->first();
            $keys = array_keys(is_array($first_row) ? $first_row : $first_row->toArray());
            $writer->addRow($keys);
            $writer->addRows($this->data->toArray());
        }
        $writer->close();
    }

}
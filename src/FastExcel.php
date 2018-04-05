<?php


namespace Rap2hpoutre\FastExcel;

use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Common\Type;
use Illuminate\Support\Collection;

/**
 * Class FastExcel
 * @package App
 */
class FastExcel
{
    /**
     * @var
     */
    private $data;

    /**
     * @var int
     */
    private $sheet_number = 1;
    /**
     * @var bool
     */
    private $with_header = true;

    /**
     * FastExcel constructor.
     * @param $data
     */
    public function __construct($data = null)
    {
        $this->data = $data;
    }

    /**
     * @param $path
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
     * @param $function
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    private function exportOrDownload($path, $function)
    {
        $writer = WriterFactory::create($this->getType($path));
        $writer->$function($path);
        if ($this->data instanceof Collection) {
            $first_row = $this->data->first();
            $keys = array_keys(is_array($first_row) ? $first_row : $first_row->toArray());
            $writer->addRow($keys);
            $writer->addRows($this->data->toArray());
        }
        $writer->close();
    }

    /**
     * @param $path
     * @return string
     */
    private function getType($path)
    {
        if (ends_with($path, Type::CSV)) {
            return Type::CSV;
        } elseif (ends_with($path, Type::ODS)) {
            return Type::ODS;
        } else {
            return Type::XLSX;
        }
    }

    /**
     * @param $sheet_number
     * @return $this
     */
    public function sheet($sheet_number)
    {
        $this->sheet_number = $sheet_number;
        return $this;
    }

    /**
     * @return $this
     */
    public function withoutHeaders()
    {
        $this->with_headers = false;
        return $this;
    }

    /**
     *
     * @param $path
     * @return Collection
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function import($path)
    {
        $headers = [];
        $collection = [];

        $reader = ReaderFactory::create($this->getType($path));
        $reader->open($path);

        foreach ($reader->getSheetIterator() as $key => $sheet) {
            if ($this->sheet_number != $key) {
                continue;
            }
            if ($this->with_header) {
                foreach ($sheet->getRowIterator() as $k => $row) {
                    if ($k == 1) {
                        $headers = $row;
                        continue;
                    }
                    if ($count_header = count($headers) > $count_row = count((array)$row)) {
                        $row = array_merge($row, array_fill(0, $count_header - $count_row, null));
                        dd($row, $headers);
                    }
                    $collection[] = array_combine($headers, $row);
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
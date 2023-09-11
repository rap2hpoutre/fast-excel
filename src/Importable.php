<?php

namespace Rap2hpoutre\FastExcel;

use Illuminate\Support\Collection;
use Illuminate\Support\Str;
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
    public function import($path, callable $callback = null)
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
     * @param string        $path
     * @param callable|null $callback
     *
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Common\Exception\IOException
     *
     * @return Collection
     */
    public function importSheets($path, callable $callback = null)
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
     * @param SheetInterface $sheet
     * @param callable|null  $callback
     *
     * @return array
     */
    private function importSheet(SheetInterface $sheet, callable $callback = null)
    {
        $headers = [];
        $collection = [];
        $count_header = 0;

        foreach ($sheet->getRowIterator() as $row => $columnsAsObject) {
            $columns = $columnsAsObject->toArray();
            $count_columns = count($columns);

            if ($row >= $this->start_row) {
                if ($this->with_header) {
                    if ($row == $this->start_row) {
                        $headers = $this->toStrings($columns);
                        $count_header = count($headers);
                        continue;
                    }
                    if ($count_header > $count_columns) {
                        $columns = array_merge($columns, array_fill(0, $count_header - $count_columns, null));
                    } elseif ($count_header < $count_columns) {
                        $columns = array_slice($columns, 0, $count_header);
                    }
                }
                if ($callback) {
                    if ($result = $callback(empty($headers) ? $columns : array_combine($headers, $columns))) {
                        $collection[] = $result;
                    }
                } else {
                     $collection[] = $this->with_header ? $this->checkHeaders($headers, $columns) : $columns;
                }
            }
        }

        if ($this->transpose) {
            return $this->transposeCollection($collection);
        }

        return $collection;
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
    
    /**
     * if not actived => Empty column headers skipped and their values(rows) not appear on your results
     *                => Duplicate headers replaced 
     *
     * if actived     => Empty column headers named by "COL_INDEX"
     *                => Duplicate headers rename and add "COL_INDEX" to it
     * just set third parameter to False and it works as always be
     *
     * @param $headers
     * @param $columns
     * @param true $active
     * @return array|mixed
     */
    private function checkHeaders($headers, $columns, $active = true){
        $result = [];
        if($active){
            foreach($headers as $index => $header){
                /* Name COL_XX for empty headers */
                $key = empty($header) ? 'COL_'.$index : $header;
                /* Duplicate headers check and add col_index to duplicated ones */
                $key = array_key_exists($key, $result) ? $key."_".$index : $key;
                $result[$key] = $columns[$index];
            }
        }else{
            $result = empty($headers) ? $columns : array_combine($headers, $columns);
        }
        return $result;
    }
}

<?php

namespace Rap2hpoutre\FastExcel;

use Generator;
use Illuminate\Support\Collection;
use OpenSpout\Reader\CSV\Options as CsvReaderOptions;
use OpenSpout\Reader\CSV\Reader;
use OpenSpout\Reader\ReaderInterface;
use OpenSpout\Writer\CSV\Options as CsvWriterOptions;
use OpenSpout\Writer\CSV\Writer;
use OpenSpout\Writer\WriterInterface;

/**
 * Class FastExcel.
 */
class FastExcel
{
    use Importable;
    use Exportable;

    /**
     * @var Collection|Generator|array
     */
    protected $data;

    /**
     * @var bool
     */
    private $with_header = true;

    /**
     * @var bool
     */
    private $with_sheets_names = false;

    /**
     * @var int
     */
    private $start_row = 1;

    /**
     * @var bool
     */
    private $transpose = false;

    /**
     * @var
     */
    private $csv_configuration = [
        'delimiter' => ',',
        'enclosure' => '"',
        'encoding'  => 'UTF-8',
        'bom'       => true,
    ];

    /**
     * @var callable
     */
    protected $reader_configurator = null;

    /**
     * @var callable
     */
    protected $writer_configurator = null;

    /**
     * FastExcel constructor.
     *
     * @param array|Generator|Collection|null $data
     */
    public function __construct(array|Generator|Collection $data = null)
    {
        $this->data = $data;
    }

    /**
     * Manually set data apart from the constructor.
     *
     * @param Collection|Generator|array $data
     *
     * @return FastExcel
     */
    public function data($data)
    {
        $this->data = $data;

        return $this;
    }

    /**
     * @param $sheet_number
     *
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
        $this->with_header = false;

        return $this;
    }

    /**
     * @return $this
     */
    public function withSheetsNames()
    {
        $this->with_sheets_names = true;

        return $this;
    }

    /**
     * @return $this
     */
    public function startRow(int $row)
    {
        $this->start_row = $row;

        return $this;
    }

    /**
     * @return $this
     */
    public function transpose()
    {
        $this->transpose = true;

        return $this;
    }

    /**
     * @param string $delimiter
     * @param string $enclosure
     * @param string $encoding
     * @param bool   $bom
     *
     * @return $this
     */
    public function configureCsv($delimiter = ',', $enclosure = '"', $encoding = 'UTF-8', $bom = false)
    {
        $this->csv_configuration = compact('delimiter', 'enclosure', 'encoding', 'bom');

        return $this;
    }

    /**
     * Configure the underlying Spout Reader using a callback.
     *
     * @param callable|null $callback
     *
     * @return $this
     */
    public function configureReaderUsing(?callable $callback = null)
    {
        $this->reader_configurator = $callback;

        return $this;
    }

    /**
     * Configure the underlying Spout Reader using a callback.
     *
     * @param callable|null $callback
     *
     * @return $this
     */
    public function configureWriterUsing(?callable $callback = null)
    {
        $this->writer_configurator = $callback;

        return $this;
    }

    /**
     * @param \OpenSpout\Reader\ReaderInterface|\OpenSpout\Writer\WriterInterface $reader_or_writer
     */
    protected function setOptions(&$reader_or_writer)
    {
        if ($reader_or_writer instanceof CsvReaderOptions || $reader_or_writer instanceof CsvWriterOptions) {
            $reader_or_writer->FIELD_DELIMITER = $this->csv_configuration['delimiter'];
            $reader_or_writer->FIELD_ENCLOSURE = $this->csv_configuration['enclosure'];
            if ($reader_or_writer instanceof CsvReaderOptions) {
                $reader_or_writer->ENCODING = $this->csv_configuration['encoding'];
            }
            if ($reader_or_writer instanceof CsvWriterOptions) {
                $reader_or_writer->SHOULD_ADD_BOM = $this->csv_configuration['bom'];
            }
        }

        if ($reader_or_writer instanceof ReaderInterface && is_callable($this->reader_configurator)) {
            call_user_func(
                $this->reader_configurator,
                $reader_or_writer
            );
        } elseif ($reader_or_writer instanceof WriterInterface && is_callable($this->writer_configurator)) {
            call_user_func(
                $this->writer_configurator,
                $reader_or_writer
            );
        }
    }
}

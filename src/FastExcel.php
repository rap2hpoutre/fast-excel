<?php

namespace Rap2hpoutre\FastExcel;

use Box\Spout\Common\Type;
use Box\Spout\Reader\CSV\Reader as CSVReader;
use Box\Spout\Reader\ReaderInterface;
use Box\Spout\Writer\CSV\Writer as CSVWriter;
use Box\Spout\Writer\WriterInterface;
use Illuminate\Support\Collection;
use Illuminate\Support\Str;

/**
 * Class FastExcel.
 */
class FastExcel
{
    use Importable, Exportable;

    /**
     * @var Collection
     */
    protected $data;

    /**
     * @var bool
     */
    private $with_header = true;

    /**
     * @var
     */
    private $csv_configuration = [
        'delimiter' => ',',
        'enclosure' => '"',
        'eol'       => "\n",
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
     * @param Collection $data
     */
    public function __construct($data = null)
    {
        $this->data = $data;
    }

    /**
     * Manually set data apart from the constructor.
     *
     * @param Collection $data
     *
     * @return FastExcel
     */
    public function data($data)
    {
        $this->data = $data;

        return $this;
    }

    /**
     * @param $path
     *
     * @return string
     */
    protected function getType($path)
    {
        if (Str::endsWith($path, Type::CSV)) {
            return Type::CSV;
        } elseif (Str::endsWith($path, Type::ODS)) {
            return Type::ODS;
        } else {
            return Type::XLSX;
        }
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
     * @param string $delimiter
     * @param string $enclosure
     * @param string $eol
     * @param string $encoding
     * @param bool   $bom
     *
     * @return $this
     */
    public function configureCsv($delimiter = ',', $enclosure = '"', $eol = "\n", $encoding = 'UTF-8', $bom = false)
    {
        $this->csv_configuration = compact('delimiter', 'enclosure', 'eol', 'encoding', 'bom');

        return $this;
    }

    /**
     * Configure the underlying Spout Reader using a callback.
     *
     * @param callable $callback
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
     * @param callable $callback
     *
     * @return $this
     */
    public function configureWriterUsing(?callable $callback = null)
    {
        $this->writer_configurator = $callback;

        return $this;
    }

    /**
     * @param \Box\Spout\Reader\ReaderInterface|\Box\Spout\Writer\WriterInterface $reader_or_writer
     */
    protected function setOptions(&$reader_or_writer)
    {
        if ($reader_or_writer instanceof CSVReader || $reader_or_writer instanceof CSVWriter) {
            $reader_or_writer->setFieldDelimiter($this->csv_configuration['delimiter']);
            $reader_or_writer->setFieldEnclosure($this->csv_configuration['enclosure']);
            if ($reader_or_writer instanceof CSVReader) {
                $reader_or_writer->setEndOfLineCharacter($this->csv_configuration['eol']);
                $reader_or_writer->setEncoding($this->csv_configuration['encoding']);
            }
            if ($reader_or_writer instanceof CSVWriter) {
                $reader_or_writer->setShouldAddBOM($this->csv_configuration['bom']);
            }
        }

        if ($reader_or_writer instanceof ReaderInterface && is_callable($this->reader_configurator)) {
            call_user_func(
                $this->reader_configurator, $reader_or_writer
            );
        } elseif ($reader_or_writer instanceof WriterInterface && is_callable($this->writer_configurator)) {
            call_user_func(
                $this->writer_configurator, $reader_or_writer
            );
        }
    }
}

<?php


namespace Rap2hpoutre\FastExcel;

use Box\Spout\Reader\CSV\Reader as CSVReader;
use Box\Spout\Reader\ReaderInterface;
use Box\Spout\Writer\CSV\Writer as CSVWriter;
use Box\Spout\Common\Type;
use Box\Spout\Writer\WriterInterface;
use Illuminate\Support\Collection;

/**
 * Class FastExcel
 * @package Rap2hpoutre\FastExcel
 */
class FastExcel
{
    use Importable, Exportable;

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
     * @var
     */
    private $csv_configuration = [
        'delimiter' => ',',
        'enclosure' => '"',
        'eol' => "\n",
        'encoding' => 'UTF-8',
        'bom' => true,
    ];

    /**
     * FastExcel constructor.
     * @param Collection $data
     */
    public function __construct($data = null)
    {
        $this->data = $data;
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
        $this->with_header = false;
        return $this;
    }

    /**
     * @param string $delimiter
     * @param string $enclosure
     * @param string $eol
     * @param string $encoding
     * @param bool $bom
     * @return $this
     */
    public function configureCsv($delimiter, $enclosure = '"', $eol = "\n", $encoding = 'UTF-8', $bom = false)
    {
        $this->csv_configuration = compact('delimiter', 'enclosure', 'eol', 'encoding', 'bom');
        return $this;
    }

    /**
     * @param ReaderInterface|WriterInterface $reader_or_writer
     */
    private function setOptions(&$reader_or_writer)
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
    }
}

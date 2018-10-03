<?php

namespace Rap2hpoutre\FastExcel;

use Box\Spout\Common\Type;
use Box\Spout\Reader\CSV\Reader as CSVReader;
use Box\Spout\Writer\CSV\Writer as CSVWriter;
use Box\Spout\Writer\Style\Color;
use Box\Spout\Writer\Style\StyleBuilder;
use Illuminate\Support\Collection;

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
     * @var bool
     */
    public $headerStyle = true;

    /**
     * @var bool
     */

    private $styleHeader;

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
     * FastExcel constructor.
     *
     * @param Collection $data
     */
    public function __construct($data = null)
    {
        $this->data = $data;
        $this->setHeaderStyle();
    }

    /**
     * @param $path
     *
     * @return string
     */
    protected function getType($path)
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
     * @param bool   $bold
     * @param int    $font_size
     * @param string $font_color
     * @param bool   $wrap_text
     * @param string $background_color
     *
     * @return $this
     */

    public function setHeaderStyle($bold = false, $font_size = 12, $font_color = Color::BLACK, $wrap_text = false, $background_color = Color::YELLOW)
    {
        $this->styleHeader = (new StyleBuilder())
          ->setFontBold($bold)
          ->setFontSize($font_size)
          ->setFontColor($font_color)
          ->setShouldWrapText($wrap_text)
          ->setBackgroundColor($background_color)
          ->build();

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
    }
}

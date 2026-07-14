<?php

namespace Rap2hpoutre\FastExcel;

use Illuminate\Support\Collection;
use OpenSpout\Reader\CSV\Options as CsvReaderOptions;
use OpenSpout\Writer\Common\AbstractOptions;
use OpenSpout\Writer\CSV\Options as CsvWriterOptions;
use Traversable;

/**
 * Class FastExcel.
 */
class FastExcel
{
    use Importable;
    use Exportable;

    /**
     * @var Collection|Traversable|array
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
     * @var int|null
     */
    private $end_row = null;

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
     * @var callable|null
     */
    protected $options_configurator = null;

    /**
     * @var callable|null
     */
    protected $writer_configurator = null;

    /**
     * FastExcel constructor.
     *
     * @param array|Traversable|null $data
     */
    public function __construct(array|Traversable|null $data = null)
    {
        $this->data = $data;
    }

    /**
     * Manually set data apart from the constructor.
     *
     * @param Collection|Traversable|array $data
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
     * Limit the number of data rows imported. Pass null to remove the limit.
     *
     * @param int|null $rows
     *
     * @return $this
     */
    public function limitRows(?int $rows = null)
    {
        $this->end_row = $rows;

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
     *
     * @deprecated Has no effect with spout v4
     * @see        configureOptionsUsing
     */
    public function configureReaderUsing(?callable $callback = null)
    {
        return $this;
    }

    /**
     * Configure a custom writer factory using a callback.
     *
     * The callback receives the configured options and file extension
     * ('csv', 'ods' or 'xlsx') and should return a \OpenSpout\Writer\WriterInterface instance.
     * Return null to fall back to the default writer for that extension.
     *
     * @param callable|null $callback function (AbstractOptions $options, string $extension): ?\OpenSpout\Writer\WriterInterface
     *
     * @return $this
     */
    public function configureWriterUsing(?callable $callback = null)
    {
        $this->writer_configurator = $callback;

        return $this;
    }

    /**
     * Configure the underlying Spout Reader options using a callback.
     *
     * @param callable|null $callback
     *
     * @return $this
     */
    public function configureOptionsUsing(?callable $callback = null)
    {
        $this->options_configurator = $callback;

        return $this;
    }

    /**
     * @param AbstractOptions $options
     */
    protected function setOptions(&$options)
    {
        if ($options instanceof CsvReaderOptions || $options instanceof CsvWriterOptions) {
            $options->FIELD_DELIMITER = $this->csv_configuration['delimiter'];
            $options->FIELD_ENCLOSURE = $this->csv_configuration['enclosure'];
            if ($options instanceof CsvReaderOptions) {
                $options->ENCODING = $this->csv_configuration['encoding'];
            }
            if ($options instanceof CsvWriterOptions) {
                $options->SHOULD_ADD_BOM = $this->csv_configuration['bom'];
            }
        }

        if (is_callable($this->options_configurator)) {
            call_user_func(
                $this->options_configurator,
                $options
            );
        }
    }
}

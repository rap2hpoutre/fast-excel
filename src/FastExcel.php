<?php

namespace Rap2hpoutre\FastExcel;

use Illuminate\Support\Collection;
use InvalidArgumentException;
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
    private $header_row = null;

    /**
     * @var int|null
     */
    private $end_row = null;

    /**
     * @var int|null
     */
    private $end_column = null;

    /**
     * 1-based column indexes to keep when importing. Null means no allowlist.
     *
     * @var int[]|null
     */
    private $only_columns = null;

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
     * @param int|string $sheet_number 1-based index or sheet name
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
     * Enable passing sheet name to callback.
     *
     * @return $this
     */
    public function withSheetContext()
    {
        $this->with_sheet_context = true;

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
     * Read the headers from a specific row instead of from startRow().
     *
     * By default `startRow(155)` also takes the headers from row 155. Set the
     * header row explicitly to keep the real headers while reading data further
     * down the file, which is what makes chunked imports possible:
     *
     *     (new FastExcel)->headerRow(1)->startRow(155)->limitRows(100)->import($file);
     *
     * Pass null to go back to taking the headers from startRow().
     *
     * @param int|null $row 1-based row index
     *
     * @return $this
     */
    public function headerRow(?int $row = 1)
    {
        $this->header_row = $row;

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
     * Stop reading each row after the given column, given either as a number of
     * columns (8) or as a column reference ('H'). Pass null to remove the limit.
     * Setting a limit clears any onlyColumns() allowlist; clearing with null does not.
     *
     * @param int|string|null $column
     *
     * @return $this
     */
    public function limitColumns(int|string|null $column = null)
    {
        if ($column === null) {
            $this->end_column = null;

            return $this;
        }

        $this->only_columns = null;
        $this->end_column = $this->columnIndex($column);

        return $this;
    }

    /**
     * Keep only the given columns when importing (letters or 1-based indexes).
     * Order is preserved. Pass null to clear the allowlist.
     * Setting an allowlist clears any limitColumns() cap; clearing with null does not.
     *
     * @param array<int, int|string>|null $columns
     *
     * @return $this
     */
    public function onlyColumns(?array $columns = null)
    {
        if ($columns === null) {
            $this->only_columns = null;

            return $this;
        }

        if ($columns === []) {
            throw new InvalidArgumentException('onlyColumns() requires at least one column.');
        }

        $this->end_column = null;
        $indexes = array_values(array_map(function ($column) {
            if (!is_int($column) && !is_string($column)) {
                throw new InvalidArgumentException('onlyColumns() accepts column letters or 1-based indexes.');
            }

            return $this->columnIndex($column);
        }, $columns));

        if (count($indexes) !== count(array_unique($indexes))) {
            throw new InvalidArgumentException('onlyColumns() does not allow duplicate columns.');
        }

        $this->only_columns = $indexes;

        return $this;
    }

    /**
     * Resolve a column reference to its 1-based index: both 8 and 'H' give 8,
     * 'AA' gives 27.
     *
     * @param int|string $column
     *
     * @return int
     */
    private function columnIndex(int|string $column)
    {
        if (is_int($column) || ctype_digit($column)) {
            $index = (int) $column;
            if ($index < 1) {
                throw new InvalidArgumentException("Column reference [$column] must be greater than zero.");
            }

            return $index;
        }

        $letters = strtoupper(trim($column));
        if ($letters === '' || !ctype_alpha($letters)) {
            throw new InvalidArgumentException("Invalid column reference [$column].");
        }

        $index = 0;
        foreach (str_split($letters) as $letter) {
            $index = $index * 26 + ord($letter) - ord('A') + 1;
        }

        return $index;
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

    public function rightToLeft(bool $value = true): static
    {
        $this->right_to_left = $value;

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

<?php

namespace Rap2hpoutre\FastExcel;

use DateTimeInterface;
use Illuminate\Support\Collection;
use InvalidArgumentException;
use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Cell\StringCell;
use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Writer\Common\AbstractOptions;
use OpenSpout\Writer\WriterInterface;
use Traversable;

/**
 * Trait Exportable.
 *
 * @property bool                           $transpose
 * @property bool                           $with_header
 * @property callable|null                  $writer_configurator
 * @property \Illuminate\Support\Collection $data
 */
trait Exportable
{
    /**
     * @var Style
     */
    private $header_style;
    private $rows_style;

    /** @var Style[] */
    private $header_column_styles = [];

    /** @var Style[] */
    private $column_styles = [];

    /** @var bool */
    private $string_values = false;

    /** @var bool */
    private $escape_formulas = false;

    /** @var array<string, string> */
    private $column_formats = [];

    /**
     * @param AbstractOptions $options
     *
     * @return mixed
     */
    abstract protected function setOptions(&$options);

    /** @param Style[] $styles */
    public function setHeaderColumnStyles($styles): static
    {
        $this->header_column_styles = $styles;

        return $this;
    }

    /** @param Style[] $styles */
    public function setColumnStyles($styles): static
    {
        $this->column_styles = $styles;

        return $this;
    }

    /**
     * Export every scalar value as a string (text cell). Useful to preserve
     * leading zeros or long numeric IDs (e.g. phone numbers) that would
     * otherwise be written as numbers. Per-column setColumnFormat() rules take
     * precedence over this flag.
     *
     * @param bool $enabled
     */
    public function stringValues(bool $enabled = true): static
    {
        $this->string_values = $enabled;

        return $this;
    }

    /**
     * Write string values as explicit text cells so a value starting with "="
     * (or other formula-triggering input) is never emitted as a live formula
     * cell. Protects against CSV/formula injection and "corrupt file" errors.
     *
     * @param bool $enabled
     */
    public function escapeFormulas(bool $enabled = true): static
    {
        $this->escape_formulas = $enabled;

        return $this;
    }

    /**
     * Force the cell type for specific columns, keyed by column name, e.g.
     * ['phone' => 'string', 'price' => 'number']. A 'string' column is written
     * as text; a 'number' column casts numeric strings back to int/float. These
     * rules win over stringValues().
     *
     * @param array<string, string> $formats
     */
    public function setColumnFormat(array $formats): static
    {
        $this->column_formats = $formats;

        return $this;
    }

    /**
     * @param string        $path
     * @param callable|null $callback
     *
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     * @throws \OpenSpout\Common\Exception\IOException
     *
     * @return string
     */
    public function export($path, ?callable $callback = null)
    {
        self::exportOrDownload($path, 'openToFile', $callback);

        return realpath($path) ?: $path;
    }

    /**
     * @param               $path
     * @param callable|null $callback
     *
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     * @throws \OpenSpout\Common\Exception\IOException
     *
     * @return \Symfony\Component\HttpFoundation\StreamedResponse|string
     */
    public function download($path, ?callable $callback = null)
    {
        if (method_exists(response(), 'streamDownload')) {
            return response()->streamDownload(function () use ($path, $callback) {
                self::exportOrDownload($path, 'openToBrowser', $callback);
            }, $path);
        }
        self::exportOrDownload($path, 'openToBrowser', $callback);

        return '';
    }

    /**
     * @param               $path
     * @param string        $function
     * @param callable|null $callback
     *
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     * @throws \OpenSpout\Common\Exception\SpoutException
     */
    private function exportOrDownload($path, $function, ?callable $callback = null)
    {
        $writer = $this->makeWriter($path);

        $has_sheets = ($writer instanceof \OpenSpout\Writer\XLSX\Writer || $writer instanceof \OpenSpout\Writer\ODS\Writer);

        // CSV (and any other single-sheet format) cannot hold multiple sheets.
        // Fail with a clear message instead of a cryptic "undefined method
        // ...::getCurrentSheet()" fatal further down.
        if (!$has_sheets && $this->data instanceof SheetCollection && $this->data->count() > 1) {
            throw new InvalidArgumentException(
                'The "'.pathinfo($path, PATHINFO_EXTENSION).'" format does not support multiple sheets; use xlsx or ods.'
            );
        }

        $writer->$function($path);

        // It can export one sheet (Collection) or N sheets (SheetCollection)
        $data = $this->transpose ? $this->transposeData() : ($this->data instanceof SheetCollection ? $this->data : collect([$this->data]));

        $last_key = $data->keys()->last();

        foreach ($data as $key => $collection) {
            if ($collection instanceof Collection) {
                $this->writeRowsFromCollection($writer, $collection, $callback);
            } elseif ($collection instanceof Traversable) {
                $this->writeRowsFromGenerator($writer, $collection, $callback);
            } elseif (is_array($collection)) {
                $this->writeRowsFromArray($writer, $collection, $callback);
            } else {
                throw new InvalidArgumentException('Unsupported type for $data');
            }
            if ($has_sheets && is_string($key)) {
                $writer->getCurrentSheet()->setName($key);
            }
            if ($has_sheets && $last_key !== $key) {
                $writer->addNewSheetAndMakeItCurrent();
            }
        }
        $writer->close();
    }

    private function makeWriter(string $path): WriterInterface
    {
        $extension = $this->resolveWriterExtension($path);

        $options = match ($extension) {
            'csv'   => new \OpenSpout\Writer\CSV\Options(),
            'ods'   => new \OpenSpout\Writer\ODS\Options(),
            default => new \OpenSpout\Writer\XLSX\Options(),
        };

        $this->setOptions($options);

        if (is_callable($this->writer_configurator ?? null)) {
            $writer = call_user_func($this->writer_configurator, $options, $extension);

            if ($writer instanceof WriterInterface) {
                return $writer;
            }
        }

        return match ($extension) {
            'csv'   => new \OpenSpout\Writer\CSV\Writer($options),
            'ods'   => new \OpenSpout\Writer\ODS\Writer($options),
            default => new \OpenSpout\Writer\XLSX\Writer($options),
        };
    }

    private function resolveWriterExtension(string $path): string
    {
        if (str_ends_with($path, 'csv')) {
            return 'csv';
        }

        if (str_ends_with($path, 'ods')) {
            return 'ods';
        }

        return 'xlsx';
    }

    /**
     * Transpose data from rows to columns.
     *
     * @return SheetCollection
     */
    private function transposeData()
    {
        $data = $this->data instanceof SheetCollection ? $this->data : collect([$this->data]);
        $transposedData = [];

        foreach ($data as $key => $collection) {
            foreach ($collection as $row => $columns) {
                foreach ($columns as $column => $value) {
                    $transposedData[$key][$column][$row] = $value;
                }
            }
        }

        return new SheetCollection($transposedData);
    }

    /**
     * Write collection rows to the writer, streaming row by row.
     *
     * @SuppressWarnings(PHPMD.StaticAccess)
     */
    private function writeRowsFromCollection($writer, Collection $collection, ?callable $callback = null)
    {
        // Apply callback
        if ($callback) {
            $collection->transform(function ($value) use ($callback) {
                return $callback($value);
            });
        }

        if ($collection->isEmpty()) {
            if ($this->data instanceof SheetCollection) {
                $this->writePlaceholderRowForEmptySheet($writer);
            }

            return;
        }

        // Prepare collection (i.e remove non-string)
        $this->prepareCollection($collection);
        // Add header row.
        if ($this->with_header) {
            $this->writeHeader($writer, $collection->first());
        }

        $use_styles = $this->rows_style || $this->column_styles || $this->escape_formulas;

        // Write rows one by one so Row objects can be garbage-collected as
        // they are written, instead of materializing them all up front.
        foreach ($collection as $values) {
            // Row::fromValues works only with arrays
            if (!is_array($values)) {
                $values = $values->toArray();
            }

            if ($use_styles) {
                // Column styles are matched against the value keys; use positional
                // keys so numeric style indexes work with associative rows.
                $writer->addRow($this->createRow(array_values($values), $this->rows_style, $this->column_styles));
            } else {
                $writer->addRow(Row::fromValues($values));
            }
        }
    }

    private function writeRowsFromGenerator($writer, Traversable $generator, ?callable $callback = null)
    {
        $hasRows = false;

        foreach ($generator as $key => $item) {
            $hasRows = true;
            // Apply callback
            if ($callback) {
                $item = $callback($item);
            }

            // Prepare row (i.e remove non-string)
            $item = $this->transformRow($item);

            // Add header row.
            if ($this->with_header && $key === 0) {
                $this->writeHeader($writer, $item);
            }
            // Write rows (one by one).
            $writer->addRow($this->createRow($item, $this->rows_style, $this->column_styles));
        }

        if (!$hasRows && $this->data instanceof SheetCollection) {
            $this->writePlaceholderRowForEmptySheet($writer);
        }
    }

    private function writeRowsFromArray($writer, array $array, ?callable $callback = null)
    {
        $collection = collect($array);

        // Rows must be arrays or objects; anything else (e.g. a flat list of
        // scalars) is silently skipped, as before. Empty-sheet placeholder
        // handling is done by writeRowsFromCollection.
        if ($collection->isNotEmpty() && !is_object($collection->first()) && !is_array($collection->first())) {
            return;
        }

        $this->writeRowsFromCollection($writer, $collection, $callback);
    }

    private function writeHeader($writer, $first_row)
    {
        if ($first_row === null) {
            return;
        }

        $keys = array_keys(is_array($first_row) ? $first_row : $first_row->toArray());
        $writer->addRow($this->createRow($keys, $this->header_style, $this->header_column_styles));
    }

    /**
     * Prepare collection by removing non string if required.
     */
    protected function prepareCollection(Collection $collection)
    {
        $need_conversion = false;
        $first_row = $collection->first();

        if (!$first_row) {
            return;
        }

        foreach ($first_row as $item) {
            if (!is_string($item)) {
                $need_conversion = true;
                break;
            }
        }
        if ($need_conversion || $this->string_values || $this->column_formats) {
            $collection->transform(fn ($data) => $this->transformRow($data));
        }
    }

    /**
     * Transform one row (i.e remove non-string).
     */
    private function transformRow($data): array
    {
        $row = [];
        foreach (is_array($data) ? $data : collect($data)->all() as $key => $value) {
            $value = is_null($value) ? '' : $this->formatValue($value, $key);
            if (is_string($value) || is_int($value) || is_float($value) || $value instanceof DateTimeInterface) {
                $row[$key] = $value;
            }
        }

        return $row;
    }

    /**
     * Apply the configured export format to a single value. A per-column rule
     * (setColumnFormat) wins over the global stringValues() flag. Dates and
     * non-numeric strings are always left untouched.
     *
     * @param mixed      $value
     * @param int|string $key
     *
     * @return mixed
     */
    private function formatValue($value, $key)
    {
        $format = $this->column_formats[$key] ?? ($this->string_values ? 'string' : null);

        if ($format === 'string' && (is_int($value) || is_float($value))) {
            return (string) $value;
        }

        if ($format === 'number' && is_string($value) && is_numeric($value)) {
            return $value + 0;
        }

        return $value;
    }

    /**
     * @param Style $style
     *
     * @return Exportable
     */
    public function headerStyle(Style $style)
    {
        $this->header_style = $style;

        return $this;
    }

    /**
     * @param Style $style
     *
     * @return Exportable
     */
    public function rowsStyle(Style $style)
    {
        $this->rows_style = $style;

        return $this;
    }

    /**
     * OpenSpout multi-sheet workbooks require at least one non-empty row per worksheet.
     * Empty rows are skipped by the XLSX writer and would otherwise produce a corrupt file.
     *
     * @param \OpenSpout\Writer\WriterInterface $writer
     */
    private function writePlaceholderRowForEmptySheet($writer): void
    {
        if (!$writer instanceof \OpenSpout\Writer\XLSX\Writer && !$writer instanceof \OpenSpout\Writer\ODS\Writer) {
            return;
        }

        $writer->addRow($this->createRow([' ']));
    }

    /**
     * Create openspout row from values with optional row and cell styling.
     *
     * @SuppressWarnings(PHPMD.StaticAccess)
     */
    private function createRow(array $values = [], ?Style $rows_style = null, array $column_styles = []): Row
    {
        if (!$this->escape_formulas) {
            return Row::fromValuesWithStyles($values, $rows_style, $column_styles);
        }

        $cells = [];
        foreach ($values as $key => $value) {
            $style = $column_styles[$key] ?? null;
            $cells[] = is_string($value)
                ? new StringCell($value, $style)
                : Cell::fromValue($value, $style);
        }

        return new Row($cells, $rows_style);
    }
}

<?php

namespace Rap2hpoutre\FastExcel;

use Illuminate\Support\Collection;
use OpenSpout\Common\Entity\Cell;
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
     * @var bool
     */
    private $with_sheet_context = false;

    /**
     * @param AbstractOptions $options
     *
     * @return mixed
     */
    abstract protected function setOptions(&$options);

    /**
     * @param string|\Symfony\Component\HttpFoundation\File\UploadedFile $path
     * @param callable|null                                              $callback
     *
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Common\Exception\IOException
     *
     * @return Collection
     */
    public function import($path, ?callable $callback = null)
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
     * @param string|\Symfony\Component\HttpFoundation\File\UploadedFile $path
     * @param callable|null                                              $callback
     *
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Common\Exception\IOException
     *
     * @return Collection
     */
    public function importSheets($path, ?callable $callback = null)
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
     * @param string|object $path
     *
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Common\Exception\IOException
     *
     * @return \OpenSpout\Reader\ReaderInterface
     */
    private function reader($path)
    {
        $type = $this->readerType($path);

        if ($type === 'csv') {
            $options = new \OpenSpout\Reader\CSV\Options();
            $this->setOptions($options);
            $reader = new \OpenSpout\Reader\CSV\Reader($options);
        } elseif ($type === 'ods') {
            $options = new \OpenSpout\Reader\ODS\Options();
            $this->setOptions($options);
            $reader = new \OpenSpout\Reader\ODS\Reader($options);
        } else {
            $options = new \OpenSpout\Reader\XLSX\Options();
            $this->setOptions($options);
            $reader = new \OpenSpout\Reader\XLSX\Reader($options);
        }

        /* @var \OpenSpout\Reader\ReaderInterface $reader */
        $reader->open($this->readerPath($path));

        return $reader;
    }

    /**
     * Resolve the reader type (csv|ods|xlsx) for the given path or uploaded file.
     *
     * Files uploaded through a form are stored under an extension-less temporary
     * path (e.g. /tmp/phpXXXX), so the type cannot be guessed from the path alone.
     * In that case we rely on the uploaded file's original extension / mime type,
     * falling back to sniffing the file contents.
     *
     * @param string|object $path
     *
     * @return string
     */
    private function readerType($path): string
    {
        // Illuminate/Symfony UploadedFile (duck-typed to avoid a hard dependency).
        if (is_object($path) && method_exists($path, 'getClientOriginalExtension')) {
            $extension = strtolower((string) $path->getClientOriginalExtension());
            if (in_array($extension, ['csv', 'ods', 'xlsx'], true)) {
                return $extension;
            }

            $mime = method_exists($path, 'getClientMimeType') ? (string) $path->getClientMimeType() : '';

            return $this->readerTypeFromMime($mime, $this->readerPath($path));
        }

        $path = (string) $path;

        if (str_ends_with($path, 'csv')) {
            return 'csv';
        }
        if (str_ends_with($path, 'ods')) {
            return 'ods';
        }
        if (str_ends_with($path, 'xlsx')) {
            return 'xlsx';
        }

        // No recognizable extension (e.g. an uploaded temp file): sniff the contents.
        return $this->readerTypeFromMime($this->guessMimeType($path), $path);
    }

    /**
     * Map a mime type to a reader type, sniffing the file as a last resort.
     *
     * @param string $mime
     * @param string $path
     *
     * @return string
     */
    private function readerTypeFromMime(string $mime, string $path): string
    {
        $mime = strtolower($mime);

        if (str_contains($mime, 'csv')) {
            return 'csv';
        }
        if (str_contains($mime, 'opendocument.spreadsheet')) {
            return 'ods';
        }
        if (str_contains($mime, 'spreadsheetml') || str_contains($mime, 'officedocument')) {
            return 'xlsx';
        }

        return $this->sniffReaderType($mime, $path);
    }

    /**
     * Determine the reader type when the mime type is inconclusive by inspecting
     * the file contents.
     *
     * @param string $mime
     * @param string $path
     *
     * @return string
     */
    private function sniffReaderType(string $mime, string $path): string
    {
        // XLSX and ODS are zip archives; distinguish them by peeking inside.
        if (is_file($path) && $this->isZipArchive($path)) {
            return $this->zipContains($path, 'mimetype', 'opendocument.spreadsheet') ? 'ods' : 'xlsx';
        }

        // Plain delimited text with no zip signature: treat as CSV.
        if ($mime === '' || str_starts_with($mime, 'text/')) {
            return 'csv';
        }

        return 'xlsx';
    }

    /**
     * @param string|object $path
     *
     * @return string
     */
    private function readerPath($path): string
    {
        if (is_object($path) && method_exists($path, 'getPathname')) {
            return (string) $path->getPathname();
        }

        return (string) $path;
    }

    /**
     * @param string $path
     *
     * @return string
     */
    private function guessMimeType($path): string
    {
        if (is_file($path) && function_exists('mime_content_type')) {
            $mime = @mime_content_type($path);
            if ($mime !== false) {
                return $mime;
            }
        }

        return '';
    }

    /**
     * @param string $path
     *
     * @return bool
     */
    private function isZipArchive($path): bool
    {
        $handle = @fopen($path, 'rb');
        if ($handle === false) {
            return false;
        }
        $signature = fread($handle, 2);
        fclose($handle);

        return $signature === 'PK';
    }

    /**
     * Check whether a zip entry's content contains a given needle.
     *
     * @param string $path
     * @param string $entry
     * @param string $needle
     *
     * @return bool
     */
    private function zipContains($path, $entry, $needle): bool
    {
        if (!class_exists(\ZipArchive::class)) {
            return false;
        }

        $zip = new \ZipArchive();
        if ($zip->open($path) !== true) {
            return false;
        }
        $content = $zip->getFromName($entry);
        $zip->close();

        return is_string($content) && str_contains($content, $needle);
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
    private function importSheet(SheetInterface $sheet, ?callable $callback = null)
    {
        $headers = [];
        $collection = [];
        $count_header = 0;
        $sheetName = $sheet->getName();

        foreach ($sheet->getRowIterator() as $k => $rowAsObject) {
            $row = array_map(function (Cell $cell) {
                return match (true) {
                    $cell instanceof Cell\FormulaCell => $cell->getComputedValue(),
                    default                           => $cell->getValue(),
                };
            }, $rowAsObject->getCells());

            if ($k >= $this->start_row) {
                if ($this->with_header) {
                    if ($k == $this->start_row) {
                        $headers = $this->uniqueHeaders($this->toStrings($row));
                        $count_header = count($headers);
                        continue;
                    }
                    if ($count_header > $count_row = count($row)) {
                        $row = array_merge($row, array_fill(0, $count_header - $count_row, null));
                    } elseif ($count_header < $count_row = count($row)) {
                        $row = array_slice($row, 0, $count_header);
                    }
                }

                $rowData = empty($headers) ? $row : array_combine($headers, $row);

                if ($callback) {
                    if ($this->with_sheet_context) {
                        $result = $callback($sheetName, $rowData);
                    } else {
                        $result = $callback($rowData);
                    }

                    if ($result) {
                        $collection[] = $result;
                    }
                } else {
                    $collection[] = $rowData;
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
     * Make header names usable as array keys. Empty headers get a positional
     * name (`column_N`) and duplicated headers are de-duplicated: the first
     * occurrence is kept and later ones get a numeric suffix (`Name`, `Name_2`).
     * Without this, columns that share a name collide in array_combine() and
     * their values are silently lost.
     *
     * @param array $headers
     *
     * @return array
     */
    private function uniqueHeaders(array $headers)
    {
        $used = [];
        foreach ($headers as $index => $header) {
            $header = (string) $header;
            if ($header === '') {
                $header = 'column_'.($index + 1);
            }

            $base = $header;
            $suffix = 1;
            while (isset($used[$header])) {
                $suffix++;
                $header = $base.'_'.$suffix;
            }

            $used[$header] = true;
            $headers[$index] = $header;
        }

        return $headers;
    }
}

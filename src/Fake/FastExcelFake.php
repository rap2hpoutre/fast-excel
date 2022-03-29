<?php

namespace Rap2hpoutre\FastExcel\Fake;

use Box\Spout\Common\Entity\Style\Style;
use PHPUnit\Framework\Assert;
use Rap2hpoutre\FastExcel\Contracts\ExcelInterface;
use Rap2hpoutre\FastExcel\Contracts\ExportInterface;
use Rap2hpoutre\FastExcel\Contracts\ImportInterface;

// use Rap2hpoutre\FastExcel\SheetCollection;

class FastExcelFake implements ExcelInterface,ExportInterface,ImportInterface
{
    /**
     * @var Collection|Generator|array
     */
    protected $data;

    /**
     * @var array
     */
    protected $exports = [];

    /**
     * @var array
     */
    protected $downloads = [];

    /**
     * @var array
     */
    protected $export_callbacks = [];

    /**
     * @var array
     */
    protected $download_callbacks = [];

    /**
     * @var Style
     */
    private $header_style;
    private $rows_style;

    /**
     * FastExcel constructor.
     *
     * @param Collection|Generator|array|null $data
     */
    public function __construct($data = null)
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

    public function sheet($sheet_number)
    {
        return $this;
    }

    public function withoutHeaders()
    {
        return $this;
    }

    public function withSheetsNames()
    {
        return $this;
    }

    public function startRow(int $row)
    {
        return $this;
    }

    public function transpose()
    {
        return $this;
    }

    public function configureCsv($delimiter = ',', $enclosure = '"', $encoding = 'UTF-8', $bom = false)
    {
        return $this;
    }

    public function configureReaderUsing(?callable $callback = null)
    {
        return $this;
    }

    public function configureWriterUsing(?callable $callback = null)
    {
        return $this;
    }

    public function import($path, callable $callback = null)
    {
        return $this;
    }

    public function importSheets($path, callable $callback = null)
    {
        return $this;
    }

    public function export($path, callable $callback = null)
    {
        // It can export one sheet (Collection) or N sheets (SheetCollection)
        // $data = $this->transpose ? $this->transposeData() : ($this->data instanceof SheetCollection ? $this->data : collect([$this->data]));
        $this->exports[$path] = $this->data;

        $this->export_callbacks[$path] = $callback;

        return '';
    }

    /**
     * @param $path
     * @param callable|null $callback
     **/
    public function download($path, callable $callback = null)
    {
        // It can export one sheet (Collection) or N sheets (SheetCollection)
        // $data = $this->transpose ? $this->transposeData() : ($this->data instanceof SheetCollection ? $this->data : collect([$this->data]));
        $this->downloads[$path] = $this->data;

        $this->download_callbacks[$path] = $callback;

        return '';
    }

    /**
     * @param string        $fileName
     * @param callable|null $callback
     */
    public function assertDownloaded(string $fileName, $callback = null)
    {
        Assert::assertArrayHasKey($fileName, $this->downloads, sprintf('%s is not downloaded', $fileName));

        $callback = $callback ?: function () {
            return true;
        };

        Assert::assertTrue(
            $callback($this->downloads[$fileName]),
            "The file [{$fileName}] was not downloaded with the expected data."
        );
    }

    public function headerStyle(Style $style)
    {
        $this->header_style = $style;

        return $this;
    }

    public function rowsStyle(Style $style)
    {
        $this->rows_style = $style;

        return $this;
    }
}

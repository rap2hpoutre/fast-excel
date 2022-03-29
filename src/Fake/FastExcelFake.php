<?php

namespace Rap2hpoutre\FastExcel\Fake;

use PHPUnit\Framework\Assert;
use Rap2hpoutre\FastExcel\SheetCollection;

class FastExcelFake
{
    /**
     * @var Collection|Generator|array
     */
    protected $data;

    /**
     * @var array
     */
    protected $downloads = [];

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

    /**
     * @param $path
     * @param callable|null $callback
     **/
    public function download($path, callable $callback = null)
    {
        // It can export one sheet (Collection) or N sheets (SheetCollection)
        // $data = $this->transpose ? $this->transposeData() : ($this->data instanceof SheetCollection ? $this->data : collect([$this->data]));
        $this->downloads[$path] = $this->data;

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

}

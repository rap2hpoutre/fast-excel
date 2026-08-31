<?php

namespace Rap2hpoutre\FastExcel\Tests;

use PHPUnit\Framework\TestCase as BaseTestCase;

/**
 * Class TestCase.
 */
class TestCase extends BaseTestCase
{
    /**
     * Temporary files created during a test; removed in tearDown().
     *
     * @var string[]
     */
    private $tempFiles = [];

    protected function tearDown(): void
    {
        foreach ($this->tempFiles as $file) {
            if (is_file($file)) {
                @unlink($file);
            }
        }
        $this->tempFiles = [];

        parent::tearDown();
    }

    /**
     * @return \Illuminate\Support\Collection
     */
    protected function collection()
    {
        return collect([
            ['col1' => 'row1 col1', 'col2' => 'row1 col2'],
            ['col1' => 'row2 col1', 'col2' => ''],
            ['col1' => 'row3 col1', 'col2' => 'row3 col2'],
        ]);
    }

    /**
     * Path under the system temp dir, tracked for automatic cleanup.
     *
     * @param string $prefix
     *
     * @return string
     */
    protected function tempXlsx(string $prefix = 'test'): string
    {
        $path = sys_get_temp_dir().'/fastexcel-'.$prefix.'-'.uniqid('', true).'.xlsx';
        $this->tempFiles[] = $path;

        return $path;
    }
}

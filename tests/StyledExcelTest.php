<?php

namespace Rap2hpoutre\FastExcel\Tests;

use Illuminate\Support\Collection;
use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Common\Entity\Style\Color;
use OpenSpout\Common\Entity\Style\Style;
use Rap2hpoutre\FastExcel\FastExcel;

/**
 * Class FastExcelTest.
 */
class StyledExcelTest extends TestCase
{
    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     */
    public function testExportWithRowsStyle()
    {
        $original_collection = $this->collection();

        $style = (new Style())
            ->setFontItalic()
            ->setFontSize(15)
            ->setFontColor(Color::BLUE)
            ->setShouldWrapText()
            ->setBackgroundColor(Color::YELLOW);

        $file = __DIR__.'/test-row-style.xlsx';

        (new FastExcel(clone $original_collection))
            ->rowsStyle($style)
            ->export($file);

        $this->assertEquals($original_collection, (new FastExcel())->import($file));

        unlink($file);
    }

    /**
     * @throws \OpenSpout\Common\Exception\IOException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     */
    public function testExportWithCellsStyle()
    {
        $original_collection = $this->collection();

        $collection = $this->applyCellStyles(clone $original_collection);

        $file = __DIR__.'/test-cells-style.xlsx';

        (new FastExcel($collection))->export($file);

        $this->assertEquals($original_collection, (new FastExcel())->import($file));

        unlink($file);
    }

    /**
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     */
    protected function applyCellStyles(Collection $data): Collection
    {
        $styles = [
            [
                'col1' => (new Style())
                    ->setFontBold()
                    ->setFontColor(Color::GREEN)
                    ->setShouldWrapText()
                    ->setBackgroundColor(Color::YELLOW),

                'col2' => (new Style())
                    ->setFontColor(Color::DARK_BLUE)
                    ->setBackgroundColor(Color::LIGHT_GREEN),
            ],

            [
                'col1' => (new Style())
                    ->setCellAlignment(CellAlignment::CENTER)
                    ->setBackgroundColor(Color::YELLOW),

                'col2' => (new Style())
                    ->setCellAlignment(CellAlignment::RIGHT)
                    ->setFontColor(Color::DARK_BLUE),
            ],

            [
                'col1' => (new Style())
                    ->setBackgroundColor(Color::ORANGE),

                'col2' => (new Style())
                    ->setCellAlignment(CellAlignment::RIGHT)
                    ->setFontColor(Color::DARK_BLUE)
                    ->setBackgroundColor(Color::LIGHT_GREEN),
            ],

        ];

        $data->transform(function ($row, $row_index) use ($styles) {
            return collect($row)->transform(function ($value, $col_index) use ($row_index, $styles) {
                return Cell::fromValue($value, $styles[$row_index][$col_index] ?? null);
            });
        });

        return $data;
    }
}

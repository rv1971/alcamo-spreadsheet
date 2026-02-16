<?php

namespace alcamo\spreadsheet;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\{Alignment, Font};
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PHPUnit\Framework\TestCase;

class MyWorksheet extends Worksheet
{
    public const PAGE_PAPERSIZE = PageSetup::PAPERSIZE_A5;
    public const PAGE_ORIENTATION = PageSetup::ORIENTATION_LANDSCAPE;
    public const PAGE_FIT_TO_WIDTH = true;
}

class WorksheetTest extends TestCase
{
    private static $spreadsheet_;

    public static function setUpBeforeClass(): void
    {
        self::$spreadsheet_ = new Spreadsheet();

        self::$spreadsheet_->createSheet();

        self::$spreadsheet_->addSheet(
            new MyWorksheet(self::$spreadsheet_, 'My sheet')
        );
    }

    public static function tearDownAfterClass(): void
    {
        // for debugging when needed

        /*
        (new Xlsx(self::$spreadsheet_))->save(
            __DIR__ . DIRECTORY_SEPARATOR . 'WorksheetTest.xlsx'
        );
        */
    }

    public function testSetLayout()
    {
        $sheet0 = self::$spreadsheet_->getSheet(0);
        $sheet1 = self::$spreadsheet_->getSheet(1);

        $this->assertSame('Worksheet', $sheet0->getTitle());

        $this->assertSame(
            Worksheet::PAGE_PAPERSIZE,
            $sheet0->getPageSetup()->getPaperSize()
        );

        $this->assertSame(
            Worksheet::PAGE_ORIENTATION,
            $sheet0->getPageSetup()->getOrientation()
        );

        $this->assertSame(
            Worksheet::PAGE_FIT_TO_WIDTH,
            (bool)$sheet0->getPageSetup()->getFitToWidth()
        );

        $this->assertSame('My sheet', $sheet1->getTitle());

        $this->assertSame(
            MyWorksheet::PAGE_PAPERSIZE,
            $sheet1->getPageSetup()->getPaperSize()
        );

        $this->assertSame(
            MyWorksheet::PAGE_ORIENTATION,
            $sheet1->getPageSetup()->getOrientation()
        );

        $this->assertSame(
            MyWorksheet::PAGE_FIT_TO_WIDTH,
            (bool)$sheet1->getPageSetup()->getFitToWidth()
        );
    }

    public function testGetSetColRow()
    {
        $sheet0 = self::$spreadsheet_->getSheet(0);

        $this->assertSame('A', (string)$sheet0->getCol());
        $this->assertSame(1, $sheet0->getRow());
        $this->assertSame('A1', $sheet0->getCurrentCell()->getCoordinate());

        $sheet0->setCol('AA');
        $this->assertSame('AA', (string)$sheet0->getCol());
        $this->assertSame('AA1', $sheet0->getCurrentCell()->getCoordinate());

        $sheet0->setRow(42);
        $this->assertSame(42, $sheet0->getRow());
        $this->assertSame('AA42', $sheet0->getCurrentCell()->getCoordinate());

        $sheet0->setColRow('AZ', 43);
        $this->assertSame('AZ', (string)$sheet0->getCol());
        $this->assertSame(43, $sheet0->getRow());
        $this->assertSame('AZ43', $sheet0->getCurrentCell()->getCoordinate());

        $sheet0->setCoordinate('B24');
        $this->assertSame('B', (string)$sheet0->getCol());
        $this->assertSame(24, $sheet0->getRow());
        $this->assertSame('B24', $sheet0->getCurrentCell()->getCoordinate());
    }

    public function testMoveCol()
    {
        $sheet0 = self::$spreadsheet_->getSheet(0);

        $sheet0->setCol('CU')->moveCol(7);

        $this->assertSame('DB', (string)$sheet0->getCol());

        $sheet0->moveCol(-12);
        $this->assertSame('CP', (string)$sheet0->getCol());
    }

    public function testMoveRow()
    {
        $sheet0 = self::$spreadsheet_->getSheet(0);

        $sheet0->setRow(42)->moveRow(42);
        $this->assertSame(84, $sheet0->getRow());

        $sheet0->moveRow(-3);
        $this->assertSame(81, $sheet0->getRow());
    }

    public function testWriteCell()
    {
        $sheet0 = self::$spreadsheet_->getSheet(0);

        $sheet0->setColRow('B', 2)->writeCell('Lorem ipsum');

        $this->assertSame(
            'Lorem ipsum',
            $sheet0->getCell('B2')->getValue()
        );

        $sheet0->writeCell(
            1,
            [ 'font' => [ 'bold' => true ] ],
            DataType::TYPE_BOOL,
            'B3'
        );

        $this->assertSame(
            true,
            $sheet0->getCell('B3')->getValue()
        );

        $this->assertSame(
            true,
            $sheet0->getCell('B3')->getStyle()->getFont()->getBold()
        );
    }

    public function testWriteRow()
    {
        $sheet0 = self::$spreadsheet_->getSheet(0);

        $sheet0->setCoordinate('D2')->writeRow(
            [
                'foo',
                'bar' => [ 'bar' ],
                [ 'baz', 'style' => [ 'font' => [ 'size' => 18 ] ] ],
                44 => [ 44, 'type' => DataType::TYPE_STRING ]
            ],
            [ 'font' => [ 'italic' => true ] ]
        );

        $this->assertSame("D3", $sheet0->getCoordinate());

        $this->assertSame(
            'D3',
            $sheet0->getCoordinate()
        );

        $this->assertSame(
            'foo',
            $sheet0->getCell('D2')->getValue()
        );

        $this->assertSame(
            true,
            $sheet0->getCell('D2')->getStyle()->getFont()->getItalic()
        );

        $this->assertSame(
            'bar',
            $sheet0->getCell('E2')->getValue()
        );

        $this->assertSame(
            'baz',
            $sheet0->getCell('F2')->getValue()
        );

        $this->assertSame(
            18.0,
            $sheet0->getCell('F2')->getStyle()->getFont()->getSize()
        );

        $this->assertSame(
            '44',
            $sheet0->getCell('G2')->getValue()
        );

        $this->assertSame(
            true,
            $sheet0->getCell('D2')->getStyle()->getFont()->getItalic()
        );

        $sheet0->writeRow([ 'qux' ], null, null, 3);

        $this->assertSame(
            'qux',
            $sheet0->getCell('D3')->getValue()
        );

        $this->assertSame(
            'D3',
            $sheet0->getCoordinate()
        );
    }

    public function testWriteCol()
    {
        $sheet0 = self::$spreadsheet_->getSheet(0);

        $sheet0->setCoordinate('D4')->writeCol(
            [
                'qux',
                [
                    'quux',
                    'style' => [
                        'alignment' => [
                            'horizontal' => Alignment::HORIZONTAL_RIGHT
                        ]
                    ]
                ]
            ],
            [ 'font' => [ 'bold' => true ] ]
        );

        $this->assertSame("E4", $sheet0->getCoordinate());

        $this->assertSame(
            'E4',
            $sheet0->getCoordinate()
        );

        $this->assertSame(
            'qux',
            $sheet0->getCell('D4')->getValue()
        );

        $this->assertSame(
            true,
            $sheet0->getCell('D4')->getStyle()->getFont()->getBold()
        );

        $this->assertSame(
            'quux',
            $sheet0->getCell('D5')->getValue()
        );

        $this->assertSame(
            true,
            $sheet0->getCell('D5')->getStyle()->getFont()->getBold()
        );

        $this->assertSame(
            Alignment::HORIZONTAL_RIGHT,
            $sheet0->getCell('D5')->getStyle()->getAlignment()->getHorizontal()
        );

        $sheet0->writeCol([ 'corge' ], null, 'E');

        $this->assertSame(
            'corge',
            $sheet0->getCell('E4')->getValue()
        );

        $this->assertSame(
            'E4',
            $sheet0->getCoordinate()
        );
    }

    public function testFormatRows()
    {
        $sheet0 = self::$spreadsheet_->getSheet(0);

        $sheet0->setCoordinate('B7')->formatRows(
            [
                [ 'alignment' => [ 'vertical' => Alignment::VERTICAL_CENTER ] ],
                [ 'font' => [ 'underline' => Font::UNDERLINE_DOUBLE ] ]
            ],
            'D',
            'F'
        );

        $this->assertSame(
            'B9',
            $sheet0->getCoordinate()
        );

        $this->assertSame(
            Alignment::VERTICAL_CENTER,
            $sheet0->getCell('D7')->getStyle()->getAlignment()->getVertical()
        );

        $this->assertSame(
            Alignment::VERTICAL_CENTER,
            $sheet0->getCell('E7')->getStyle()->getAlignment()->getVertical()
        );

        $this->assertSame(
            Alignment::VERTICAL_CENTER,
            $sheet0->getCell('F7')->getStyle()->getAlignment()->getVertical()
        );

        $this->assertSame(
            Font::UNDERLINE_DOUBLE,
            $sheet0->getCell('D8')->getStyle()->getFont()->getUnderline()
        );

        $this->assertSame(
            Font::UNDERLINE_DOUBLE,
            $sheet0->getCell('E8')->getStyle()->getFont()->getUnderline()
        );

        $this->assertSame(
            Font::UNDERLINE_DOUBLE,
            $sheet0->getCell('F8')->getStyle()->getFont()->getUnderline()
        );

        $sheet0->formatRows(
            [
                [ 'alignment' => [ 'vertical' => Alignment::VERTICAL_CENTER ] ],
                [ 'font' => [ 'underline' => Font::UNDERLINE_DOUBLE ] ]
            ],
            'D',
            'F',
            9
        );

        $this->assertSame(
            'B9',
            $sheet0->getCoordinate()
        );
    }

    public function testFormatCols()
    {
        $sheet0 = self::$spreadsheet_->getSheet(0);

        $sheet0->setCoordinate('B10')->formatCols(
            [
                [ 'font' => [ 'size' => 32 ] ],
                [ 'font' => [ 'size' => 33 ] ],
                [ 'font' => [ 'size' => 34 ] ],
            ],
            '11',
            '15'
        );

        $this->assertSame(
            'E10',
            $sheet0->getCoordinate()
        );

        $this->assertSame(
            32.0,
            $sheet0->getCell('B11')->getStyle()->getFont()->getSize()
        );

        $this->assertSame(
            32.0,
            $sheet0->getCell('B12')->getStyle()->getFont()->getSize()
        );

        $this->assertSame(
            32.0,
            $sheet0->getCell('B13')->getStyle()->getFont()->getSize()
        );

        $this->assertSame(
            32.0,
            $sheet0->getCell('B14')->getStyle()->getFont()->getSize()
        );

        $this->assertSame(
            32.0,
            $sheet0->getCell('B15')->getStyle()->getFont()->getSize()
        );

        $this->assertSame(
            33.0,
            $sheet0->getCell('C11')->getStyle()->getFont()->getSize()
        );

        $this->assertSame(
            33.0,
            $sheet0->getCell('C12')->getStyle()->getFont()->getSize()
        );

        $this->assertSame(
            33.0,
            $sheet0->getCell('C13')->getStyle()->getFont()->getSize()
        );

        $this->assertSame(
            33.0,
            $sheet0->getCell('C14')->getStyle()->getFont()->getSize()
        );

        $this->assertSame(
            33.0,
            $sheet0->getCell('C15')->getStyle()->getFont()->getSize()
        );

        $this->assertSame(
            34.0,
            $sheet0->getCell('D11')->getStyle()->getFont()->getSize()
        );

        $this->assertSame(
            34.0,
            $sheet0->getCell('D12')->getStyle()->getFont()->getSize()
        );

        $this->assertSame(
            34.0,
            $sheet0->getCell('D13')->getStyle()->getFont()->getSize()
        );

        $this->assertSame(
            34.0,
            $sheet0->getCell('D14')->getStyle()->getFont()->getSize()
        );

        $this->assertSame(
            34.0,
            $sheet0->getCell('D15')->getStyle()->getFont()->getSize()
        );

        $sheet0->formatCols(
            [
                [ 'font' => [ 'size' => 32 ] ],
                [ 'font' => [ 'size' => 33 ] ],
                [ 'font' => [ 'size' => 34 ] ],
            ],
            '11',
            '15',
            'E'
        );

        $this->assertSame(
            'E10',
            $sheet0->getCoordinate()
        );
    }
}

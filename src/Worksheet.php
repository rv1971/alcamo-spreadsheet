<?php

namespace alcamo\spreadsheet;

use alcamo\exception\SyntaxError;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet as PhpOfficeWorksheet;

/// Worksheet which has a current position and convenience methods
class Worksheet extends PhpOfficeWorksheet
{
    public const PAGE_PAPERSIZE = PageSetup::PAPERSIZE_A4;
    public const PAGE_ORIENTATION = PageSetup::ORIENTATION_PORTRAIT;
    public const PAGE_FIT_TO_HEIGHT = false;
    public const PAGE_FIT_TO_PAGE = false;
    public const PAGE_FIT_TO_WIDTH = false;

    public const ODD_HEADER = '&L&F&R&A';
    public const ODD_FOOTER = '&L&D &T&R&P / &N';

    public const SHOW_GRIDLINES = false;

    private $col_; ///< Col
    private $row_; ///< int

    public function __construct(
        ?Spreadsheet $parent,
        ?string $title = null
    ) {
        parent::__construct($parent, $title ?? 'Worksheet');

        $this->col_ = Col::newFromString('A');
        $this->row_ = 1;

        $this->setLayout();
    }

    public function setLayout(): void
    {
        $this->getPageSetup()
            ->setPaperSize(static::PAGE_PAPERSIZE)
            ->setOrientation(static::PAGE_ORIENTATION)
            ->setFitToWidth(static::PAGE_FIT_TO_HEIGHT)
            ->setFitToWidth(static::PAGE_FIT_TO_PAGE)
            ->setFitToWidth(static::PAGE_FIT_TO_WIDTH);

        $this->getHeaderFooter()
            ->setOddHeader(static::ODD_HEADER)
            ->setOddFooter(static::ODD_FOOTER);

        $this->setShowGridlines(static::SHOW_GRIDLINES);
    }

    public function getCol(): Col
    {
        return $this->col_;
    }

    public function getRow(): int
    {
        return $this->row_;
    }

    public function getCoordinate(): string
    {
        return $this->col_ . $this->row_;
    }

    public function setCol($col): self
    {
        $this->col_ = $col instanceof Col
            ? $col
            : Col::newFromString($col);

        return $this;
    }

    public function setRow(int $row): self
    {
        $this->row_ = $row;

        return $this;
    }

    public function setColRow($col, int $row): self
    {
        $this->col_ = $col instanceof Col
            ? $col
            : Col::newFromString($col);
        $this->row_ = $row;

        return $this;
    }

    public function setCoordinate(string $coordinate): self
    {
        $len = strcspn($coordinate, '0123456789');

        $this->col_ = Col::newFromString(substr($coordinate, 0, $len));
        $this->row_ = (int)substr($coordinate, $len);

        return $this;
    }

    public function moveCol(?int $distance = 1): self
    {
        $this->col_ = $this->col_->add($distance);
        return $this;
    }

    public function moveRow(?int $distance = 1): self
    {
        $this->row_ += $distance;
        return $this;
    }

    /**
     * @param $coordinate The cell coordinate. If not given, @ref $col_
     * and @ref $row_ are used.
     *
     * @return $this
     *
     * @invariant Does not modify @ref $col_ nor @ref $row_.
     */
    public function writeCell(
        $value,
        ?array $style = null,
        ?string $type = null,
        string $coordinate = null
    ): self {
        $cell = $this->getCell($coordinate ?? $this->col_ . $this->row_);

        if (isset($type)) {
            $cell->setValueExplicit($value, $type);
        } else {
            $cell->setValue($value);
        }

        if (isset($style)) {
            $cell->getStyle()->applyFromArray($style);
        }

        return $this;
    }

    /**
     * @brief Write a horizontal range of cells.
     *
     * @param $rowData Array of items, each of which is
     * - either a non-array value
     * - or an array consisting of the value (with key 0) and optionally a
     * style with key `style` and/or a data type with key `type`.
     *
     * @param $rowStyle Style data applied to the cells written.
     *
     * @return $this
     *
     * @invariant Does not modify @ref $col_.
     *
     * Increments @ref $row_.
     */
    public function writeRow(array $rowData, array $rowStyle = null): self
    {
        $col = $this->col_;

        foreach ($rowData as $cellData) {
            $cellData = (array)$cellData;

            $this->writeCell(
                $cellData[0],
                $cellData['style'] ?? null,
                $cellData['type'] ?? null,
                $col . $this->row_
            );
            $col = $col->inc();
        }

        if (isset($rowStyle)) {
            $col = $col->dec();

            $this->getStyle("{$this->col_}{$this->row_}:{$col}{$this->row_}")
                ->applyFromArray($rowStyle);
        }

        $this->row_++;

        return $this;
    }

    /**
     * @brief Write a vertical range of cells.
     *
     * @param $colData Array of items, each of which is
     * - either a non-array value
     * - or an array consisting of the value (with key 0) and optionally a
     * style with key `style` and/or a data type with key `type`.
     *
     * @param $colStyle Style data applied to the cells written.
     *
     * @return $this
     *
     * @invariant Does not modify @ref $row_.
     *
     * Increments @ref $col_.
     */
    public function writeCol(array $colData, array $colStyle = null): self
    {
        $row = $this->row_;

        foreach ($colData as $cellData) {
            $cellData = (array)$cellData;

            $this->writeCell(
                $cellData[0],
                $cellData['style'] ?? null,
                $cellData['type'] ?? null,
                $this->col_ . $row++
            );
        }

        if (isset($colStyle)) {
            $row--;

            $this->getStyle("{$this->col_}{$this->row_}:{$this->col_}{$row}")
                ->applyFromArray($colStyle);
        }

        $this->col_ = $this->col_->inc();

        return $this;
    }

    /**
     * @brief Apply styles to successive horizontal ranges.
     *
     * @param $styles Array of styles, one per rows. Indexes are irrelevant.
     *
     * @return $this
     *
     * @invariant Does not modify @ref $col_.
     *
     * Apply styles to rows starting as @ref $row_. Increment @ref $row_ for
     * each entry.
     */
    public function formatRows(
        array $styles,
        $fromCol,
        $toCol
    ): self {
        foreach ($styles as $style) {
            if (isset($style)) {
                $this->getStyle("$fromCol{$this->row_}:$toCol{$this->row_}")
                    ->applyFromArray($style);
            }

            $this->row_++;
        }

        return $this;
    }

    /**
     * @brief Apply styles to successive vertical ranges.
     *
     * @param $styles Array of styles, one per column. Indexes are irrelevant.
     *
     * @return $this
     *
     * @invariant Does not modify @ref $row_.
     *
     * Apply styles to columns starting as @ref $col_. Increment @ref $col_
     * for each entry.
     */
    public function formatCols(
        array $styles,
        int $fromRow,
        int $toRow
    ): self {
        foreach ($styles as $style) {
            if (isset($style)) {
                $this->getStyle("{$this->col_}$fromRow:{$this->col_}$toRow")
                    ->applyFromArray($style);
            }

            $this->col_ = $this->col_->inc();
        }

        return $this;
    }
}

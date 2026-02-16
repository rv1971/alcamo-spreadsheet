<?php

namespace alcamo\spreadsheet;

use alcamo\exception\SyntaxError;
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Worksheet\{
    ColumnCellIterator,
    PageSetup,
    RowCellIterator,
    Worksheet as PhpOfficeWorksheet
};

/**
 * @brief Worksheet which has a current position and some convenience methods
 *
 * @date Last reviewed 2026-01-15
 */
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

        /** Set current position to top left cell. */
        $this->col_ = Col::newFromString('A');
        $this->row_ = 1;

        /** Call setLayout(). */
        $this->setLayout();
    }

    /// Set worksheet layout from the above class constants
    public function setLayout(): void
    {
        $this->getPageSetup()
            ->setPaperSize(static::PAGE_PAPERSIZE)
            ->setOrientation(static::PAGE_ORIENTATION)
            ->setFitToHeight(static::PAGE_FIT_TO_HEIGHT)
            ->setFitToPage(static::PAGE_FIT_TO_PAGE)
            ->setFitToWidth(static::PAGE_FIT_TO_WIDTH);

        $this->getHeaderFooter()
            ->setOddHeader(static::ODD_HEADER)
            ->setOddFooter(static::ODD_FOOTER);

        $this->setShowGridlines(static::SHOW_GRIDLINES);
    }

    /// Get column of current position
    public function getCol(): Col
    {
        return $this->col_;
    }

    /// Get row of current position
    public function getRow(): int
    {
        return $this->row_;
    }

    /// Get current position
    public function getCoordinate(): string
    {
        return "{$this->col_}{$this->row_}";
    }

    /**
     * @brief Set column of current position
     *
     * @param $col Col|string New column.
     */
    public function setCol($col): self
    {
        $this->col_ = $col instanceof Col
            ? $col
            : Col::newFromString($col);

        return $this;
    }

    /// Set row of current position
    public function setRow(int $row): self
    {
        $this->row_ = $row;

        return $this;
    }

    /**
     * @brief Set column and row of current position
     *
     * @param $col Col|string New column.
     *
     * @param $row New row.
     */
    public function setColRow($col, int $row): self
    {
        $this->col_ = $col instanceof Col
            ? $col
            : Col::newFromString($col);

        $this->row_ = $row;

        return $this;
    }

    /// Set current position from coordinate string
    public function setCoordinate(string $coordinate): self
    {
        $len = strcspn($coordinate, '0123456789');

        $this->col_ = Col::newFromString(substr($coordinate, 0, $len));
        $this->row_ = (int)substr($coordinate, $len);

        return $this;
    }

    /// Move current column by indicated distance
    public function moveCol(?int $distance = 1): self
    {
        $this->col_ = $this->col_->add($distance);
        return $this;
    }

    /// Move current row by indicated distance
    public function moveRow(?int $distance = 1): self
    {
        $this->row_ += $distance;
        return $this;
    }

    /// Cell at current position
    public function getCurrentCell(): Cell
    {
        return parent::getCell("{$this->col_}{$this->row_}");
    }

    /**
     * @param $value Value to write to the cell.
     *
     * @param $style Array of style data to apply.
     *
     * @param $type Data type to set explicitly, if given.
     *
     * @param Cell|string $cell The cell to write to, or its coordinates. If
     * not given, use the cell at the current position.
     *
     * @return $this
     *
     * @invariant Does not modify the current position.
     */
    public function writeCell(
        $value,
        ?array $style = null,
        ?string $type = null,
        $cell = null
    ): self {
        switch (true) {
            case !isset($cell):
                $cell = $this->getCurrentCell();
                break;

            case !($cell instanceof Cell):
                $cell = $this->getCell($cell);
                break;
        }

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
     * @brief Write a data to cells.
     *
     * @param $rowData Array of items, each of which is
     * - either a non-array value
     * - or an array consisting of the value (with key 0) and optionally a
     * style with key `style` and/or a data type with key `type`.
     *
     * @param iterable $cells An iterator whose values are of type Cell.
     *
     * @return $this
     */
    public function writeCellIterator(array $data, iterable $cells): self
    {
        $data = array_values($data);

        $i = 0;

        foreach ($cells as $cell) {
            if (!isset($data[$i])) {
                break;
            }

            $cellData = $data[$i++];

            if (is_array($cellData)) {
                $this->writeCell(
                    $cellData[0],
                    $cellData['style'] ?? null,
                    $cellData['type'] ?? null,
                    $cell
                );
            } else {
                $cell->setValue($cellData);
            }
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
     * @param $style Style data applied to the cells written.
     *
     * @param $col Col|string Column to start at, default current column.
     *
     * @param $row Row to start at, default current row.
     *
     * @return $this
     *
     * @invariant Does not modify @ref $col_.
     *
     * Increments @ref $row_ if $row is not given.
     */
    public function writeRow(
        array $data,
        ?array $style = null,
        $col = null,
        ?int $row = null
    ): self {
        if (!($col instanceof Col)) {
            $col = isset($col) ? Col::newFromString($col) : $this->col_;
        }

        if (!isset($row)) {
            $row = $this->row_;
            $this->row_++;
        }

        if (isset($style)) {
            $col2 = $col->add(count($data) - 1);

            $this->getStyle("$col$row:$col2$row")->applyFromArray($style);
        }

        return $this->writeCellIterator(
            $data,
            new RowCellIterator(
                $this,
                $row,
                (string)$col,
                (string)($col->add(count($data) - 1))
            )
        );
    }

    /**
     * @brief Write a vertical range of cells.
     *
     * @param $data Array of items, each of which is
     * - either a non-array value
     * - or an array consisting of the value (with key 0) and optionally a
     * style with key `style` and/or a data type with key `type`.
     *
     * @param $style Style data applied to the cells written.
     *
     * @param $col Col|string Column to start at, default current column.
     *
     * @param $row Row to start at, default current row.
     *
     * @return $this
     *
     * @invariant Does not modify @ref $row_.
     *
     * Increments @ref $col_ if $col is not given..
     */
    public function writeCol(
        array $data,
        ?array $style = null,
        $col = null,
        ?int $row = null
    ): self {
        if (isset($col)) {
            if (!($col instanceof Col)) {
                $col = Col::newFromString($col);
            }
        } else {
            $col = clone $this->col_;
            $this->col_ = $this->col_->inc();
        }

        if (!isset($row)) {
            $row = $this->row_;
        }

        if (isset($style)) {
            $row2 = $row + count($data) - 1;

            $this->getStyle("$col$row:$col$row2")->applyFromArray($style);
        }

        return $this->writeCellIterator(
            $data,
            new ColumnCellIterator(
                $this,
                (string)$col,
                $row,
                $row + count($data) - 1
            )
        );
    }

    /**
     * @brief Apply styles to successive horizontal ranges.
     *
     * @param $styles Array of styles, one per row. Indexes are irrelevant.
     *
     * @param $fromCol Col|string First column to apply styles to.
     *
     * @param $toCol Col|string Last column to apply styles to.
     *
     * @param $row Row to start at, default current row.
     *
     * @return $this
     *
     * @invariant Does not modify @ref $col_.
     *
     * Increments @ref $row_ by the numer of rows processed, if $row is not
     * given.
     */
    public function formatRows(
        array $styles,
        $fromCol,
        $toCol,
        ?int $row = null
    ): self {
        if (!isset($row)) {
            $row = $this->row_;

            $this->row_ += count($styles);
        }

        foreach ($styles as $style) {
            if (isset($style)) {
                $this->getStyle("$fromCol$row:$toCol$row")
                    ->applyFromArray($style);
            }

            $row++;
        }

        return $this;
    }

    /**
     * @brief Apply styles to successive vertical ranges.
     *
     * @param $styles Array of styles, one per column. Indexes are irrelevant.
     *
     * @param $fromRow  First row to apply styles to.
     *
     * @param $toRow Last row to apply styles to.
     *
     * @param $col Col|string|null column start at, default current column.
     *
     * @return $this
     *
     * @invariant Does not modify @ref $row_.
     *
     * Increments @ref $col_ by the numer of columns processed, if $col is not
     * given.
     */
    public function formatCols(
        array $styles,
        int $fromRow,
        int $toRow,
        $col = null
    ): self {
        if (!isset($col)) {
            $col = clone $this->col_;

            $this->col_ = $this->col_->add(count($styles));
        } elseif (!($col instanceof Col)) {
            $col = Col::newFromString($col);
        }

        foreach ($styles as $style) {
            if (isset($style)) {
                $this->getStyle("$col$fromRow:$col$toRow")
                    ->applyFromArray($style);
            }

            $col = $col->inc();
        }

        return $this;
    }
}

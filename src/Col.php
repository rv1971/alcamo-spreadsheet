<?php

namespace alcamo\spreadsheet;

use alcamo\exception\{SyntaxError, Underflow};
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

/**
 * @brief Identifier for a worksheet column
 *
 * @invariant Immutable object.
 *
 * @date Last reviewed 2026-01-14
 */
class Col
{
    /// Regfeulare expression for a valid identifier
    public const VALID_COL_REGEXP = '/^[A-Z]{1,3}$/';

    private $value_;

    public static function newFromString(string $value): self
    {
        if (!preg_match(static::VALID_COL_REGEXP, $value)) {
            /** @throw alcamo::exception::SyntaxError if $value does not match
             *  alcamo::spreadsheet::Col::VALID_COL_REGEXP. */
            throw (new SyntaxError())->setMessageContext(
                [ 'inData' => $value ]
            );
        }

        return new self($value);
    }

    private function __construct(string $value)
    {
        $this->value_ = $value;
    }

    public function __toString()
    {
        return $this->value_;
    }

    /// Return a new object with incremented value
    public function inc(): self
    {
        $value = $this->value_;

        return new self(++$value);
    }

    /// Return a new object with decremented value
    public function dec(): self
    {
        return new self($this->decString($this->value_));
    }

    /// Return a new object with $diff added
    public function add(int $diff): self
    {
        /* Choose the fastest approach depending on the distance. */
        switch (true) {
            case $diff == 0:
                return clone $this;

            case $diff == 1:
                $value = $this->value_;

                return new self(++$value);

            case $diff > 0:
                $value = $this->value_;

                for ($i = 0; $i < $diff; $i++) {
                    $value++;
                }

                return new self($value);

            case $diff == -1:
                return new self($this->decString($this->value_));

            case $diff == -2:
                return new self(
                    $this->decString($this->decString($this->value_))
                );

                /* $diff < 0 */
            default:
                return new self(
                    Coordinate::stringFromColumnIndex(
                        Coordinate::columnIndexFromString($this->value_)
                            + $diff
                    )
                );
        }
    }

    /// Decrement a string representing a column
    private function decString(string $value): string
    {
        if ($value == 'A') {
            /** @throw alcamo::exception::Underflow on attempt to decrement
             *  the first column. */
            throw new Underflow();
        }

        for ($i = strlen($value) - 1; $i >= 0; $i--) {
            if ($value[$i] != 'A') {
                $value[$i] = chr(ord($value[$i]) - 1);
                return $value;
            } else {
                if ($i > 0) {
                    $value[$i] = 'Z';
                } else {
                    /* This is the step from n*A to (n-1)*Z */
                    return substr($value, 1);
                }
            }
        }
    }
}

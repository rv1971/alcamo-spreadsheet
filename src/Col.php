<?php

namespace alcamo\spreadsheet;

use alcamo\exception\{SyntaxError, Underflow};

class Col
{
    public const REGEXP = '/^[A-Z]{1,3}$/';

    private $value_;

    public static function newFromString(string $value): self
    {
        if (!preg_match(static::REGEXP, $value)) {
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

    /// return a new object with incremented value
    public function inc(): self
    {
        $value = $this->value_;
        return new self(++$value);
    }

    /// return a new object with decremented value
    public function dec(): self
    {
        return new self($this->decString($this->value_));
    }

    /// return a new object with $diff added
    public function add(int $diff): self
    {
        switch (true) {
            case $diff > 0:
                $result = clone ($this);

                for ($i = 0; $i < $diff; $i++) {
                    $result->value_++;
                }

                return $result;

            case $diff < 0:
                $value = $this->value_;

                for ($i = $diff; $i < 0; $i++) {
                    $value = $this->decString($value);
                }

                return new self($value);

            default:
                return clone $this;
        }
    }

    /// Decrement a string representing a column
    private function decString(string $value): string
    {
        for ($i = strlen($value) - 1; $i >= 0; $i--) {
            if ($value[$i] != 'A') {
                $value[$i] = chr(ord($value[$i]) - 1);
                return $value;
            } else {
                $value[$i] = 'Z';

                if ($i == 0) {
                    if ($value == 'Z') {
                        throw new Underflow();
                    }

                    return substr($value, 1);
                }
            }
        }
    }
}

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
        $value = $this->value_;

        for ($i = strlen($value) - 1; $i >= 0; $i--) {
            if ($value[$i] != 'A') {
                $value[$i] = chr(ord($value[$i]) - 1);
                return new self($value);
            } else {
                $value[$i] = 'Z';

                if ($i == 0) {
                    if ($value == 'Z') {
                        throw new Underflow();
                    }

                    return new self(substr($value, 1));
                }
            }
        }
    }

}

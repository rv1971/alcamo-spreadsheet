<?php

namespace alcamo\spreadsheet;

use alcamo\exception\{SyntaxError, Underflow};
use PHPUnit\Framework\TestCase;

class ColTest extends TestCase
{
    /**
     * @dataProvider incDecProvider
     */
    public function testIncDec($col, $incCol): void
    {
        $this->assertSame(
            (string)((Col::newFromString($col))->inc()),
            $incCol
        );

        $this->assertSame(
            (string)((Col::newFromString($incCol))->dec()),
            $col
        );
    }

    public function incDecProvider(): array
    {
        return [
            [ 'A', 'B' ],
            [ 'B', 'C' ],
            [ 'Z', 'AA' ],
            [ 'AA', 'AB' ],
            [ 'AB', 'AC' ],
            [ 'AZ', 'BA' ],
            [ 'BA', 'BB' ],
            [ 'BX', 'BY' ],
            [ 'BZ', 'CA' ],
            [ 'ZZ', 'AAA' ],
            [ 'AAA', 'AAB' ],
            [ 'AAB', 'AAC' ],
            [ 'AAZ', 'ABA' ],
            [ 'ABZ', 'ACA' ],
            [ 'AZZ', 'BAA' ],
            [ 'YZZ', 'ZAA' ]
        ];
    }

    /**
     * @dataProvider addProvider
     */
    public function testAdd($value, $diff, $expectedValue): void
    {
        $this->assertEquals(
            Col::newFromString($expectedValue),
            (Col::newFromString($value))->add($diff)
        );
    }

    public function addProvider(): array
    {
        return [
            [ 'AAA', 0, 'AAA' ],
            [ 'A', 1, 'B' ],
            [ 'B', 2, 'D' ],
            [ 'D', 3, 'G' ],
            [ 'AZW', 5, 'BAB' ],
            [ 'G', -1, 'F' ],
            [ 'AAA', -1, 'ZZ' ],
            [ 'X', -2, 'V' ],
            [ 'AGB', -2, 'AFZ' ],
            [ 'AB', -4, 'X' ],
            [ 'IAD', -6, 'HZX' ]
        ];
    }

    public function testSyntaxException(): void
    {
        $this->expectException(SyntaxError::class);
        $this->expectExceptionMessage('Syntax error in "Aa"');

        Col::newFromString('Aa');
    }

    public function testUnderflowException(): void
    {
        $this->expectException(Underflow::class);
        $this->expectExceptionMessage(
            'Underflow in object <alcamo\spreadsheet\Col>"A"'
        );

        Col::newFromString('A')->dec();
    }
}

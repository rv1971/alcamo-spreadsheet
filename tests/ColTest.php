<?php

namespace alcamo\spreadsheet;

use alcamo\exception\{SyntaxError, Underflow};
use PHPUnit\Framework\TestCase;

class ColTest extends TestCase
{
    /**
     * @dataProvider incDecProvider
     */
    public function testIncDec($col, $incCol)
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

    public function incDecProvider()
    {
        return [
            [ 'A', 'B' ],
            [ 'B', 'C' ],
            [ 'Z', 'AA' ],
            [ 'AA', 'AB' ],
            [ 'AB', 'AC' ],
            [ 'AZ', 'BA' ],
            [ 'BA', 'BB' ],
            [ 'BB', 'BC' ],
            [ 'BZ', 'CA' ],
            [ 'ZZ', 'AAA' ],
            [ 'AAA', 'AAB' ],
            [ 'AAB', 'AAC' ],
            [ 'AAZ', 'ABA' ],
            [ 'ABZ', 'ACA' ],
            [ 'AZZ', 'BAA' ]
        ];
    }

    public function testSyntaxException()
    {
        $this->expectException(SyntaxError::class);
        $this->expectExceptionMessage('Syntax error in "Aa"');

        Col::newFromString('Aa');
    }

    public function testUnderflowException()
    {
        $this->expectException(Underflow::class);
        $this->expectExceptionMessage(
            'Underflow in object <alcamo\spreadsheet\Col>"A"'
        );

        Col::newFromString('A')->dec();
    }
}

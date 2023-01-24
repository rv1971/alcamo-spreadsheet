<?php

namespace alcamo\spreadsheet;

use alcamo\rdfa\MediaType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PHPUnit\Framework\TestCase;

class XlsxResponseTest extends TestCase
{
    public function testEmit()
    {
        $outFilename = __DIR__ . DIRECTORY_SEPARATOR . 'XlsxResponseAux.xlsx';

        $wsTitle = 'lorem-ipsum';

        $xlsx = shell_exec(
            'PHPUNIT_COMPOSER_INSTALL="' . PHPUNIT_COMPOSER_INSTALL . '" php '
            . __DIR__ . DIRECTORY_SEPARATOR
            . "XlsxResponseAux.php $wsTitle > $outFilename"
        );

        $this->assertSame(
            XlsxResponse::MEDIA_TYPE,
            (string)MediaType::newFromFilename($outFilename)
        );

        $this->assertSame(
            $wsTitle,
            IOFactory::load($outFilename)->getActiveSheet()->getTitle()
        );

        unlink($outFilename);
    }
}

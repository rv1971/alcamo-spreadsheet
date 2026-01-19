<?php

namespace alcamo\spreadsheet;

use alcamo\process\Process;
use alcamo\rdfa\MediaType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PHPUnit\Framework\TestCase;

class XlsxResponseTest extends TestCase
{
    public function testEmit()
    {
        $outFilename = __DIR__ . DIRECTORY_SEPARATOR . 'XlsxResponseAux.xlsx';

        $wsTitle = 'lorem-ipsum';

        $process = new Process(
            "php XlsxResponseAux.php $wsTitle",
            __DIR__,
            [ 'PHPUNIT_COMPOSER_INSTALL' => PHPUNIT_COMPOSER_INSTALL ],
            [ 1 => [ 'file', $outFilename, 'w' ] ]
        );

        $outputLine = fgets($process->getStderr());

        $output = explode(' ', $outputLine);

        $process->close();

        $this->assertSame(XlsxResponse::MEDIA_TYPE, $output[0]);

        $this->assertSame('lorem-ipsum_3.14.xlsx', $output[1]);

        $this->assertSame(filesize($outFilename), (int)$output[2]);

        $this->assertSame(
            XlsxResponse::MEDIA_TYPE,
            (string)MediaType::newFromFilename($outFilename)
        );

        $spreadsheet = IOFactory::load($outFilename);

        $this->assertSame(
            $wsTitle,
            $spreadsheet->getActiveSheet()->getTitle()
        );

        $this->assertSame(
            'lorem-ipsum',
            $spreadsheet->getProperties()->getCustomPropertyValue('Identifier')
        );

        $this->assertSame(
            '3.14',
            $spreadsheet->getProperties()->getCustomPropertyValue('Version')
        );

        unlink($outFilename);
    }
}

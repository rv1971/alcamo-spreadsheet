<?php

namespace alcamo\spreadsheet;

use alcamo\rdfa\{DateTimeLiteral, RdfaData, StringLiteral};
use PHPUnit\Framework\TestCase;

class MySpreadsheet extends Spreadsheet
{
    public const DEFAULT_STYLE = [
        'alignment' => [ 'vertical' => 'top' ]
    ];
}

class SpreadsheetTest extends TestCase
{
    public function testBasics(): void
    {
        $rdfaData = RdfaData::newFromIterable(
            [
                [ 'dc:created',      '2021-12-27' ],
                [ 'dc:creator',      'rv1971' ],
                [ 'dc:modified',     '2021-12-28' ],
                [ 'dc:publisher',    'Example.org ltd.' ],
                [ 'dc:title',        'Foo vs. Bar' ],
                [ 'dc:audience',     'World' ],
                [ 'dc:identifier',   '123e4567-e89b-12d3-a456-426614174000' ],
                [ 'dc:language',     'eu' ],
                [ 'owl:versionInfo', '0.42.1' ]
            ]
        );

        $spreadsheet = new MySpreadsheet($rdfaData);

        $created = new DateTimeLiteral(
            (new \DateTime())
                ->setTimestamp($spreadsheet->getProperties()->getCreated())
        );

        $this->assertEquals(
            $rdfaData['dc:created']->first()->getObject(),
            $created
        );

        $this->assertEquals(
            $rdfaData['dc:creator']->first()->getObject(),
            new StringLiteral($spreadsheet->getProperties()->getCreator())
        );

        $modified = new DateTimeLiteral(
            (new \DateTime())
                ->setTimestamp($spreadsheet->getProperties()->getModified())
        );

        $this->assertEquals(
            $rdfaData['dc:modified']->first()->getObject(),
            $modified
        );

        $this->assertEquals(
            $rdfaData['dc:publisher']->first()->getObject(),
            new StringLiteral($spreadsheet->getProperties()->getCompany())
        );

        $this->assertEquals(
            (string)$rdfaData['dc:title']->first(),
            new StringLiteral($spreadsheet->getProperties()->getTitle())
        );

        $this->assertEquals(
            $rdfaData['dc:audience']->first()->getObject(),
            new StringLiteral(
                $spreadsheet->getProperties()->getCustomPropertyValue('Audience')
            )
        );

        $this->assertEquals(
            $rdfaData['dc:identifier']->first()->getObject(),
            new StringLiteral(
                $spreadsheet->getProperties()->getCustomPropertyValue('Identifier')
            )
        );

        $this->assertEquals(
            (string)$rdfaData['dc:language']->first(),
            $spreadsheet->getProperties()->getCustomPropertyValue('Language')
        );

        $this->assertEquals(
            $rdfaData['owl:versionInfo']->first()->getObject(),
            new StringLiteral(
                $spreadsheet->getProperties()->getCustomPropertyValue('Version')
            )
        );

        $this->assertSame(0, count($spreadsheet->getAllSheets()));

        $this->assertSame(
            MySpreadsheet::DEFAULT_STYLE['alignment']['vertical'],
            $spreadsheet->getDefaultStyle()
                ->getAlignment()
                ->getVertical()
        );
    }
}

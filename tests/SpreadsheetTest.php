<?php

namespace alcamo\spreadsheet;

use alcamo\rdfa\RdfaData;
use PHPUnit\Framework\TestCase;

class MySpreadsheet extends Spreadsheet
{
    public const DEFAULT_STYLE = [
        'alignment' => [ 'vertical' => 'top' ]
    ];
}

class SpreadsheetTest extends TestCase
{
    public function testBasics()
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

        $created = new \DateTime();
        $created->setTimestamp($spreadsheet->getProperties()->getCreated());

        $this->assertEquals(
            $rdfaData['dc:created']->getObject(),
            $created
        );

        $this->assertSame(
            array_values($rdfaData['dc:creator'])[0]->getObject(),
            $spreadsheet->getProperties()->getCreator()
        );

        $modified = new \DateTime();
        $modified->setTimestamp($spreadsheet->getProperties()->getModified());

        $this->assertEquals(
            $rdfaData['dc:modified']->getObject(),
            $modified
        );

        $this->assertSame(
            array_values($rdfaData['dc:publisher'])[0]->getObject(),
            $spreadsheet->getProperties()->getCompany()
        );

        $this->assertSame(
            array_values($rdfaData['dc:title'])[0]->getObject(),
            $spreadsheet->getProperties()->getTitle()
        );

        $this->assertSame(
            array_values($rdfaData['dc:audience'])[0]->getObject(),
            $spreadsheet->getProperties()->getCustomPropertyValue('Audience')
        );

        $this->assertSame(
            array_values($rdfaData['dc:identifier'])[0]->getObject(),
            $spreadsheet->getProperties()->getCustomPropertyValue('Identifier')
        );

        $this->assertEquals(
            (string)array_values($rdfaData['dc:language'])[0],
            $spreadsheet->getProperties()->getCustomPropertyValue('Language')
        );

        $this->assertSame(
            $rdfaData['owl:versionInfo']->getObject(),
            $spreadsheet->getProperties()->getCustomPropertyValue('Version')
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

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
                'dc:created'      => '2021-12-27',
                'dc:creator'      => 'rv1971',
                'dc:modified'     => '2021-12-28',
                'dc:publisher'    => 'Example.org ltd.',
                'dc:title'        => 'Foo vs. Bar',
                'dc:audience'     => 'World',
                'owl:versionInfo' => '0.42.1'
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
            $rdfaData['dc:creator']->getObject(),
            $spreadsheet->getProperties()->getCreator()
        );

        $modified = new \DateTime();
        $modified->setTimestamp($spreadsheet->getProperties()->getModified());

        $this->assertEquals(
            $rdfaData['dc:modified']->getObject(),
            $modified
        );

        $this->assertSame(
            $rdfaData['dc:publisher']->getObject(),
            $spreadsheet->getProperties()->getCompany()
        );

        $this->assertSame(
            $rdfaData['dc:title']->getObject(),
            $spreadsheet->getProperties()->getTitle()
        );

        $this->assertSame(
            $rdfaData['dc:audience']->getObject(),
            $spreadsheet->getProperties()->getCustomPropertyValue('Audience')
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

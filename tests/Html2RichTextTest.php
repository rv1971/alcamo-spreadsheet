<?php

namespace alcamo\spreadsheet;

use PhpOffice\PhpSpreadsheet\RichText\{RichText as RT, Run as R};
use PhpOffice\PhpSpreadsheet\Style\{Color, Font};
use PHPUnit\Framework\TestCase;

class Html2RichTextTest extends TestCase
{
    /**
     * @dataProvider createProvider
     */
    public function testCreate($htmlText, $expectedRichText)
    {
        $this->assertEquals(
            $expectedRichText,
            (new Html2RichText())->create($htmlText)
        );
    }

    public function createProvider()
    {
        return [
            [
                'Lorem ipsum.',
                (new RT())->addText(new R('Lorem ipsum.'))
            ],
            [
                '<b color="#ff0000" face="Times" size="18">Lorem <font>ip</font>sum.</b>',
                (new RT())
                    ->addText(
                        (new R('Lorem ipsum.'))
                            ->setFont(
                                (new Font())
                                    ->setBold(true)
                                    ->setColor(new Color('ffff0000'))
                                    ->setName('Times')
                                    ->setSize(18)
                            )
                    )
            ],
            [
                ' <B>Lorem <i>ipsum</i> dolor</B>  ',
                (new RT())
                    ->addText(
                        (new R('Lorem '))
                            ->setFont((new Font())->setBold(true))
                    )
                    ->addText(
                        (new R('ipsum'))
                            ->setFont(
                                (new Font())->setBold(true)->setItalic(true)
                            )
                    )
                    ->addText(
                        (new R(' dolor'))
                            ->setFont((new Font())->setBold(true))
                    )
            ],
            [
                "\n <B><u>L</u>orem\t<i>ip<sUb>sum</sUb><suP>.</suP></i></B> dolor \n ",
                (new RT())
                    ->addText(
                        (new R('L'))
                            ->setFont(
                                (new Font())->setBold(true)->setUnderline(true)
                            )
                    )
                    ->addText(
                        (new R('orem '))
                            ->setFont((new Font())->setBold(true))
                    )
                    ->addText(
                        (new R('ip'))
                            ->setFont(
                                (new Font())->setBold(true)->setItalic(true)
                            )
                    )
                    ->addText(
                        (new R('sum'))
                            ->setFont(
                                (new Font())
                                    ->setBold(true)
                                    ->setItalic(true)
                                    ->setSubscript(true)
                            )
                    )
                    ->addText(
                        (new R('.'))
                            ->setFont(
                                (new Font())
                                    ->setBold(true)
                                    ->setItalic(true)
                                    ->setSuperscript(true)
                            )
                    )
                    ->addText(new R(' dolor'))
            ],
            [
                '<p><code>Lorem </code><del>ipsum <br/></del><em>dolor</em></p>'
                . '<ins>sit </ins><strong>amet,</strong>',
                (new RT())
                    ->addText(
                        (new R('Lorem '))
                            ->setFont((new Font())->setName('Courier New'))
                    )
                    ->addText(
                        (new R("ipsum \n"))
                            ->setFont((new Font())->setStrikethrough(true))
                    )
                    ->addText(
                        (new R("dolor"))
                            ->setFont((new Font())->setItalic(true))
                    )
                    ->addText(new R("\n\n"))
                    ->addText(
                        (new R('sit '))
                            ->setFont((new Font())->setUnderline(true))
                    )
                    ->addText(
                        (new R('amet,'))
                            ->setFont((new Font())->setBold(true))
                    )
            ],
            [
                '<h1><span color="#00ff00">H</span>eading 1</h1>'
                . '<h2>Heading 2</h2><h3>Heading 3</h3><h4>Heading 4</h4>',
                (new RT())
                    ->addText(
                        (new R("H"))
                            ->setFont(
                                (new Font())
                                    ->setBold(true)
                                    ->setSize(16)
                                    ->setColor(new Color('ff00ff00'))
                            )
                    )
                    ->addText(
                        (new R("eading 1\n\n"))
                            ->setFont(
                                (new Font())
                                    ->setBold(true)
                                    ->setSize(16)
                            )
                    )
                    ->addText(
                        (new R("Heading 2\n\n"))
                            ->setFont(
                                (new Font())
                                    ->setBold(true)
                                    ->setSize(14)
                            )
                    )
                    ->addText(
                        (new R("Heading 3\n\n"))
                            ->setFont(
                                (new Font())
                                    ->setBold(true)
                                    ->setSize(12)
                            )
                    )
                    ->addText(
                        (new R("Heading 4"))
                            ->setFont((new Font())->setBold(true))
                    )
            ],
            [
                'Lorem<ul><li>ip</li><li><u>sum</u></li></ul>dolor',
                (new RT())
                    ->addText(new R("Lorem\n • ip\n • "))
                    ->addText(
                        (new R("sum"))
                            ->setFont((new Font())->setUnderline(true))
                    )
                    ->addText(new R("\ndolor"))
            ],
            [
                '<p>Heading</p>'
                . '<ul>'
                . '<li>1</li>'
                . '<li>2<ul><li>2.1</li><li>2.2</li></ul></li>'
                . '<li>2</li>'
                . '</ul>',
                (new RT())
                    ->addText(new R("Heading\n\n • 1\n • 2\n    • 2.1\n    • 2.2\n • 2"))
            ],


        ];
    }
}

<?php

namespace alcamo\spreadsheet;

use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Style\{Color, Font};

/**
 * @brief Transform (pseudo-)HTML to rich text.
 *
 * Accepts pseudo-HTML in the sense that the attributes `color`, `face` and
 * `size` are accepted for any element, which was never the case in any HTML
 * version.
 *
 * @sa Basic idea inspired by [PhpSpreadsheet\Helper\Html](https://github.com/PHPOffice/PhpSpreadsheet/blob/master/src/PhpSpreadsheet/Helper/Html.php)
 *
 * @date Last reviewed 2026-01-15
 */
class Html2RichText
{
    public const MONOSPACE_FONT = 'Courier New';
    public const H1_SIZE = 16;
    public const H2_SIZE = 14;
    public const H3_SIZE = 12;

    /**
     * @brief Map some tag names to start hooks
     *
     * A start hook consists of the name of a method and a boolean telling
     * whether the method might modfy the font.
     */
    public const START_HOOKS = [
        'b'      => [ 'setBold', true ],
        'code'   => [ 'setMonospace', true ],
        'del'    => [ 'setStrikethrough', true ],
        'em'     => [ 'setItalic', true ],
        'h1'     => [ 'startH1', true ],
        'h2'     => [ 'startH2', true ],
        'h3'     => [ 'startH3', true ],
        'h4'     => [ 'setBold', true ],
        'i'      => [ 'setitalic', true ],
        'ins'    => [ 'setUnderline', true ],
        'li'     => [ 'appendBullet', false ],
        'strong' => [ 'setBold', true ],
        'sub'    => [ 'setSubscript', true ],
        'sup'    => [ 'setSuperscript', true ],
        'u'      => [ 'setUnderline', true ],
        'ul'     => [ 'startUl', false ]
    ];

    /// Map some tag names to end hooks
    public const END_HOOKS = [
        'br' => 'appendLf',
        'h1' => 'appendLfLf',
        'h2' => 'appendLfLf',
        'h3' => 'appendLfLf',
        'h4' => 'appendLfLf',
        'li' => 'endLi',
        'p'  => 'appendLfLf',
        'ul' => 'endUl'
    ];

    private $richText_;   ///< RichText
    private $font_;       ///< current Font
    private $fontStack_;  ///< stack of Font objects in outer contexts
    private $listLevel_;  ///< nesting level of \<ul>
    private $stringData_; ///< string data collected so far

    /**
     * @brief Create rich text from HTML
     *
     * @param DOMNode|string $html HTML data
     */
    public function create($html): RichText
    {
        $this->init();

        return $this->parse($html);
    }

    /**
     * @brief Append rich text created from HTML to existing rich text
     *
     * @param DOMNode|string $html HTML data
     */
    public function append(RichText $richText, $html): RichText
    {
        $this->init($richText);

        return $this->parse($html);
    }

    private function init(?RichText $richText = null)
    {
        $this->richText_ = $richText ?? new RichText();
        $this->font_ = new Font();
        $this->fontStack_ = [];
        $this->listLevel_ = 0;
        $this->stringData_ = '';
    }

    private function parse($html): RichText
    {
        if ($html instanceof \DOMNode) {
            $dom = $html;
        } else {
            $dom = new \DOMDocument();

            $dom->preserveWhiteSpace = false;

            /* Load as XML rather than HTML because in HTML mode, PHP applies a
             * lot of unwanted modifications to the content. */
            $dom->loadXML("<body>$html</body>");
        }

        $this->parseChildren($dom);

        $this->addRun();

        /* Remove trailing whitespace from last rich text element. */

        $richTextElements = $this->richText_->getRichTextElements();

        $lastElement = end($richTextElements);

        $lastElement->setText(rtrim($lastElement->getText()));

        return $this->richText_;
    }

    private function parseChildren(\DOMNode $parent): void
    {
        foreach ($parent->childNodes as $node) {
            if ($node instanceof \DOMText) {
                $this->stringData_ .=
                    preg_replace('/\s+/', ' ', $node->nodeValue);
            } elseif ($node instanceof \DOMElement) {
                $tagName = strtolower($node->tagName);

                $startHook = static::START_HOOKS[$tagName] ?? null;

                $modifiesFont = isset($startHook) && $startHook[1]
                || $node->attributes->length;

                if ($modifiesFont) {
                    $this->addRun();

                    $this->fontStack_[] = $this->font_;
                    $this->font_ = clone $this->font_;
                }

                if (isset($startHook)) {
                    $startHookMethod = $startHook[0];
                    $this->$startHookMethod($node);
                }

                if (
                    $node->hasAttribute('color')
                    && $node->getAttribute('color')[0] == '#'
                ) {
                    $this->font_->setColor(
                        new Color(
                            'ff' . substr($node->getAttribute('color'), 1)
                        )
                    );
                }

                if ($node->hasAttribute('face')) {
                    $this->font_->setName($node->getAttribute('face'));
                }

                if ($node->hasAttribute('size')) {
                    $this->font_->setSize($node->getAttribute('size'));
                }

                $this->parseChildren($node);

                $endHook = static::END_HOOKS[$tagName] ?? null;

                if (isset($endHook)) {
                    $this->$endHook($node);
                }

                if ($modifiesFont) {
                    $this->addRun();

                    $this->font_ = array_pop($this->fontStack_);
                }
            }
        }
    }

    private function addRun(): void
    {
        if ($this->stringData_ != '') {
            $this->richText_
                ->createTextRun($this->stringData_)
                ->setFont($this->font_);

            $this->stringData_ = '';
        }
    }

    private function setBold(): void
    {
        $this->font_->setBold(true);
    }

    private function setMonospace(): void
    {
        $this->font_->setName(static::MONOSPACE_FONT);
    }

    private function setStrikethrough(): void
    {
        $this->font_->setStrikethrough(true);
    }

    private function startH1(): void
    {
        $this->font_->setBold(true)->setSize(static::H1_SIZE);
    }

    private function startH2(): void
    {
        $this->font_->setBold(true)->setSize(static::H2_SIZE);
    }

    private function startH3(): void
    {
        $this->font_->setBold(true)->setSize(static::H3_SIZE);
    }

    private function startUl(\DOMElement $node): void
    {
        if (
            !($node->previousSibling instanceof \DOMElement
              && in_array(
                  strtolower($node->previousSibling->nodeName),
                  [ 'br', 'h1', 'h2', 'h3', 'p' ]
              )
            )
        ) {
            $this->stringData_ .= "\n";
        }

        $this->listLevel_++;
    }

    private function setSubscript(): void
    {
        $this->font_->setSubscript(true);
    }

    private function setSuperscript(): void
    {
        $this->font_->setSuperscript(true);
    }

    private function setUnderline(): void
    {
        $this->font_->setUnderline(true);
    }

    private function appendBullet(\DOMElement $node): void
    {
        $this->stringData_ .= str_pad('', 3 * $this->listLevel_ - 3) . ' • ';
    }

    private function setItalic(): void
    {
        $this->font_->setItalic(true);
    }

    private function appendLf(): void
    {
        $this->stringData_ .= "\n";
    }

    private function appendLfLf(): void
    {
        $this->stringData_ .= "\n\n";
    }

    private function endLi(\DOMElement $node): void
    {
        if (
            !($node->lastChild instanceof \DOMElement
              && in_array(
                  strtolower($node->lastChild->nodeName),
                  [ 'br', 'h1', 'h2', 'h3', 'p', 'ul' ]
              )
            )
        ) {
            $this->stringData_ .= "\n";
        }
    }

    private function endUl(): void
    {
        $this->listLevel_--;
    }
}

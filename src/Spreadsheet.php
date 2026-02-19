<?php

/**
 * @namespace alcamo::spreadsheet
 *
 * @brief Convenience package built on top of PhpSpreadsheet
 */

namespace alcamo\spreadsheet;

use alcamo\rdfa\{HavingRdfaDataInterface, RdfaData, StringLiteral};
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Spreadsheet as PhpOfficeSpreadsheet;

/**
 * @brief Spreadsheet with RDFa data
 *
 * @date Last reviewed 2026-01-14
 */
class Spreadsheet extends PhpOfficeSpreadsheet implements HavingRdfaDataInterface
{
    public const WORKSHEET_CLASS = Worksheet::class;

    /// Map of RDFa properties to methods of the Properties class
    public const PROP_TO_METHODS = [
        'dc:created'   => 'setCreated',
        'dc:creator'   => [ 'setCreator', 'setLastModifiedBy' ],
        'dc:modified'  => 'setModified',
        'dc:publisher' => 'setCompany',
        'dc:title'     => 'setTitle'
    ];

    /// Map of RDFa properties to spreadsheet custom properties.
    public const PROP_TO_CUSTOM_PROP = [
        'dc:audience'     => 'Audience',
        'dc:identifier'   => 'Identifier',
        'dc:language'     => 'Language',
        'owl:versionInfo' => 'Version'
    ];

    /// Default style used as argument for applyFromArray()
    public const DEFAULT_STYLE = [];

    private $rdfaData_;

    /**
     * @brief Create new spreadsheet and set its metadata
     *
     * Furthermore:
     * - Remove the initial worksheet.
     * - Set the default style to @ref DEFAULT_STYLE.
     */
    public function __construct($rdfaData = null)
    {
        parent::__construct();

        switch (true) {
            case $rdfaData instanceof RdfaData:
                $this->rdfaData_ = $rdfaData;
                break;

            case isset($rdfaData):
                $this->rdfaData_ = RdfaData::newFromIterable($rdfaData);
                break;

            default:
                $this->rdfaData_ = RdfaData::newEmpty();
        }

        foreach (static::PROP_TO_METHODS as $prop => $methods) {
            if (isset($this->rdfaData_[$prop])) {
                foreach ((array)$methods as $method) {
                    $this->getProperties()->$method(
                        (string)$this->rdfaData_->getFirstStmt($prop)
                    );
                }
            }
        }

        foreach (static::PROP_TO_CUSTOM_PROP as $prop => $customProp) {
            if (isset($this->rdfaData_[$prop])) {
                $dataType =
                    $this->rdfaData_->getFirstObject($prop)
                    instanceof StringLiteral
                    ? DataType::TYPE_STRING
                    : null;

                $this->getProperties()->setCustomProperty(
                    $customProp,
                    (string)$this->rdfaData_->getFirstStmt($prop),
                    $dataType
                );
            }
        }

        $this->removeSheetByIndex(0);

        $this->getDefaultStyle()->applyFromArray(static::DEFAULT_STYLE);
    }

    public function getRdfaData(): RdfaData
    {
        return $this->rdfaData_;
    }

    public function createSheet($sheetIndex = null)
    {
        $class = static::WORKSHEET_CLASS;

        $newSheet = new $class($this);

        $this->addSheet($newSheet, $sheetIndex);

        return $newSheet;
    }
}

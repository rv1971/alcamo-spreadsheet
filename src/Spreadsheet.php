<?php

/**
 * @namespace alcamo::spreadsheet
 *
 * @brief Convenience package built on top of PhpSpreadsheet
 */

namespace alcamo\spreadsheet;

use alcamo\rdfa\{HasRdfaDataTrait, RdfaData};
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Spreadsheet as PhpOfficeSpreadsheet;

/// Spreadsheet with RDFa data
class Spreadsheet extends PhpOfficeSpreadsheet
{
    use HasRdfaDataTrait;

    /// Map of RDFa properties to methods of the Properties class
    public const PROP_2_METHODS = [
        'dc:created'   => 'setCreated',
        'dc:creator'   => [ 'setCreator', 'setLastModifiedBy' ],
        'dc:modified'  => 'setModified',
        'dc:publisher' => 'setCompany',
        'dc:title'     => 'setTitle'
    ];

    /// Map of RDFa properties to spreadsheet custom properties.
    public const PROP_2_CUSTOM_PROP = [
        'dc:audience'     => 'Audience',
        'dc:identifier'   => 'Identifier',
        'dc:language'     => 'Language',
        'owl:versionInfo' => 'Version'
    ];

    /// Default style used as argument for applyFromArray()
    public const DEFAULT_STYLE = [];

    /**
     * @brief Create new spreadsheet and set its metadata
     *
     * Furthermore:
     * - Remove the initial worksheet.
     * - Set the default style to @ref DEFAULT_STYLE.
     */
    public function __construct(?RdfaData $rdfaData = null)
    {
        parent::__construct();

        $this->rdfaData_ = $rdfaData ?? new RdfaData();

        foreach (static::PROP_2_METHODS as $prop => $methods) {
            if (isset($this->rdfaData_[$prop])) {
                foreach ((array)$methods as $method) {
                    $this->getProperties()->$method(
                        (string)$this->rdfaData_[$prop]
                    );
                }
            }
        }

        foreach (static::PROP_2_CUSTOM_PROP as $prop => $customProp) {
            if (isset($this->rdfaData_[$prop])) {
                $dataType =
                    is_string($this->rdfaData_[$prop]->getObject())
                    ? DataType::TYPE_STRING
                    : null;

                $this->getProperties()->setCustomProperty(
                    $customProp,
                    (string)$this->rdfaData_[$prop]->getObject(),
                    $dataType
                );
            }
        }

        $this->removeSheetByIndex(0);

        $this->getDefaultStyle()->applyFromArray(static::DEFAULT_STYLE);
    }
}

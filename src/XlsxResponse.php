<?php

namespace alcamo\spreadsheet;

use alcamo\http\{ResourceStream, Response};
use alcamo\rdfa\RdfaData;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * @brief HTTP response containing an xlsx document
 *
 * @date Last reviewed 2026-01-14
 */
class XlsxResponse extends Response
{
    public const MEDIA_TYPE =
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    /**
     * @brief Construct a complete response from a spreadsheet.
     *
     * Takes over any relevant RDFa data from $spreadsheet and adds data
     * needed for HTTP headers such as `Content-Type`, `Content-Length` and
     * `Content-Disposition`. For `Content-Disposition`, a file name is
     * constructed from `dc:identifier` (which is expected to be present) and
     * `owl:versionInfo` (if present).
     */
    public function __construct(Spreadsheet $spreadsheet)
    {
        $resource = fopen('php://memory', 'wb+');

        (new Xlsx($spreadsheet))->save($resource);

        /* Needed to get content-length. */
        fseek($resource, 0, SEEK_END);

        $rdfaData = $spreadsheet->getRdfaData();

        $rdfaData = $rdfaData->add(
            RdfaData::newFromIterable(
                [
                    [ 'dc:format', static::MEDIA_TYPE ],
                    [ 'http:content-length', ftell($resource) ],
                    [
                        'http:content-disposition',
                        $rdfaData->getFirstStmt('dc:identifier')
                            . (isset($rdfaData['owl:versionInfo'])
                               ? '_' . $rdfaData->getFirstStmt('owl:versionInfo')
                               : '')
                            . '.xlsx'
                    ]
                ]
            )
        );

        parent::__construct($rdfaData, new ResourceStream($resource));
    }
}

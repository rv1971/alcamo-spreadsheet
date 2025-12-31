<?php

namespace alcamo\spreadsheet;

use alcamo\http\{ResourceStream, Response};
use alcamo\rdfa\RdfaData;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/// HTTP response containing an xlsx document
class XlsxResponse extends Response
{
    public const MEDIA_TYPE =
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    public function __construct(Spreadsheet $spreadsheet)
    {
        $resource = fopen('php://memory', 'wb+');

        (new Xlsx($spreadsheet))->save($resource);

        fseek($resource, 0, SEEK_END);

        $rdfaData = $spreadsheet->getRdfaData();

        $rdfaData = $rdfaData->add(
            RdfaData::newFromIterable(
                [
                    [ 'dc:format', static::MEDIA_TYPE ],
                    [ 'http:content-length', ftell($resource) ],
                    [
                        'http:content-disposition',
                        array_values($rdfaData['dc:identifier'])[0]
                            . (isset($rdfaData['owl:versionInfo'])
                               ? '_' . $rdfaData['owl:versionInfo']
                               : '')
                            . '.xlsx'
                    ]
                ]
            )
        );

        parent::__construct($rdfaData, new ResourceStream($resource));
    }
}

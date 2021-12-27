<?php

namespace alcamo\spreadsheet;

use alcamo\http\Response;
use alcamo\rdfa\RdfaData;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/// HTTP response containing an xlsx document
class XlsxResponse extends Response
{
    public const MEDIA_TYPE =
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    public function __construct(Spreadsheet $spreadsheet)
    {
        $rdfaData = $spreadsheet->getRdfaData();

        $rdfaData = $rdfaData->add(
            RdfaData::newFromIterable(
                [
                    'dc:format' => static::MEDIA_TYPE,
                    'header:content-disposition'
                    => $rdfaData['dc:identifier']
                    . (isset($rdfaData['owl:versionInfo'])
                       ? '_' . $rdfaData['owl:versionInfo']
                       : '')
                    . '.xlsx'
                ]
            )
        );

        $resource = fopen('php://memory', 'wb+');

        parent::__construct($rdfaData, $resource);

        (new Xlsx($spreadsheet))->save($resource);
    }
}

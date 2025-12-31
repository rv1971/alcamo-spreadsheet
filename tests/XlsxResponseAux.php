<?php

namespace alcamo\spreadsheet;

use alcamo\rdfa\RdfaData;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

require getenv('PHPUNIT_COMPOSER_INSTALL');

[ , $wsTitle ] = $argv;

$spreadsheet = new Spreadsheet(
    RdfaData::newFromIterable(
        [
            [ 'dc:identifier', 'lorem-ipsum' ],
            [ 'owl:versionInfo', '3.14' ]
        ]
    )
);

$spreadsheet->addSheet(new Worksheet($spreadsheet, $wsTitle));

(new XlsxResponse($spreadsheet))->emit();

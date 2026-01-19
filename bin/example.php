<?php

use alcamo\spreadsheet\{Spreadsheet, XlsxResponse};

include $_composer_autoload_path ?? __DIR__ . '/../vendor/autoload.php';

$spreadsheet = new Spreadsheet(
    [
        [ 'dc:title',        'Example spreadsheet' ],
        [ 'dc:identifier',   'my-example' ],
        [ 'owl:versionInfo', '2.71' ]
    ]
);

$worksheet = $spreadsheet->createSheet();

$worksheet->setCoordinate('B2')->writeRow(
    [
        [ 'Hello, ', 'style' => [ 'font' => [ 'bold' => true ] ] ],
        'world!'
    ],
    [ 'font' => [ 'color' => [ 'rgb' => '008000' ] ] ]
);

$worksheet->writeRow(
    [ 'foo', 'bar' ],
    [ 'font' => [ 'color' => [ 'rgb' => '808080' ] ] ]
);

(new XlsxResponse($spreadsheet))->emit();

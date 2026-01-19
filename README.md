# Usage example

~~~
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
~~~

This example is contained in this package as a file in the `bin`
directory. If called through a webserver, it returns a spreadsheet
(with `Content-Length` and `Content-Type` headers) and suggests to
save it as `my-example_2.71.xlsx`.

The spreadsheet contains data in the ranges B2:C2 and B3:C3 since in
this example `writeRow()` automatically moves the current cell to the
start of the next row.

# Classes

* `Col` - arithmetics with alphabetic spreadsheet column counters.
* `Html2RichText` - a converter that creates Rich Text from simple HTML
  with formatting such as font color.
* `Spreadsheet` - an extension of the PhpOffice Spreadsheet class.
* `Worksheet` - an extension of the PhpOffice Worksheet class.
* `XlsxResponse` - an HTTP response containing a worksheet.

See the doxygen documentation for details.

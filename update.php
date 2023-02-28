<?php

require __DIR__ . '/vendor/autoload.php';

// The NAWS import spreadsheet from NAWS no longer imports directly into BMLT due to changes in the NAWS database.
// Fix up some problems so that it can be imported.

function update_file($inputFile, $outputFile)
{
    // $formatWorldCodeToName is copied from world_format_codes in src/legacy/local_server/server_admin/lang/en/server_admin_strings.inc.php
    // Just have code to add these to the array $formatNameToWorldId rather than reversing all the key/value pairs.
    $formatWorldIdToName = array(
        'OPEN' => 'Open',
        'CLOSED' => 'Closed',
        'WCHR' => 'Wheelchair-Accessible',
        'BEG' => 'Beginner/Newcomer',
        'BT' => 'Basic Text',
        'CAN' => 'Candlelight',
        'CPT' => '12 Concepts',
        'CW' => 'Children Welcome',
        'DISC' => 'Discussion/Participation',
        'GL' => 'Gay/Lesbian',
        'IP' => 'IP Study',
        'IW' => 'It Works Study',
        'JFT' => 'Just For Today Study',
        'LC' => 'Living Clean',
        'LIT' => 'Literature Study',
        'M' => 'Men',
        'MED' => 'Meditation',
        'NS' => 'Non-Smoking',
        'QA' => 'Questions & Answers',
        'RA' => 'Restricted Access',
        'S-D' => 'Speaker/Discussion',
        'SMOK' => 'Smoking',
        'SPK' => 'Speaker',
        'STEP' => 'Step',
        'SWG' => 'Step Working Guide Study',
        'TOP' => 'Topic',
        'TRAD' => 'Tradition',
        'VAR' => 'Format Varies',
        'W' => 'Women',
        'Y' => 'Young People',
        'LANG' => 'Alternate Language',
        'GP' => 'Guiding Principles',
        'NC' => 'No Children',
        'CH' => 'Closed Holidays',
        'VM' => 'Virtual',
        'HYBR' => 'Virtual and In-Person',
        'TC' => 'Temporarily Closed Facility',
        'SPAD' => 'Spiritual Principle a Day',
    );
    // There are some additional names not in the $formatWorldIdToName table -- add those explicitly
    $formatNameToWorldId = array(
        'Venue Temporarily Closed' => 'TC',
        'Basic Text Study' => 'BT',
        'Virtual Meeting' => 'VM',
        'Living Clean Study' => 'LC',
        'Step Study' => 'STEP',
        'Hybrid - Meets Online & In-person' => 'HYBR',
        'LGBTQ+' => 'GL',
        'Guiding Principles Study' => 'GP',
        'No Smoking' => 'NS',
        'Closed On Holidays' => 'CH',
        'Spiritual Principle a Day Study' => 'SPAD'
    );
    // $expectedColumns are ones that we fix up
    $expectedColumns = array('directions', 'additionaldirections', 'closed', 'format1', 'format2', 'format3', 'format4', 'format5');
    $directionsIndex = null;
    $additionalDirectionsIndex = null;
    $closedIndex = null;
    $roomIndex = null;
    $formatIndices = array();
    $spreadsheet = null;
    $columnNames = array();

    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($inputFile);
    $spreadsheet = $reader->load($inputFile);
    // get the column names from the first row in the spreadsheet
    $worksheet = $spreadsheet->getActiveSheet();
    // Get column names from first row.  Make them all lowercase for case insensitive string matching.
    $highestColumn = $worksheet->getHighestColumn();
    $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
    for ($i = 1; $i < $highestColumnIndex; $i++) {
        array_push($columnNames, strtolower($worksheet->getCellByColumnAndRow($i, 1)->getValue()));
    }

    // Validate there are no missing columns
    $missingValues = array();
    foreach ($expectedColumns as $expectedColumnName) {
        $idx = array_search($expectedColumnName, $columnNames);
        if (is_bool($idx)) {
            array_push($missingValues, $expectedColumnName);
        }
    }
    if (count($missingValues) > 0) {
        throw new Exception('NAWS export is missing required columns: ' . implode(', ', $missingValues));
    }

    // the PHP spreadsheet package starts with column 1 rather than 0, so add 1 to each of these values
    $directionsIndex = 1 + array_search('directions', $columnNames);
    $additionalDirectionsIndex = 1 + array_search('additionaldirections', $columnNames);
    $closedIndex = 1 + array_search('closed', $columnNames);
    $roomIndex = 1 + array_search('room', $columnNames);
    $formatIndices[0] = 1 + array_search('format1', $columnNames);
    $formatIndices[1] = 1 + array_search('format2', $columnNames);
    $formatIndices[2] = 1 + array_search('format3', $columnNames);
    $formatIndices[3] = 1 + array_search('format4', $columnNames);
    $formatIndices[4] = 1 + array_search('format5', $columnNames);

    // add names from formatWorldIdToName to the formatNameToWorldId array
    foreach ($formatWorldIdToName as $id => $name) {
        $formatNameToWorldId[$name] = $id;
    }

    // update columns as needed
    $iter = $spreadsheet->getActiveSheet()->getRowIterator();
    $n = 0;
    foreach ($iter as $row) {
        $n++;
        // skip the header row
        if ($n > 1) {
            // combine directions and additionalDirections, and write the result in the directions column
            $directions = $spreadsheet->getActiveSheet()->getCellByColumnAndRow($directionsIndex, $n)->getValue();
            $additionalDirections = $spreadsheet->getActiveSheet()->getCellByColumnAndRow($additionalDirectionsIndex, $n)->getValue();
            if ($directions) {
                if ($additionalDirections) {
                    $directions = $directions . '; ' . $additionalDirections;
                }
            } else {
                $directions = $additionalDirections;
            }
            $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($directionsIndex, $n, $directions);
            $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($additionalDirectionsIndex, $n, '');
            // the value for 'closed' needs to be a capitalized rather than a boolean
            $isClosed = $spreadsheet->getActiveSheet()->getCellByColumnAndRow($closedIndex, $n)->getValue();
            if (is_bool($isClosed)) {
                $isClosed = $isClosed ? 'CLOSED' : 'OPEN';
            } else {
                $isClosed = strtoupper($isClosed);
            }
            $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($closedIndex, $n, $isClosed);
            // the formats in the original spreadsheet are in the format1 column, separated by semicolons
            $formatNames = explode('; ', $spreadsheet->getActiveSheet()->getCellByColumnAndRow($formatIndices[0], $n)->getValue());
            // clear out format1 (in case there aren't any formats that map to a world ID)
            $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($formatIndices[0], $n, '');
            // if the format name isn't in the World ID table, or if there are more than 5 formats, append the format name to the Room column
            // (weird but that's what the standard seems to be for the old NAWS import format)
            $room = $spreadsheet->getActiveSheet()->getCellByColumnAndRow($roomIndex, $n)->getValue();
            $i = 0;
            foreach ($formatNames as $name) {
                if ($i < 5 && array_key_exists($formatNames[$i], $formatNameToWorldId)) {
                    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($formatIndices[$i], $n, $formatNameToWorldId[$formatNames[$i]]);
                    $i++;
                } else {
                    if ($room) {
                        $room = $room . '; ' . $name;
                    } else {
                        $room = $name;
                    }
                }
            }
            $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($roomIndex, $n, $room);
        }
    }
    $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
    $writer->save($outputFile);

}

$in = $argv[1];
$out = $argv[2];

echo "Updating NAWS import file\nInput file: $in \nOutput file: $out \n\n";
update_file($in, $out);

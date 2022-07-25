<?php

use Google\Service\Sheets;
use Suyash\GoogleSheetsCrud\SheetOperations;

require_once __DIR__ . '/./vendor/autoload.php';

/**
 * @return Sheets
 * @throws \Google\Exception
 */
function getService(): Sheets
{
    $client = new Google\Client();
    $client -> setApplicationName('Google Sheets API CRUD Testing');
    $client -> setScopes(Google_Service_Sheets::SPREADSHEETS);
    $client -> setAuthConfig('credentials.json');
    $client -> setAccessType('offline');
    return new Google\Service\Sheets($client);
}

try {
    $service = getService();
    $spreadsheetId = '1PJlIRk4PNwBZrhWTRA4VodT2-lSRLnHX27enkL5ghhE';
    $range = 'Sheet1';
    $values = [
        ['Column 1', 'Column 2'],
        ['Value 1', 'Value 2']
    ];

    $sheet1 = new SheetOperations($service, $spreadsheetId);
    echo $sheet1->insertInto($range, $values) . ' rows updated';
} catch (\Google\Exception $e) {
    echo $e -> getMessage();
}

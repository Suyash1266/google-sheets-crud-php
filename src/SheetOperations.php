<?php

namespace Suyash\GoogleSheetsCrud;

use Google\Service\Sheets;
use Google_Service_Sheets_ValueRange;

class SheetOperations
{
    /**
     * @var Sheets Google\Service\Sheets Object obtained after creating new Google client.
     */
    private Sheets $service;

    /**
     * @var string The ID of the Google Spreadsheet. Can be obtained from the URL of the concerned Google spreadsheet.
     */
    private string $spreadsheetId;

    /**
     * @param string $spreadsheetId The spreadsheet on which this object will work
     */
    public function __construct(Sheets $service, string $spreadsheetId)
    {
        $this->service = $service;
        $this->spreadsheetId = $spreadsheetId;
    }

    /**
     * Inserts single or multiple rows at a time without the regard of number of values provided in each row. Each cell in a row will be inserted sequentially.
     *
     * @param string $range Range in the Google Sheet to be considered. Name of the work sheet, e.g., Sheet1, means whole sheet will be considered and hence new rows will be inserted just after last entry in the sheet.
     * @param array $values 2D array of values to be inserted with each outer element corresponding to a row and each inner element corresponding to the value to be inserted.
     * @param string $insertAs How the input data should be interpreted. Available options are:
     * - RAW: The values the user has entered will not be parsed and will be stored as-is
     * - USER_ENTERED: The values will be parsed as if the user typed them into the UI. Numbers will stay as numbers, but strings may be converted to number, dates, etc. following the same rules that are applied when entering text into a cell via the Google Sheets UI.
     * Default is RAW.
     *
     * @return int Number of rows updated after the operation
     */
    public function insertInto(string $range, array $values, string $insertAs = 'RAW'): int
    {
        $body = new Google_Service_Sheets_ValueRange([
            'values' => $values
        ]);
        $params = [
            'valueInputOption' => $insertAs
        ];
        $result = $this->service->spreadsheets_values->append($this->spreadsheetId, $range, $body, $params);
        return $result->getUpdates()->getUpdatedRows();
    }
}
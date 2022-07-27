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
     * @param string $range The A1 notation or R1C1 notation of the range to retrieve values from. Formats: Sheet1!A1:B2 | Sheet1!R1C1:R2C2 | Sheet1 (for entire sheet)
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
            'valueInputOption' => $insertAs,
            'insertDataOption' => 'INSERT_ROWS'
        ];
        $result = $this->service->spreadsheets_values->append($this->spreadsheetId, $range, $body, $params);
        return $result->getUpdates()->getUpdatedRows();
    }

    /**
     * Returns the column names (first row) of the sheet along with the column index (in A1 notation).
     *
     * @param string $range The A1 notation or R1C1 notation of the range to retrieve values from. Formats: Sheet1!A1:B2 | Sheet1!R1C1:R2C2 | Sheet1 (for entire sheet)
     * @return array Associative array of column names and their index in A1 notation.
     */
    public function getColumns(string $range): array
    {
        $result = $this->service->spreadsheets_values->get($this->spreadsheetId, $range);
        $columns = $result->getValues()[0];
        $initialCount = count($columns);
        for ($i = 0; $i < $initialCount; $i++) {
            if ($columns[$i] !== '') {
                $column = chr(65 + $i);
                $columns[$columns[$i]] = $column;
            }
            unset($columns[$i]);
        }
        return $columns;
    }

    /**
     * Inserts values into the spreadsheet at specified columns.
     *
     * @param string $range - The A1 notation or R1C1 notation of the range to retrieve values from. Formats: Sheet1!A1:B2 | Sheet1!R1C1:R2C2 | Sheet1 (for entire sheet)
     * @param array $values - 2D array of values with first row being the column names.
     * @param string $insertAs - How the input data should be interpreted. Available options are:
     * - RAW: The values the user has entered will not be parsed and will be stored as-is
     * - USER_ENTERED: The values will be parsed as if the user typed them into the UI. Numbers will stay as numbers, but strings may be converted to number, dates, etc. following the same rules that are applied when entering text into a cell via the Google Sheets UI.
     * Default is RAW.
     *
     * @return int - Number of rows updated after the operation.
     */
    public function insertIntoColumns(string $range, array $values, string $insertAs = 'RAW'): int
    {
        $allColumns = $this->getColumns($range);
        $allColumnsCount = count($allColumns);
        $requiredColumns = $values[0];
        $requiredColumnsCount = count($requiredColumns);
        $requiredColumnsPositions = [];
        $rowTemplate = [];
        for ($i = 0; $i < $requiredColumnsCount; $i++) {
            $requiredColumnsPositions[$i] = ord($allColumns[$requiredColumns[$i]]) - 65;
        }
        for ($i = 0; $i < $allColumnsCount; $i++) {
            if (in_array($i, $requiredColumnsPositions)) {
                $rowTemplate[$i] = true;
            } else {
                $rowTemplate[$i] = false;
            }
        }
        $values = array_slice($values, 1);
        $valuesRestructured = [];
        for ($i = 0; $i < count($values); $i++) {
            $columnIndex = 0;
            for ($j = 0; $j < $allColumnsCount; $j++) {
                if ($rowTemplate[$j]) {
                    $valuesRestructured[$i][$j] = $values[$i][$columnIndex++];
                } else {
                    $valuesRestructured[$i][$j] = '';
                }
            }
        }
        return $this->insertInto($range, $valuesRestructured, $insertAs);
    }
}
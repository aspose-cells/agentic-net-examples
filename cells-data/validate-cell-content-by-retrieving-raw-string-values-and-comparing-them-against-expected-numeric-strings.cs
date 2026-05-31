using System;
using Aspose.Cells;

class ValidateCellContent
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate cells with various values
        cells["A1"].PutValue("123");          // numeric string
        cells["A2"].PutValue("45.67");        // numeric string with decimal
        cells["A3"].PutValue("ABC");          // non‑numeric string
        cells["A4"].PutValue(89);             // actual numeric value
        cells["A5"].PutValue("2021-06-20");   // date string

        // Expected numeric strings to validate against
        string[] expectedNumericStrings = { "123", "45.67", "89", "2021-06-20" };

        // Determine the used range of the worksheet
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        // Iterate through each used cell
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];

                // Retrieve the raw string representation of the cell's value
                string rawString = cell.StringValue;

                // Check if the raw string matches any of the expected numeric strings
                bool matchesExpected = Array.Exists(expectedNumericStrings, s => s == rawString);

                // Additionally, use IsNumericValue to see if the cell is considered numeric by Aspose.Cells
                bool isNumeric = cell.IsNumericValue;

                Console.WriteLine($"Cell {cell.Name}: Raw='{rawString}' | MatchesExpected={matchesExpected} | IsNumericValue={isNumeric}");
            }
        }

        // Save the workbook (optional)
        workbook.Save("ValidateCellContent.xlsx");
    }
}
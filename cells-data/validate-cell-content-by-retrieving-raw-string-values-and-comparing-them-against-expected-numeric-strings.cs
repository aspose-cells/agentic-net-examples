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

        // Populate cells with string values (some numeric, some not)
        cells["A1"].PutValue("123");
        cells["A2"].PutValue("45.67");
        cells["A3"].PutValue("NotANumber");
        cells["A4"].PutValue("00100");

        // Expected numeric strings to validate against
        string[] expectedNumeric = { "123", "45.67", "00100" };

        // Determine the used range of the worksheet
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        // Iterate through each populated cell
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];

                // Retrieve the raw string value from the cell
                string rawValue = cell.StringValue;

                // Check if the raw string matches any expected numeric string
                bool matchesExpected = Array.Exists(expectedNumeric, s => s.Equals(rawValue, StringComparison.Ordinal));

                // Additionally, use IsNumericValue to see if the cell is considered numeric by Aspose.Cells
                bool isNumeric = cell.IsNumericValue;

                Console.WriteLine($"Cell {cell.Name}: Raw=\"{rawValue}\", IsNumeric={isNumeric}, MatchesExpected={matchesExpected}");
            }
        }

        // Save the workbook (optional)
        workbook.Save("ValidateCellContent.xlsx");
    }
}
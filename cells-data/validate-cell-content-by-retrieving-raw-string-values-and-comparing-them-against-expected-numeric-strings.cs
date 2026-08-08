// Title: C# – Validate Excel cell strings against expected numeric values using Aspose.Cells
// Description: Creates a workbook, writes numeric, decimal, text and date strings to column A, reads each cell’s raw string via StringValue, compares it with a predefined list of expected numeric strings, outputs match results and the IsNumericValue flag, then saves the file.
// Keywords: Aspose.Cells C# raw string value | StringValue | IsNumericValue | validate Excel cell content | compare cell string to expected value | numeric string validation | date string handling | Excel data validation .NET
// Common Searches: Aspose.Cells get raw string from cell | C# compare cell value to expected string Aspose.Cells | How to check if cell is numeric Aspose.Cells | Validate Excel data against list of strings C# | IsNumericValue for date cells Aspose.Cells
// Developer Intent: Read each cell’s original string, compare it to a predefined numeric string array, and determine whether the cell is treated as numeric by Aspose.Cells.
// Use Cases: Validate user‑entered numeric strings in an uploaded Excel file before processing. | Ensure date strings are correctly recognized as numeric values during import. | Generate a validation report that lists mismatched cells together with their IsNumericValue status.
// AI Prompts: Write C# code with Aspose.Cells that iterates through column A, retrieves each cell’s raw string via StringValue, compares it to an array of expected numeric strings, and prints the match result and IsNumericValue flag. | Show how to log validation failures, including cell name, raw value, expected value, and IsNumericValue, for later review. | Explain Aspose.Cells’ logic for setting IsNumericValue on date strings and suggest handling strategies when validating mixed data types.

using System;
using Aspose.Cells;

// Creates a workbook, writes numeric, decimal, text and date strings to column A, reads each cell’s raw string via StringValue, compares it with a predefined list of expected numeric strings, outputs match results and the IsNumericValue flag, then saves the file.
class ValidateCellContent
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate cells with various string values
        cells["A1"].PutValue("123");          // numeric string
        cells["A2"].PutValue("45.67");        // numeric string with decimal
        cells["A3"].PutValue("ABC");          // non‑numeric string
        cells["A4"].PutValue("2021-06-20");   // date string (treated as numeric after conversion)

        // Expected numeric strings for validation
        string[] expectedValues = { "123", "45.67", "100", "2021-06-20" };

        // Iterate through the cells and compare raw string values with expected strings
        for (int row = 0; row < expectedValues.Length; row++)
        {
            Cell cell = cells[row, 0];               // Column A (index 0)
            string rawString = cell.StringValue;     // Retrieve raw string representation
            bool isMatch = rawString == expectedValues[row];

            Console.WriteLine($"Cell {cell.Name}: Raw='{rawString}' Expected='{expectedValues[row]}' Match={isMatch}");
            // Additional check: whether the cell content is considered numeric by Aspose.Cells
            Console.WriteLine($"  IsNumericValue = {cell.IsNumericValue}");
        }

        // Save the workbook (optional, demonstrates the required save lifecycle step)
        workbook.Save("ValidateCellContent.xlsx");
    }
}

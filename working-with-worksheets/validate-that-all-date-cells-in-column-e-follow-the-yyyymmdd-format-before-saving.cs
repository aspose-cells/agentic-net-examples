// Title: Validate yyyy‑MM‑dd date format in column E of an Excel workbook using Aspose.Cells for .NET before saving
// AI Prompts: Loop through column E with Aspose.Cells, read each cell's displayed string, apply DateTime.TryParseExact with the 'yyyy-MM-dd' pattern, and throw an InvalidOperationException if any cell fails the check. | Inspect every DateTime cell in column E of a loaded workbook, confirm its text matches the ISO date format, and only invoke Workbook.Save after all cells pass validation.
// Common Searches: c# aspocells verify date format in specific column before saving workbook | how to enforce ISO yyyy-mm-dd dates in Excel column using Aspose.Cells .NET | throw exception for invalid date string in Excel cell with Aspose.Cells | validate date cells in column E of an .xlsx file using Aspose.Cells C#
// Tags: Aspose.Cells column date format validation | C# DateTime.TryParseExact with Aspose.Cells | Excel column E date pattern enforcement | InvalidOperationException for mismatched Excel dates | Workbook.Save after date validation Aspose

using Aspose.Cells;
using System;
using System.Globalization;

// The program loads an Excel workbook, iterates over all rows in column E, validates that each DateTime cell's displayed text conforms to the 'yyyy-MM-dd' pattern using DateTime.TryParseExact, throws an InvalidOperationException on any mismatch, and saves the workbook only after successful validation.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Expected date format
        const string expectedFormat = "yyyy-MM-dd";

        // Determine the last used row in the sheet
        int lastRow = cells.MaxDataRow;

        // Iterate through all rows in column E (zero‑based index 4)
        for (int row = 0; row <= lastRow; row++)
        {
            Cell cell = cells[row, 4]; // Column E

            // Process only cells that contain a DateTime value
            if (cell.Type == CellValueType.IsDateTime)
            {
                // Retrieve the cell's displayed text
                string cellText = cell.StringValue;

                // Validate the text against the expected format
                if (!DateTime.TryParseExact(cellText, expectedFormat, CultureInfo.InvariantCulture,
                                            DateTimeStyles.None, out _))
                {
                    // Throw an exception if the format is incorrect
                    throw new InvalidOperationException(
                        $"Cell {cell.Name} does not follow the required format '{expectedFormat}'.");
                }
            }
        }

        // Save the workbook after successful validation
        workbook.Save("output.xlsx");
    }
}

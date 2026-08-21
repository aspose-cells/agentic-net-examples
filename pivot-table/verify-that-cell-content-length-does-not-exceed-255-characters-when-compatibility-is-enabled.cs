// Title: Check and warn if any cell exceeds 255 characters in Aspose.Cells (C#)
// Description: Creates a workbook, inserts a 300‑character string, then scans all worksheets to identify string cells longer than 255 characters. Each violation is logged as a console warning, allowing the file to be saved without compatibility errors.
// Keywords: Aspose.Cells | C# | cell length limit | 255 characters | Excel compatibility | validate cell content | detect long strings | warning log | workbook verification
// Common Searches: Aspose.Cells limit cell length 255 | C# verify Excel cell string length | log warning for oversized cell Aspose | check cell content length across worksheets | prevent compatibility errors Aspose.Cells
// Developer Intent: Detect cells whose text exceeds the 255‑character limit when Excel compatibility mode is enabled.
// Use Cases: Run a pre‑save scan that reports cells over 255 characters to avoid Excel compatibility failures. | Integrate the check into an ETL pipeline that imports data into Excel files. | Use the warning output to trigger automatic truncation or flagging of oversized cell values in reporting tools.
// AI Prompts: Generate a method that truncates cell strings to 255 characters instead of only logging a warning using Aspose.Cells. | Show how to enable Excel compatibility mode in Aspose.Cells before executing the VerifyCellContentLength check. | Create a unit test that confirms VerifyCellContentLength logs a warning for a cell containing more than 255 characters.

using System;
using Aspose.Cells;

// Creates a workbook, inserts a 300‑character string, then scans all worksheets to identify string cells longer than 255 characters. Each violation is logged as a console warning, allowing the file to be saved without compatibility errors.
class VerifyCellLength
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Example: put a string longer than 255 characters into a cell
            string longString = new string('A', 300);
            cells["A1"].PutValue(longString);

            // Verify that no cell exceeds 255 characters when compatibility is enabled
            VerifyCellContentLength(workbook, 255);

            // Save the workbook (using the standard save method)
            workbook.Save("VerifiedWorkbook.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Checks each cell in the workbook and logs a warning if a string exceeds the specified max length
    static void VerifyCellContentLength(Workbook workbook, int maxLength)
    {
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Type == CellValueType.IsString)
                    {
                        string value = cell.StringValue;
                        if (!string.IsNullOrEmpty(value) && value.Length > maxLength)
                        {
                            // Log a warning instead of throwing an exception
                            Console.WriteLine(
                                $"Warning: Cell {cell.Name} contains {value.Length} characters, exceeding the allowed limit of {maxLength}.");
                        }
                    }
                }
            }
        }
    }
}

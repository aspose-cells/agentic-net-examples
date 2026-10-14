// Title: Log a warning in C# with Aspose.Cells when a worksheet’s MaxDataRow is zero while MaxDataColumn is greater than zero
// AI Prompts: Generate C# code that uses Aspose.Cells to iterate all worksheets in a workbook and writes a console warning for any sheet where Cells.MaxDataRow equals 0 and Cells.MaxDataColumn is greater than 0. | Add a diagnostic message to an existing Aspose.Cells processing loop that includes the worksheet name and the MaxDataColumn value when the sheet contains columns but no rows of data. | Create a .NET method that checks each worksheet’s MaxDataRow and MaxDataColumn properties and logs a warning if the row count is zero while the column count is non‑zero.
// Common Searches: Aspose.Cells how to detect worksheets with data only in columns in C# | C# log warning when Excel sheet has MaxDataRow = 0 but MaxDataColumn > 0 using Aspose | check Cells.MaxDataRow and Cells.MaxDataColumn for empty rows in Aspose.Cells workbook | identify Excel worksheets that contain columns but no rows with Aspose.Cells API | Aspose.Cells warning for sheets with zero rows and non‑zero columns
// Tags: Aspose.Cells worksheet MaxDataRow validation | C# log warning for empty rows Aspose.Cells | detect columns without rows Aspose.Cells | Cells.MaxDataColumn greater than zero check | Excel sheet data range analysis with Aspose.Cells

using Aspose.Cells;
using System;

// C# program using Aspose.Cells loads a workbook, iterates each worksheet, reads Cells.MaxDataRow and Cells.MaxDataColumn, and writes a console warning when a sheet has no data rows but has one or more data columns.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Retrieve the maximum data row and column indices (zero‑based)
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Log a warning if MaxDataRow is zero but MaxDataColumn is greater than zero
            if (maxRow == 0 && maxCol > 0)
            {
                Console.WriteLine($"Warning: Worksheet '{sheet.Name}' has MaxDataRow = 0 but MaxDataColumn = {maxCol}.");
            }
        }

        // Save the workbook if any modifications were made (optional)
        // workbook.Save("output.xlsx");
    }
}

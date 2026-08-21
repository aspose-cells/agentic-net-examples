// Title: Log warning for worksheets where MaxDataRow = 0 and MaxDataColumn > 0 – Aspose.Cells for .NET (C#)
// Description: Creates or loads a Workbook, iterates each Worksheet, reads Cells.MaxDataRow and Cells.MaxDataColumn, and writes a console warning when the row index is zero while the column index is greater than zero, then saves the file.
// Keywords: Aspose.Cells | MaxDataRow | MaxDataColumn | C# | .NET | worksheet validation | empty rows | column‑only sheet | warning log | Excel data range | cells.MaxDataRow | cells.MaxDataColumn
// Common Searches: Aspose.Cells check MaxDataRow zero | C# detect worksheet with columns but no rows | log warning when MaxDataColumn > 0 and MaxDataRow = 0 | validate Excel sheet data extents Aspose.Cells | how to use Cells.MaxDataRow and MaxDataColumn
// Developer Intent: Identify worksheets that contain column data without any data rows and output a warning message.
// Use Cases: Pre‑process uploaded Excel files to confirm at least one data row exists. | Flag header‑only sheets before importing data into a database. | Prevent downstream errors in reporting pipelines caused by missing rows. | Automate quality checks in ETL workflows that consume Excel sources.
// AI Prompts: Write C# code using Aspose.Cells that prints a warning when a worksheet's MaxDataRow is 0 while MaxDataColumn is greater than 0. | Explain the behavior of Cells.MaxDataRow and Cells.MaxDataColumn when a sheet contains only column headers. | Suggest alternative handling (e.g., skip sheet, add placeholder row) for worksheets with columns but no rows in Aspose.Cells. | Generate unit tests for the warning logic in a .NET project.

using System;
using Aspose.Cells;

// Creates or loads a Workbook, iterates each Worksheet, reads Cells.MaxDataRow and Cells.MaxDataColumn, and writes a console warning when the row index is zero while the column index is greater than zero, then saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        // Example of loading: workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Retrieve the maximum data row and column indices
            int maxDataRow = cells.MaxDataRow;       // Zero‑based, -1 if no data
            int maxDataColumn = cells.MaxDataColumn; // Zero‑based, -1 if no data

            // Log a warning when MaxDataRow is zero but MaxDataColumn is greater than zero
            if (maxDataRow == 0 && maxDataColumn > 0)
            {
                Console.WriteLine($"Warning: Worksheet \"{sheet.Name}\" has MaxDataRow = 0 but MaxDataColumn = {maxDataColumn}.");
            }
        }

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}

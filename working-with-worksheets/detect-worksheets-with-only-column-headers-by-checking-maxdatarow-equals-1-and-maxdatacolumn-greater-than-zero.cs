// Title: Detect Worksheets Containing Only Headers with Aspose.Cells for .NET (C#)
// Description: C# code that opens an Excel workbook using Aspose.Cells, scans each worksheet, and leverages Cells.MaxDataRow and Cells.MaxDataColumn to flag sheets where MaxDataRow equals -1 (no data rows) while MaxDataColumn is greater than zero (at least one header column). The worksheet names are output to the console; the file can be saved unchanged.
// Keywords: Aspose.Cells | C# | .NET | detect header‑only worksheets | MaxDataRow | MaxDataColumn | Excel workbook analysis | identify empty data rows | worksheet header detection | Excel automation example
// Common Searches: Aspose.Cells find sheets with only headers | MaxDataRow -1 meaning in Aspose.Cells | how to detect header‑only worksheet in C# | check if Excel sheet has no data rows but has headers | C# code to list worksheets that contain only column titles
// Developer Intent: Locate worksheets that consist solely of column headers and lack any data rows.
// Use Cases: Skip header‑only sheets during bulk import to avoid processing empty tables. | Generate a report of worksheets that need data population before further analysis. | Validate workbook structure by flagging sheets that are missing data but contain header rows.
// AI Prompts: Write C# code with Aspose.Cells that deletes worksheets identified as header‑only. | Explain the relationship between MaxDataRow and MaxDataColumn and how they can be combined to detect empty data sections. | Modify the example to write the names of header‑only worksheets to a new summary sheet instead of the console.

using System;
using Aspose.Cells;

// C# code that opens an Excel workbook using Aspose.Cells, scans each worksheet, and leverages Cells.MaxDataRow and Cells.MaxDataColumn to flag sheets where MaxDataRow equals -1 (no data rows) while MaxDataColumn is greater than zero (at least one header column). The worksheet names are output to the console; the file can be saved unchanged.
class DetectHeaderOnlySheets
{
    static void Main()
    {
        // Load an existing workbook (adjust the file path as needed)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // MaxDataRow == -1 means there are no data rows (only possible headers)
            // MaxDataColumn > 0 means at least one column contains a value (the header)
            if (sheet.Cells.MaxDataRow == -1 && sheet.Cells.MaxDataColumn > 0)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\" contains only column headers.");
            }
        }

        // Save the workbook (optional, can be omitted if no changes are made)
        workbook.Save("output.xlsx");
    }
}
